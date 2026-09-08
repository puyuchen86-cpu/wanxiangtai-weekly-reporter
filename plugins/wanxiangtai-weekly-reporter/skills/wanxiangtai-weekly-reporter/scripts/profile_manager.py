from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


PROFILE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
FORBIDDEN_KEYS = {
    "password",
    "cookie",
    "cookies",
    "access_token",
    "refresh_token",
    "app_secret",
    "secret",
}


def default_config_dir() -> Path:
    return Path.home() / ".codex" / "wanxiangtai-weekly-reporter" / "profiles"


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    temporary.replace(path)


def walk_keys(value: Any) -> set[str]:
    keys: set[str] = set()
    if isinstance(value, dict):
        for key, child in value.items():
            keys.add(str(key).lower())
            keys.update(walk_keys(child))
    elif isinstance(value, list):
        for child in value:
            keys.update(walk_keys(child))
    return keys


def validate_profile(profile: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    profile_id = str(profile.get("profile_id", ""))
    if not PROFILE_ID.fullmatch(profile_id):
        errors.append("profile_id 只允许小写字母、数字和单连字符")
    if profile.get("schema_version") != 1:
        errors.append("schema_version 必须为 1")
    if not str(profile.get("brand_name", "")).strip():
        errors.append("brand_name 不能为空")

    spreadsheet = profile.get("spreadsheet")
    if not isinstance(spreadsheet, dict) or not str(spreadsheet.get("token", "")).strip():
        errors.append("spreadsheet.token 不能为空")

    plans = profile.get("plans")
    if not isinstance(plans, list) or not plans:
        errors.append("plans 必须是非空数组")
        plans = []

    seen_campaigns: set[str] = set()
    seen_sheet_ids: set[str] = set()
    seen_sheet_names: set[str] = set()
    for index, plan in enumerate(plans, start=1):
        if not isinstance(plan, dict):
            errors.append(f"plans[{index}] 不是对象")
            continue
        required = ("campaign_id", "plan_name", "sheet_id", "sheet_name")
        for field in required:
            if not str(plan.get(field, "")).strip():
                errors.append(f"plans[{index}].{field} 不能为空")
        campaign_id = str(plan.get("campaign_id", ""))
        sheet_id = str(plan.get("sheet_id", ""))
        sheet_name = str(plan.get("sheet_name", ""))
        if campaign_id in seen_campaigns:
            errors.append(f"campaign_id 重复：{campaign_id}")
        if sheet_id in seen_sheet_ids:
            errors.append(f"sheet_id 重复：{sheet_id}")
        if sheet_name in seen_sheet_names:
            errors.append(f"sheet_name 重复：{sheet_name}")
        seen_campaigns.add(campaign_id)
        seen_sheet_ids.add(sheet_id)
        seen_sheet_names.add(sheet_name)

    forbidden = sorted(walk_keys(profile).intersection(FORBIDDEN_KEYS))
    if forbidden:
        errors.append("配置包含禁止的凭据字段：" + ", ".join(forbidden))
    return errors


def resolve_profile_path(config_dir: Path, profile: str) -> Path:
    candidate = Path(profile)
    if candidate.exists():
        return candidate
    return config_dir / f"{profile}.json"


def cmd_init(args: argparse.Namespace) -> int:
    config_dir = Path(args.config_dir).expanduser().resolve()
    if not PROFILE_ID.fullmatch(args.profile_id):
        raise SystemExit("profile-id 只允许小写字母、数字和单连字符")
    mappings = read_json(Path(args.mappings_file))
    if not isinstance(mappings, list):
        raise SystemExit("mappings-file 必须是 JSON 数组")

    data_root = (
        Path(args.data_root).expanduser().resolve()
        if args.data_root
        else config_dir.parent / "data" / args.profile_id
    )
    profile = {
        "schema_version": 1,
        "profile_id": args.profile_id,
        "brand_name": args.brand_name,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "spreadsheet": {"token": args.spreadsheet_token, "url": args.spreadsheet_url or ""},
        "excluded_sheets": ["计划月数据"],
        "template": {
            "first_block_start": 3,
            "block_height": 27,
            "block_stride": 28,
            "material_start_column": "G",
            "default_material_capacity": args.material_capacity,
        },
        "plans": mappings,
        "data_root": str(data_root),
    }
    errors = validate_profile(profile)
    if errors:
        raise SystemExit("配置校验失败：\n- " + "\n- ".join(errors))

    target = config_dir / f"{args.profile_id}.json"
    if target.exists() and not args.force:
        raise SystemExit(f"配置已存在，拒绝覆盖：{target}")
    write_json_atomic(target, profile)
    print(json.dumps({"ok": True, "profile": str(target)}, ensure_ascii=False))
    return 0


def cmd_validate(args: argparse.Namespace) -> int:
    config_dir = Path(args.config_dir).expanduser().resolve()
    path = resolve_profile_path(config_dir, args.profile)
    profile = read_json(path)
    errors = validate_profile(profile)
    result = {"ok": not errors, "profile": str(path), "errors": errors}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not errors else 1


def cmd_show(args: argparse.Namespace) -> int:
    config_dir = Path(args.config_dir).expanduser().resolve()
    path = resolve_profile_path(config_dir, args.profile)
    profile = read_json(path)
    errors = validate_profile(profile)
    if errors:
        raise SystemExit("配置校验失败：\n- " + "\n- ".join(errors))
    print(json.dumps(profile, ensure_ascii=False, indent=2))
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    config_dir = Path(args.config_dir).expanduser().resolve()
    rows: list[dict[str, Any]] = []
    if config_dir.exists():
        for path in sorted(config_dir.glob("*.json")):
            try:
                profile = read_json(path)
                errors = validate_profile(profile)
                rows.append(
                    {
                        "profile_id": profile.get("profile_id", path.stem),
                        "brand_name": profile.get("brand_name", ""),
                        "plans": len(profile.get("plans", [])),
                        "valid": not errors,
                    }
                )
            except Exception as exc:  # pragma: no cover - diagnostic path
                rows.append({"profile_id": path.stem, "valid": False, "error": str(exc)})
    print(json.dumps(rows, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage local Wanxiangtai brand profiles")
    parser.add_argument("--config-dir", default=str(default_config_dir()))
    subparsers = parser.add_subparsers(dest="command", required=True)

    init = subparsers.add_parser("init")
    init.add_argument("--profile-id", required=True)
    init.add_argument("--brand-name", required=True)
    init.add_argument("--spreadsheet-token", required=True)
    init.add_argument("--spreadsheet-url")
    init.add_argument("--mappings-file", required=True)
    init.add_argument("--data-root")
    init.add_argument("--material-capacity", type=int, default=24)
    init.add_argument("--force", action="store_true")
    init.set_defaults(handler=cmd_init)

    validate = subparsers.add_parser("validate")
    validate.add_argument("--profile", required=True)
    validate.set_defaults(handler=cmd_validate)

    show = subparsers.add_parser("show")
    show.add_argument("--profile", required=True)
    show.set_defaults(handler=cmd_show)

    listing = subparsers.add_parser("list")
    listing.set_defaults(handler=cmd_list)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return int(args.handler(args))


if __name__ == "__main__":
    sys.exit(main())
