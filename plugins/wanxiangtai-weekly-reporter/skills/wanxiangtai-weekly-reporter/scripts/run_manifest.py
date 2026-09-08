from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from profile_manager import default_config_dir, read_json, resolve_profile_path, validate_profile, write_json_atomic


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


FINAL_PLAN_STATUS = "completed_and_verified"
FINAL_RUN_STATUS = "complete_and_globally_verified"


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("日期必须是 YYYY-MM-DD") from exc


def period_label(start: date, end: date) -> str:
    return f"{start.month}.{start.day}-{end.month}.{end.day}"


def manifest_path(profile: dict[str, Any], start: date, end: date) -> Path:
    root = Path(profile["data_root"]).expanduser().resolve()
    return root / f"{start.isoformat()}_{end.isoformat()}" / "task-list.json"


def load_profile(args: argparse.Namespace) -> tuple[Path, dict[str, Any]]:
    config_dir = Path(args.config_dir).expanduser().resolve()
    path = resolve_profile_path(config_dir, args.profile)
    profile = read_json(path)
    errors = validate_profile(profile)
    if errors:
        raise SystemExit("品牌配置无效：\n- " + "\n- ".join(errors))
    return path, profile


def cmd_start(args: argparse.Namespace) -> int:
    profile_path, profile = load_profile(args)
    if args.start > args.end:
        raise SystemExit("开始日期不能晚于结束日期")
    target = manifest_path(profile, args.start, args.end)
    if target.exists():
        existing = read_json(target)
        if existing.get("status") == FINAL_RUN_STATUS:
            raise SystemExit(f"该周期已完成，拒绝重复运行：{target}")
        if not args.resume:
            raise SystemExit(f"该周期存在未完成任务；使用 --resume 继续：{target}")
        print(json.dumps({"ok": True, "resumed": True, "manifest": str(target)}, ensure_ascii=False))
        return 0

    plans = []
    for order, plan in enumerate(profile["plans"], start=1):
        plans.append(
            {
                "order": order,
                "campaign_id": plan["campaign_id"],
                "plan_name": plan["plan_name"],
                "sheet_id": plan["sheet_id"],
                "sheet_name": plan["sheet_name"],
                "status": "pending",
            }
        )
    payload = {
        "schema_version": 1,
        "profile_id": profile["profile_id"],
        "profile_path": str(profile_path),
        "period": {
            "start": args.start.isoformat(),
            "end": args.end.isoformat(),
            "label": period_label(args.start, args.end),
        },
        "spreadsheet_token": profile["spreadsheet"]["token"],
        "excluded_sheets": profile.get("excluded_sheets", ["计划月数据"]),
        "status": "in_progress",
        "created_at": datetime.now(timezone.utc).isoformat(),
        "plans": plans,
    }
    write_json_atomic(target, payload)
    print(json.dumps({"ok": True, "created": True, "manifest": str(target)}, ensure_ascii=False))
    return 0


def cmd_mark(args: argparse.Namespace) -> int:
    path = Path(args.manifest).expanduser().resolve()
    payload = read_json(path)
    matched = [p for p in payload.get("plans", []) if str(p.get("campaign_id")) == args.campaign_id]
    if len(matched) != 1:
        raise SystemExit(f"campaign ID 映射数量必须为 1，实际为 {len(matched)}")
    plan = matched[0]
    plan["status"] = args.status
    if args.raw_json:
        plan["raw_json"] = args.raw_json
    if args.enrichment_json:
        plan["enrichment_json"] = args.enrichment_json
    if args.images is not None or args.videos is not None:
        plan["counts"] = {"images": args.images or 0, "videos": args.videos or 0}
    plan["updated_at"] = datetime.now(timezone.utc).isoformat()
    completed = sum(1 for item in payload["plans"] if item.get("status") == FINAL_PLAN_STATUS)
    payload["status"] = f"in_progress_{completed}_of_{len(payload['plans'])}_verified"
    write_json_atomic(path, payload)
    print(json.dumps({"ok": True, "completed": completed, "total": len(payload["plans"])}, ensure_ascii=False))
    return 0


def cmd_complete(args: argparse.Namespace) -> int:
    path = Path(args.manifest).expanduser().resolve()
    payload = read_json(path)
    unfinished = [p["sheet_name"] for p in payload.get("plans", []) if p.get("status") != FINAL_PLAN_STATUS]
    if unfinished:
        raise SystemExit("仍有未验证计划，不能完成：" + "、".join(unfinished))
    payload["status"] = FINAL_RUN_STATUS
    payload["verified_at"] = datetime.now(timezone.utc).isoformat()
    write_json_atomic(path, payload)
    print(json.dumps({"ok": True, "status": FINAL_RUN_STATUS, "plans": len(payload["plans"])}, ensure_ascii=False))
    return 0


def cmd_summary(args: argparse.Namespace) -> int:
    path = Path(args.manifest).expanduser().resolve()
    payload = read_json(path)
    counts: dict[str, int] = {}
    materials = 0
    for plan in payload.get("plans", []):
        status = str(plan.get("status", "unknown"))
        counts[status] = counts.get(status, 0) + 1
        materials += int(plan.get("counts", {}).get("images", 0))
        materials += int(plan.get("counts", {}).get("videos", 0))
    print(
        json.dumps(
            {
                "status": payload.get("status"),
                "period": payload.get("period"),
                "plans": len(payload.get("plans", [])),
                "status_counts": counts,
                "materials": materials,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage append-only weekly run manifests")
    parser.add_argument("--config-dir", default=str(default_config_dir()))
    subparsers = parser.add_subparsers(dest="command", required=True)

    start = subparsers.add_parser("start")
    start.add_argument("--profile", required=True)
    start.add_argument("--start", required=True, type=parse_date)
    start.add_argument("--end", required=True, type=parse_date)
    start.add_argument("--resume", action="store_true")
    start.set_defaults(handler=cmd_start)

    mark = subparsers.add_parser("mark")
    mark.add_argument("--manifest", required=True)
    mark.add_argument("--campaign-id", required=True)
    mark.add_argument("--status", default=FINAL_PLAN_STATUS)
    mark.add_argument("--raw-json")
    mark.add_argument("--enrichment-json")
    mark.add_argument("--images", type=int)
    mark.add_argument("--videos", type=int)
    mark.set_defaults(handler=cmd_mark)

    complete = subparsers.add_parser("complete")
    complete.add_argument("--manifest", required=True)
    complete.set_defaults(handler=cmd_complete)

    summary = subparsers.add_parser("summary")
    summary.add_argument("--manifest", required=True)
    summary.set_defaults(handler=cmd_summary)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return int(args.handler(args))


if __name__ == "__main__":
    sys.exit(main())
