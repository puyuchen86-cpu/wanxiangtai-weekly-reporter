from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[1]
SCRIPTS = (
    REPO
    / "plugins"
    / "wanxiangtai-weekly-reporter"
    / "skills"
    / "wanxiangtai-weekly-reporter"
    / "scripts"
)
PROFILE = SCRIPTS / "profile_manager.py"
MANIFEST = SCRIPTS / "run_manifest.py"


def run(*args: str, expected: int = 0) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        [sys.executable, *args],
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=False,
    )
    if result.returncode != expected:
        raise AssertionError(
            f"returncode={result.returncode}, expected={expected}\nstdout={result.stdout}\nstderr={result.stderr}"
        )
    return result


class StateToolsTest(unittest.TestCase):
    def test_profile_manifest_lifecycle_and_duplicate_guard(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = root / "profiles"
            data = root / "data"
            mappings = root / "mappings.json"
            mappings.write_text(
                json.dumps(
                    [
                        {
                            "campaign_id": "1001",
                            "plan_name": "人群推广_计划A",
                            "sheet_id": "sheet-a",
                            "sheet_name": "计划A",
                        },
                        {
                            "campaign_id": "1002",
                            "plan_name": "人群推广_计划B",
                            "sheet_id": "sheet-b",
                            "sheet_name": "计划B",
                        },
                    ],
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            run(
                str(PROFILE),
                "--config-dir",
                str(config),
                "init",
                "--profile-id",
                "brand-a",
                "--brand-name",
                "品牌A",
                "--spreadsheet-token",
                "spreadsheet-token",
                "--mappings-file",
                str(mappings),
                "--data-root",
                str(data),
            )
            validated = run(
                str(PROFILE), "--config-dir", str(config), "validate", "--profile", "brand-a"
            )
            self.assertTrue(json.loads(validated.stdout)["ok"])

            started = run(
                str(MANIFEST),
                "--config-dir",
                str(config),
                "start",
                "--profile",
                "brand-a",
                "--start",
                "2026-09-07",
                "--end",
                "2026-09-13",
            )
            manifest = Path(json.loads(started.stdout)["manifest"])
            for campaign in ("1001", "1002"):
                run(
                    str(MANIFEST),
                    "mark",
                    "--manifest",
                    str(manifest),
                    "--campaign-id",
                    campaign,
                    "--images",
                    "24",
                    "--videos",
                    "3",
                )
            completed = run(str(MANIFEST), "complete", "--manifest", str(manifest))
            self.assertEqual(json.loads(completed.stdout)["status"], "complete_and_globally_verified")

            duplicate = subprocess.run(
                [
                    sys.executable,
                    str(MANIFEST),
                    "--config-dir",
                    str(config),
                    "start",
                    "--profile",
                    "brand-a",
                    "--start",
                    "2026-09-07",
                    "--end",
                    "2026-09-13",
                ],
                text=True,
                encoding="utf-8",
                capture_output=True,
                check=False,
            )
            self.assertNotEqual(duplicate.returncode, 0)
            self.assertIn("拒绝重复运行", duplicate.stderr)

    def test_profile_rejects_credential_fields(self) -> None:
        sys.path.insert(0, str(SCRIPTS))
        try:
            from profile_manager import validate_profile

            profile = {
                "schema_version": 1,
                "profile_id": "unsafe",
                "brand_name": "品牌",
                "spreadsheet": {"token": "sheet", "access_token": "must-not-store"},
                "plans": [
                    {
                        "campaign_id": "1",
                        "plan_name": "计划",
                        "sheet_id": "s1",
                        "sheet_name": "计划",
                    }
                ],
            }
            errors = validate_profile(profile)
            self.assertTrue(any("禁止的凭据字段" in error for error in errors))
        finally:
            sys.path.remove(str(SCRIPTS))


if __name__ == "__main__":
    unittest.main()
