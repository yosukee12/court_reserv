# -*- coding: utf-8 -*-
"""Compatibility wrapper for the renamed lottery automation dry-run script."""

from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.lottery_automation_dry_run import main


if __name__ == "__main__":
    raise SystemExit(main())
