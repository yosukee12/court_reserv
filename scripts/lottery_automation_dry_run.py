# -*- coding: utf-8 -*-
"""Dry-run entrypoint for Phase 2 lottery automation."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from court_reserv.config import (
    get_output_base_path,
    load_config,
    load_reservation_preference,
)
from court_reserv.services import (
    LotteryAutomationDryRunService,
    SlotCollectionAdapter,
    SlotRankingService,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Rank lottery application candidates from preferences without submitting lottery entries."
    )
    parser.add_argument(
        "--preferences",
        required=True,
        help="Path to a YAML or JSON preference file.",
    )
    parser.add_argument(
        "--source-csv",
        help="Optional path to an available_slots_*.csv file used as a candidate source.",
    )
    parser.add_argument(
        "--output-dir",
        help="Optional output directory for JSON and CSV dry-run results.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Explicit no-op flag kept for clarity. Dry-run is always enforced in this issue.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    load_config()
    preference = load_reservation_preference(args.preferences)
    slot_adapter = SlotCollectionAdapter()
    ranking_service = SlotRankingService()
    dry_run_service = LotteryAutomationDryRunService(
        slot_adapter=slot_adapter,
        slot_ranking_service=ranking_service,
    )

    output_base_path = (
        Path(args.output_dir)
        if args.output_dir
        else get_output_base_path() / "lottery_automation"
    )
    search_dirs = [
        get_output_base_path(),
        get_output_base_path() / "debug_pages",
        Path("court_reserv/debug_pages"),
    ]

    result = dry_run_service.run(
        preference=preference,
        source_csv=args.source_csv,
        search_dirs=search_dirs,
    )
    dry_run_service.print_result(result)
    json_path, csv_path = dry_run_service.save_result(result, output_base_path)
    print(f"Saved dry-run JSON: {json_path}")
    print(f"Saved dry-run CSV: {csv_path}")
    print("Dry-run mode only. Lottery submission is intentionally disabled in this issue.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
