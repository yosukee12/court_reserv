# -*- coding: utf-8 -*-
"""Run lottery entry candidate selection with manual submission confirmation."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
import sys
import time

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from court_reserv.browser import BrowserSession, LoginService, NavigationService
from court_reserv.config import (
    get_output_base_path,
    load_config,
    load_reservation_preference,
)
from court_reserv.services import (
    IdManagerService,
    LotteryEntryWorkflowService,
    LotteryService,
    SlotCollectionAdapter,
    SlotRankingService,
)


def _show_info(title, message):
    print(f"[{title}] {message}")


def _ask_yes_no(title, message):
    print(f"[{title}] {message}")
    answer = input("Continue? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def _output_id_dict(id_dict, output_file_path):
    return output_file_path


def _confirm_submission(result):
    selected_count = sum(
        1 for selected in result.get("selection_result", {}).values() if selected
    )
    if selected_count <= 0:
        return ""
    print(
        "Submit selected lottery entries? Type 'yes' to submit. Any other input will cancel."
    )
    return input("Submit? [yes/no]: ").strip()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Select ranked lottery entry candidates on the lottery page and submit only after explicit 'yes' confirmation."
        )
    )
    parser.add_argument(
        "--preferences",
        default="config/preferences.example.yaml",
        help="Path to a YAML or JSON preference file. Defaults to config/preferences.example.yaml.",
    )
    parser.add_argument(
        "--source-csv",
        help="Optional path to an available_slots_*.csv file used as a candidate source.",
    )
    parser.add_argument(
        "--id-csv",
        help="Optional ID CSV path. When provided, credentials are resolved from this CSV first.",
    )
    parser.add_argument(
        "--account-id",
        help="Optional account ID to use from the ID CSV. Defaults to the first CSV entry.",
    )
    parser.add_argument(
        "--output-dir",
        help="Optional output directory for the workflow summary JSON.",
    )
    parser.add_argument(
        "--max-select",
        type=int,
        default=2,
        help="Maximum number of ranked candidates to select on the page. Capped at 2.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    config = load_config()
    preference = load_reservation_preference(args.preferences)
    logger = logging.getLogger("lottery_entry_workflow")

    browser_session = BrowserSession(config)
    login_service = LoginService(
        top_url=config["URL"]["TOP_URL"],
        wait_factory=browser_session.get_wait,
        logger=logger,
        show_info=_show_info,
        ask_yes_no=_ask_yes_no,
        sleep_func=time.sleep,
    )
    navigation_service = NavigationService(
        wait_factory=browser_session.get_wait,
        sleep_func=time.sleep,
    )
    lottery_service = LotteryService(
        config=config,
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        logger=logger,
        show_info=_show_info,
        output_id_dict=_output_id_dict,
        sleep_func=time.sleep,
    )
    workflow_service = LotteryEntryWorkflowService(
        config=config,
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        lottery_service=lottery_service,
        id_manager_service=IdManagerService(config=config, sleep_func=time.sleep),
        slot_adapter=SlotCollectionAdapter(),
        slot_ranking_service=SlotRankingService(),
        logger=logger,
    )

    def preview_result(result):
        source_csv = result.get("source_csv")
        if source_csv:
            print(f"Candidate source: {source_csv}")
        print(f"Credential source: {result.get('credential_source')}")
        print(f"Target user ID: {result.get('user_id')}")
        if result.get("account_label"):
            print(f"Account label: {result.get('account_label')}")
        print(f"Total candidates: {result.get('total_slots', 0)}")

        selected_candidates = result.get("selected_candidates", [])
        if not selected_candidates:
            print("No lottery entry candidates were selected.")
            return

        print("Selected lottery entry candidates:")
        for index, ranked in enumerate(selected_candidates, start=1):
            reasons = ", ".join(ranked.reasons) if ranked.reasons else "no preference match"
            print(
                f"{index}. score={ranked.score} slot={workflow_service._to_slot_text(ranked)} reasons={reasons}"
            )

        selection_result = result.get("selection_result", {})
        if selection_result:
            print("Lottery page selection result:")
            for slot_text, selected in selection_result.items():
                print(f"- {slot_text}: {'selected' if selected else 'not selected'}")

    search_dirs = [
        get_output_base_path(),
        get_output_base_path() / "debug_pages",
        Path("court_reserv/debug_pages"),
    ]
    result = workflow_service.run(
        preference=preference,
        source_csv=args.source_csv,
        search_dirs=search_dirs,
        id_csv=args.id_csv,
        account_id=args.account_id,
        max_select=args.max_select,
        display_result_callback=preview_result,
        confirm_submit_callback=_confirm_submission,
    )
    workflow_service.print_result(result)

    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else get_output_base_path() / "lottery_automation"
    )
    result_path = workflow_service.save_result(result, output_dir)
    print(f"Saved workflow summary: {result_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
