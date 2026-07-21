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
    get_debug_output_dir,
    get_output_base_path,
    load_config,
    load_reservation_preference,
)
from court_reserv.services import (
    IdManagerService,
    LotteryEntrySlotCollector,
    LotteryEntryWorkflowService,
    LotteryService,
)


def _show_info(title, message):
    print(f"[{title}] {message}")
    if title == "手動最終送信":
        try:
            input("Press Enter after manually clicking the submit button: ")
        except EOFError:
            pass


def _ask_yes_no(title, message):
    print(f"[{title}] {message}")
    answer = input("Continue? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def _output_id_dict(id_dict, output_file_path):
    return output_file_path


def _build_entry_confirm_message(account_result, entry_result):
    slot = (entry_result or {}).get("slot", {})
    account_part = account_result.get("masked_user_id", "")
    if account_result.get("account_label"):
        account_part = f"{account_part} ({account_result.get('account_label')})"
    lines = [
        "今回申し込む枠:",
        f"ID / アカウント: {account_part}",
        f"申込み: {entry_result.get('apply_label', '')}",
        f"日付: {slot.get('date', '')}",
        f"曜日: {slot.get('weekday', '')}",
        f"時間帯: {slot.get('time_range', '')}",
        f"施設名: {slot.get('facility', '')}",
        f"現在申込数: {slot.get('current_entry_count', '')}",
    ]
    return "\n".join(lines)


def _confirm_submission(account_result, entry_result):
    slot = (entry_result or {}).get("slot", {})
    if not slot:
        return ""
    print(_build_entry_confirm_message(account_result, entry_result))
    print("送信する場合は yes を入力してください。")
    try:
        return input("送信しますか？ [yes/no]: ").strip()
    except EOFError:
        return ""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Collect lottery entry slots from the current page, apply default_entries/account_overrides, and submit only after explicit 'yes' confirmation."
        )
    )
    parser.add_argument(
        "--preferences",
        default="config/preferences.example.yaml",
        help="Path to a YAML or JSON preference file. Defaults to config/preferences.example.yaml.",
    )
    parser.add_argument(
        "--id-csv",
        help="Optional ID CSV path. When provided, credentials are resolved from this CSV first.",
    )
    parser.add_argument(
        "--account-id",
        help="Optional account ID to use from the ID CSV. Defaults to all resolved accounts.",
    )
    parser.add_argument(
        "--output-dir",
        help="Optional output directory for the workflow summary JSON.",
    )
    parser.add_argument(
        "--max-select",
        type=int,
        default=2,
        help="Maximum number of planned entries per account. Capped at 2.",
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
        logger=logger,
        get_debug_output_dir=get_debug_output_dir,
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
        slot_collector=LotteryEntrySlotCollector(
            navigation_service=navigation_service,
            browser_session=browser_session,
            logger=logger,
        ),
        logger=logger,
    )

    def preview_account_result(account_result):
        account_part = account_result.get("masked_user_id")
        if account_result.get("account_label"):
            account_part = f"{account_part} ({account_result.get('account_label')})"
        print(f"Account: {account_part}")
        print(f"Status: {account_result.get('status')}")
        if account_result.get("error"):
            print(f"Error: {account_result.get('error')}")
        print(
            f"Target weekdays: {', '.join(account_result.get('target_weekdays', []))}"
        )
        print(f"Collected slots: {len(account_result.get('collected_slots', []))}")
        planned_slots = account_result.get("planned_slots", [])
        if planned_slots:
            print("Planned entry slots:")
            for index, slot in enumerate(planned_slots, start=1):
                facility = " ".join(
                    part
                    for part in (
                        slot.get("park_name", ""),
                        slot.get("facility_name", ""),
                    )
                    if part
                ).strip()
                print(
                    f"{index}. {slot.get('date')} {slot.get('weekday')} {slot.get('time_range')} "
                    f"facility={facility} applied={slot.get('applied_count')}"
                )
        else:
            print("No planned entry slots.")

        missing_slots = account_result.get("missing_slots", [])
        if missing_slots:
            print("Warnings:")
            for warning in missing_slots:
                print(
                    f"- {warning.get('date')} {warning.get('time_range')} "
                    f"facility={warning.get('facility')} warning={warning.get('warning')}"
                )

    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else get_output_base_path() / "lottery_automation"
    )

    result = workflow_service.run(
        preference=preference,
        id_csv=args.id_csv,
        account_id=args.account_id,
        max_select=args.max_select,
        display_result_callback=preview_account_result,
        confirm_submit_callback=(
            (lambda account_result, entry_result: "dry-run")
            if preference.lottery_dry_run
            else _confirm_submission
        ),
        output_dir=output_dir,
    )
    workflow_service.print_result(result)

    result_path = workflow_service.save_result(result, output_dir)
    print(f"Saved workflow summary: {result_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
