# -*- coding: utf-8 -*-
"""Run reservation confirmation assist workflow."""

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
from court_reserv.config import get_debug_output_dir, get_output_base_path, load_config
from court_reserv.manage_id import Manage_Id
from court_reserv.services import (
    IdManagerService,
    LotteryResultWorkflowService,
    LotteryService,
    ReservationConfirmationWorkflowService,
    ReservationService,
)


def _show_info(title, message):
    print(f"[{title}] {message}")


def _ask_yes_no(title, message):
    print(f"[{title}] {message}")
    answer = input("Continue? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def _output_id_dict(id_dict, output_file_path):
    return output_file_path


def _select_entries(won_entries):
    if not won_entries:
        return []
    raw = input(
        "Select won entries by number (comma separated). "
        "Note: confirmation runs per account, so select all won entries for an account. "
    ).strip()
    if not raw:
        return []
    indices = []
    for part in raw.split(","):
        try:
            number = int(part.strip())
        except Exception:
            continue
        if 1 <= number <= len(won_entries):
            indices.append(number - 1)
    return sorted(set(indices))


def _confirm_result(result):
    selected_accounts = result.get("selected_accounts", [])
    if not selected_accounts:
        return ""
    print("Reservation confirmation will run for these accounts:")
    for account in selected_accounts:
        masked = result_workflow_service.lottery_result_workflow_service.mask_user_id(
            account["user_id"]
        )
        label = account.get("account_label", "")
        print(f"- {masked} {label}".strip())
    return input("Confirm reservation? Type 'yes' to continue: ").strip()


def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Show won lottery entries, let the user choose confirmation targets, "
            "and confirm reservations only after explicit 'yes'."
        )
    )
    parser.add_argument(
        "--id-csv",
        help="Optional ID CSV path. When provided, credentials are resolved from this CSV first.",
    )
    parser.add_argument(
        "--account-id",
        help="Optional account ID to use from the ID CSV. Defaults to all CSV entries.",
    )
    parser.add_argument(
        "--output-dir",
        help="Optional output directory for workflow result JSON.",
    )
    return parser


def main():
    global result_workflow_service

    parser = build_parser()
    args = parser.parse_args()

    config = load_config()
    logger = logging.getLogger("reservation_confirmation_workflow")

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
    reservation_service = ReservationService(
        config=config,
        browser_session=browser_session,
        navigation_service=navigation_service,
        logger=logger,
        get_id_dict_from_csv=Manage_Id.get_id_dict_from_csv,
        output_id_dict=Manage_Id.output_csv_from_id_dict,
        sleep_func=time.sleep,
    )
    lottery_result_workflow_service = LotteryResultWorkflowService(
        config=config,
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        lottery_service=lottery_service,
        id_manager_service=IdManagerService(config=config, sleep_func=time.sleep),
        logger=logger,
    )
    result_workflow_service = ReservationConfirmationWorkflowService(
        lottery_result_workflow_service=lottery_result_workflow_service,
        reservation_service=reservation_service,
        login_service=login_service,
        logger=logger,
    )

    result = result_workflow_service.run(
        id_csv=args.id_csv,
        account_id=args.account_id,
        select_entries_callback=_select_entries,
        confirm_callback=_confirm_result,
    )
    result_workflow_service.print_result(result)

    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else get_output_base_path() / "lottery_automation"
    )
    json_path = result_workflow_service.save_result(result, output_dir)
    print(f"Saved result JSON: {json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
