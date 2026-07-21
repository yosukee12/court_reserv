# -*- coding: utf-8 -*-
"""Run lottery result workflow without reservation confirmation."""

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
from court_reserv.services import (
    IdManagerService,
    LotteryResultWorkflowService,
    LotteryService,
)


def _show_info(title, message):
    print(f"[{title}] {message}")


def _ask_yes_no(title, message):
    print(f"[{title}] {message}")
    answer = input("Continue? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def _output_id_dict(id_dict, output_file_path):
    return output_file_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch and classify lottery result rows without reservation confirmation."
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
        help="Optional output directory for workflow result JSON and CSV.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    config = load_config()
    logger = logging.getLogger("lottery_result_workflow")

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
    workflow_service = LotteryResultWorkflowService(
        config=config,
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        lottery_service=lottery_service,
        id_manager_service=IdManagerService(config=config, sleep_func=time.sleep),
        logger=logger,
    )

    result = workflow_service.run(
        id_csv=args.id_csv,
        account_id=args.account_id,
    )
    workflow_service.print_result(result)

    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else get_output_base_path() / "lottery_automation"
    )
    json_path, csv_path = workflow_service.save_result(result, output_dir)
    print(f"Saved result JSON: {json_path}")
    print(f"Saved result CSV: {csv_path}")
    print("Reservation confirmation is intentionally disabled in this workflow.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
