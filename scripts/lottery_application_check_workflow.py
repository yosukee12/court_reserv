# -*- coding: utf-8 -*-
"""Run lottery application status workflow and save a timestamped CSV."""

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
from court_reserv.config import get_debug_output_dir, load_config
from court_reserv.services import (
    IdManagerService,
    LotteryApplicationCheckWorkflowService,
)


def _show_info(title, message):
    print(f"[{title}] {message}")


def _ask_yes_no(title, message):
    print(f"[{title}] {message}")
    answer = input("Continue? [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch lottery application status rows and save a CSV with tennis entries only."
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
        help="Optional output directory for the timestamped CSV file.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    config = load_config()
    logger = logging.getLogger("lottery_application_check_workflow")

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
    workflow_service = LotteryApplicationCheckWorkflowService(
        config=config,
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        id_manager_service=IdManagerService(config=config, sleep_func=time.sleep),
        logger=logger,
        sleep_func=time.sleep,
    )

    result = workflow_service.run(
        id_csv=args.id_csv,
        account_id=args.account_id,
    )
    workflow_service.print_result(result)

    output_dir = workflow_service.resolve_output_dir(
        id_csv=args.id_csv,
        output_dir=args.output_dir,
    )
    csv_path = workflow_service.save_result(
        result,
        output_dir=output_dir,
        id_csv=args.id_csv,
    )
    print(f"Saved result CSV: {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
