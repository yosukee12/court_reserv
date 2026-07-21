# -*- coding: utf-8 -*-

import csv
from configparser import ConfigParser
from pathlib import Path

from court_reserv.services.lottery_result_workflow import LotteryResultWorkflowService
from court_reserv.services.reservation_confirmation_workflow import (
    ReservationConfirmationWorkflowService,
)


class _DummyIdManagerService:
    def __init__(self, accounts):
        self.accounts = accounts

    def load_accounts(self, csv_file_path):
        return self.accounts

    def save_accounts(self, id_dict, output_file_path):
        with open(output_file_path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            for user_id, values in id_dict.items():
                writer.writerow([user_id, *values])


def _service(accounts=None):
    return LotteryResultWorkflowService(
        config=ConfigParser(),
        browser_session=None,
        login_service=None,
        navigation_service=None,
        lottery_service=None,
        id_manager_service=_DummyIdManagerService(accounts or {}),
    )


def test_save_result_writes_csv_next_to_input_csv(tmp_path):
    input_csv = tmp_path / "input" / "ids.csv"
    input_csv.parent.mkdir()
    service = _service()

    json_path, csv_path = service.save_result(
        {"results": []},
        output_dir=tmp_path / "output",
        id_csv=input_csv,
    )

    assert json_path.parent == tmp_path / "output"
    assert csv_path.parent == input_csv.resolve().parent
    assert csv_path.exists()


def test_save_won_accounts_csv_keeps_only_winning_ids(tmp_path):
    input_csv = tmp_path / "ids.csv"
    output_csv = tmp_path / "lottery_result_won.csv"
    service = _service(
        {
            "10000001": ["山田太郎", "ヤマダタロウ", "pass1"],
            "10000002": ["佐藤花子", "サトウハナコ", "pass2"],
        }
    )

    service.save_won_accounts_csv(
        {
            "account_summaries": [
                {
                    "user_id": "10000001",
                    "results": [{"result": "won"}],
                },
                {
                    "user_id": "10000002",
                    "results": [{"result": "lost"}],
                },
            ]
        },
        input_csv,
        output_csv,
    )

    with output_csv.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    assert rows == [["10000001", "山田太郎", "ヤマダタロウ", "pass1"]]


def test_save_confirmed_accounts_csv_keeps_only_successful_ids(tmp_path):
    input_csv = tmp_path / "lottery_result_won.csv"
    output_csv = tmp_path / "reservation_confirmed.csv"
    lottery_service = _service(
        {
            "10000001": ["山田太郎", "ヤマダタロウ", "pass1"],
            "10000002": ["佐藤花子", "サトウハナコ", "pass2"],
        }
    )
    service = ReservationConfirmationWorkflowService(
        lottery_result_workflow_service=lottery_service,
        reservation_service=None,
        login_service=None,
    )

    service.save_confirmed_accounts_csv(
        {
            "reservation_result": {
                "10000001": {"confirmed": ["7月15日 09時00分～11時00分"]},
                "10000002": {"confirmed": []},
            }
        },
        input_csv,
        output_csv,
    )

    with output_csv.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    assert rows == [
        [
            "10000001",
            "山田太郎",
            "ヤマダタロウ",
            "pass1",
            "7月15日 09時00分～11時00分",
        ]
    ]


def test_reservation_confirmation_uses_existing_lottery_result():
    class _LotteryService:
        id_manager_service = _DummyIdManagerService(
            {"10000001": ["山田太郎", "ヤマダタロウ", "pass1"]}
        )

        def resolve_accounts(self, id_csv=None, account_id=None):
            return [
                {
                    "user_id": "10000001",
                    "password": "pass1",
                    "account_label": "山田太郎",
                }
            ]

        def run(self, **kwargs):
            raise AssertionError("lottery result must not be fetched again")

        @staticmethod
        def mask_user_id(user_id):
            return "10***01"

    class _ReservationService:
        def __init__(self):
            self.accounts = None

        def confirm_accounts(
            self, accounts, login_service=None, decision_callback=None
        ):
            self.accounts = accounts
            return {"10000001": {"status": "completed", "confirmed": ["date"]}}

    reservation_service = _ReservationService()
    service = ReservationConfirmationWorkflowService(
        lottery_result_workflow_service=_LotteryService(),
        reservation_service=reservation_service,
        login_service=None,
    )

    result = service.run(
        id_csv="won.csv",
        lottery_result={
            "account_source": "id_csv",
            "results": [
                {
                    "account": "10***01",
                    "account_label": "山田太郎",
                    "result": "won",
                }
            ],
        },
        decision_callback=lambda account: True,
    )

    assert result["confirmed"] is True
    assert [account["user_id"] for account in reservation_service.accounts] == [
        "10000001"
    ]


def test_lottery_result_page_ready_requires_result_page_dom():
    class _Driver:
        page_source = Path("tests/抽選結果画面.html").read_text(encoding="utf-8")

    assert LotteryResultWorkflowService._is_lottery_result_page_ready(_Driver())
