# -*- coding: utf-8 -*-
"""CLI-oriented reservation confirmation workflow helpers."""

from __future__ import annotations

import json
import logging
from pathlib import Path


class ReservationConfirmationWorkflowService:
    """Confirm reservations only after user selection and explicit approval."""

    def __init__(
        self,
        lottery_result_workflow_service,
        reservation_service,
        login_service,
        logger=None,
    ):
        self.lottery_result_workflow_service = lottery_result_workflow_service
        self.reservation_service = reservation_service
        self.login_service = login_service
        self.logger = logger or logging.getLogger(__name__)

    def run(
        self,
        id_csv=None,
        account_id=None,
        select_entries_callback=None,
        confirm_callback=None,
    ):
        accounts = self.lottery_result_workflow_service.resolve_accounts(
            id_csv=id_csv,
            account_id=account_id,
        )
        lottery_result = self.lottery_result_workflow_service.run(
            id_csv=id_csv,
            account_id=account_id,
        )
        won_entries = [
            row for row in lottery_result.get("results", []) if row.get("result") == "won"
        ]
        result = {
            "status": "completed",
            "account_source": lottery_result.get("account_source"),
            "won_entries": won_entries,
            "selected_entries": [],
            "selected_accounts": [],
            "confirmation_response": None,
            "confirmed": False,
            "reservation_result": {},
            "selection_error": None,
        }
        if not won_entries:
            return result

        if select_entries_callback is None:
            return result

        selected_indices = select_entries_callback(won_entries)
        selected_entries = [
            won_entries[index]
            for index in selected_indices
            if 0 <= index < len(won_entries)
        ]
        result["selected_entries"] = selected_entries
        if not selected_entries:
            return result

        validation = self._resolve_selected_accounts(selected_entries, won_entries, accounts)
        if validation["error"]:
            result["selection_error"] = validation["error"]
            return result
        result["selected_accounts"] = validation["selected_accounts"]

        if confirm_callback is None:
            return result

        confirmation_response = confirm_callback(result)
        result["confirmation_response"] = confirmation_response
        if str(confirmation_response).strip().lower() != "yes":
            return result

        reservation_result = self.reservation_service.confirm_accounts(
            validation["selected_accounts"],
            login_service=self.login_service,
        )
        result["reservation_result"] = reservation_result
        result["confirmed"] = True
        return result

    def _resolve_selected_accounts(self, selected_entries, won_entries, accounts):
        selected_keys = {
            self._account_key(entry.get("account"), entry.get("account_label", ""))
            for entry in selected_entries
        }
        all_won_by_key = {}
        for entry in won_entries:
            key = self._account_key(entry.get("account"), entry.get("account_label", ""))
            all_won_by_key.setdefault(key, []).append(entry)

        for key in selected_keys:
            if len(all_won_by_key.get(key, [])) != len(
                [entry for entry in selected_entries if self._account_key(entry.get("account"), entry.get("account_label", "")) == key]
            ):
                return {
                    "selected_accounts": [],
                    "error": (
                        "Partial selection for the same account is not supported. "
                        "Select all won entries for an account or none."
                    ),
                }

        account_map = {}
        for account in accounts:
            masked = self.lottery_result_workflow_service.mask_user_id(account["user_id"])
            key = self._account_key(masked, account.get("account_label", ""))
            if key in account_map:
                return {
                    "selected_accounts": [],
                    "error": "Ambiguous account mapping detected. Narrow the selection to a single account.",
                }
            account_map[key] = account

        selected_accounts = []
        for key in selected_keys:
            account = account_map.get(key)
            if account is None:
                return {
                    "selected_accounts": [],
                    "error": "Failed to resolve selected account from lottery results.",
                }
            selected_accounts.append(account)

        return {"selected_accounts": selected_accounts, "error": None}

    def print_won_entries(self, won_entries):
        if not won_entries:
            print("No won entries found.")
            return
        print("Won entries:")
        for index, row in enumerate(won_entries, start=1):
            account_part = row.get("account")
            if row.get("account_label"):
                account_part = f"{account_part} ({row.get('account_label')})"
            print(
                f"{index}. account={account_part} date={row.get('date')} "
                f"time={row.get('time_range')} facility={row.get('facility')}"
            )

    def print_result(self, result):
        print(f"Account source: {result.get('account_source')}")
        self.print_won_entries(result.get("won_entries", []))
        if result.get("selection_error"):
            print(f"Selection error: {result.get('selection_error')}")
            return
        selected_entries = result.get("selected_entries", [])
        if selected_entries:
            print("Selected entries:")
            for index, row in enumerate(selected_entries, start=1):
                account_part = row.get("account")
                if row.get("account_label"):
                    account_part = f"{account_part} ({row.get('account_label')})"
                print(
                    f"{index}. account={account_part} date={row.get('date')} "
                    f"time={row.get('time_range')} facility={row.get('facility')}"
                )
        if result.get("confirmed"):
            print("Reservation confirmation result:")
            for user_id, reservation_result in result.get("reservation_result", {}).items():
                print(
                    f"- account={self.lottery_result_workflow_service.mask_user_id(user_id)} "
                    f"status={reservation_result.get('status')} "
                    f"confirmed={reservation_result.get('confirmed', [])}"
                )
        else:
            print("Reservation confirmation was not executed.")

    def save_result(self, result, output_dir):
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        json_path = output_path / "reservation_confirmation_workflow_result.json"
        payload = {
            "status": result.get("status"),
            "account_source": result.get("account_source"),
            "won_entries": result.get("won_entries", []),
            "selected_entries": result.get("selected_entries", []),
            "selected_accounts": [
                {
                    "account": self.lottery_result_workflow_service.mask_user_id(
                        account["user_id"]
                    ),
                    "account_label": account.get("account_label", ""),
                }
                for account in result.get("selected_accounts", [])
            ],
            "confirmation_response": result.get("confirmation_response"),
            "confirmed": result.get("confirmed", False),
            "reservation_result": {
                self.lottery_result_workflow_service.mask_user_id(user_id): value
                for user_id, value in result.get("reservation_result", {}).items()
            },
            "selection_error": result.get("selection_error"),
        }
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return json_path

    def _account_key(self, masked_account, account_label):
        return f"{masked_account}::{account_label or ''}"
