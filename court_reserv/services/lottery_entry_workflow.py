# -*- coding: utf-8 -*-
"""CLI-oriented lottery entry workflow helpers."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path

from court_reserv.config import get_default_credentials
from court_reserv.services.lottery_automation import LotteryAutomationDryRunService
from court_reserv.services.slot_ranking import RankedSlot


class LotteryEntryWorkflowService:
    """Drive ranked candidate selection up to the pre-submit lottery page."""

    def __init__(
        self,
        config,
        browser_session,
        login_service,
        navigation_service,
        lottery_service,
        id_manager_service,
        slot_adapter,
        slot_ranking_service,
        logger=None,
    ):
        self.config = config
        self.browser_session = browser_session
        self.login_service = login_service
        self.navigation_service = navigation_service
        self.lottery_service = lottery_service
        self.id_manager_service = id_manager_service
        self.slot_adapter = slot_adapter
        self.slot_ranking_service = slot_ranking_service
        self.logger = logger or logging.getLogger(__name__)
        self.dry_run_service = LotteryAutomationDryRunService(
            slot_adapter=self.slot_adapter,
            slot_ranking_service=self.slot_ranking_service,
            logger=self.logger,
        )

    def resolve_credentials(self, id_csv=None, account_id=None):
        """Resolve credentials by issue-defined priority."""
        if id_csv:
            id_dict = self.id_manager_service.load_accounts(id_csv)
            if not id_dict:
                raise ValueError(f"No accounts found in CSV: {id_csv}")

            if account_id:
                if account_id not in id_dict:
                    raise ValueError(f"Account ID not found in CSV: {account_id}")
                user_id = account_id
            else:
                user_id = next(iter(id_dict))

            values = id_dict[user_id]
            password = values[2] if len(values) >= 3 else ""
            if not password:
                raise ValueError(f"Password is empty for account: {user_id}")

            return {
                "user_id": user_id,
                "password": password,
                "source": "id_csv",
                "account_label": values[0] if values else "",
            }

        config_user_id = self.config.get("AUTH", "USER_ID", fallback="").strip()
        config_password = self.config.get("AUTH", "PASSWORD", fallback="").strip()
        if config_user_id and config_password:
            return {
                "user_id": config_user_id,
                "password": config_password,
                "source": "config.local.ini",
                "account_label": "",
            }

        user_id, password = get_default_credentials()
        if user_id and password:
            env_values = self._read_env_file()
            source = ".env"
            if os.environ.get("COURT_RESERV_USER_ID") or os.environ.get(
                "COURT_RESERV_PASSWORD"
            ):
                source = "environment"
            elif env_values.get("COURT_RESERV_USER_ID") or env_values.get(
                "COURT_RESERV_PASSWORD"
            ):
                source = ".env"
            return {
                "user_id": user_id,
                "password": password,
                "source": source,
                "account_label": "",
            }

        raise ValueError(
            "Credentials were not found. Configure an ID CSV, config.local.ini, or .env."
        )

    def run(
        self,
        preference,
        source_csv=None,
        search_dirs=None,
        id_csv=None,
        account_id=None,
        max_select=2,
        display_result_callback=None,
        confirm_submit_callback=None,
    ):
        dry_run_result = self.dry_run_service.run(
            preference=preference,
            source_csv=source_csv,
            search_dirs=search_dirs,
        )
        selected_candidates, skipped_duplicates = self.select_candidates(
            dry_run_result.get("ranked_candidates", []),
            max_select=max_select,
        )

        result = {
            "credential_source": None,
            "user_id": None,
            "account_label": "",
            "source_csv": dry_run_result.get("source_csv"),
            "total_slots": dry_run_result.get("total_slots", 0),
            "selected_candidates": selected_candidates,
            "skipped_duplicate_datetimes": skipped_duplicates,
            "selection_result": {},
            "submission_requested": False,
            "confirmation_response": None,
            "submission_result": {
                "requested_count": 0,
                "submitted_count": 0,
                "completed": False,
            },
            "submitted": False,
        }

        if not selected_candidates:
            return result

        auth = self.resolve_credentials(id_csv=id_csv, account_id=account_id)
        result["credential_source"] = auth["source"]
        result["user_id"] = auth["user_id"]
        result["account_label"] = auth.get("account_label", "")

        selected_slot_texts = [
            self._to_slot_text(ranked_slot) for ranked_slot in selected_candidates
        ]

        driver = self.browser_session.create_driver()
        try:
            if not self.login_service.login(driver, auth["user_id"], auth["password"]):
                raise RuntimeError(f"Login failed for user: {auth['user_id']}")

            self.navigation_service.go_to_lottery_entry(driver)
            self.navigation_service.select_lottery_tennis_park(driver)
            selection_result = self.lottery_service.auto_select_and_submit_slots(
                driver,
                selected_slot_texts,
                submit=False,
            )
            result["selection_result"] = selection_result
            if display_result_callback is not None:
                display_result_callback(result)

            if confirm_submit_callback is not None:
                confirmation_response = confirm_submit_callback(result)
                result["confirmation_response"] = confirmation_response
                result["submission_requested"] = str(confirmation_response).strip().lower() == "yes"
                if result["submission_requested"]:
                    selected_count = sum(
                        1 for selected in selection_result.values() if selected
                    )
                    submission_result = self.lottery_service.submit_selected_slots(
                        driver,
                        success_count=selected_count,
                    )
                    result["submission_result"] = submission_result
                    result["submitted"] = submission_result.get("submitted_count", 0) > 0
            return result
        finally:
            self.browser_session.safe_close(driver)

    def select_candidates(self, ranked_candidates, max_select=2):
        """Choose up to two unique datetime candidates."""
        selected = []
        skipped_duplicates = []
        seen_datetimes = set()
        limit = max(1, min(int(max_select), 2))

        for ranked in ranked_candidates:
            key = (ranked.slot.date, ranked.slot.time_range)
            if key in seen_datetimes:
                skipped_duplicates.append(
                    {
                        "date": ranked.slot.date,
                        "time_range": ranked.slot.time_range,
                        "raw_text": ranked.slot.raw_text,
                    }
                )
                continue
            seen_datetimes.add(key)
            selected.append(ranked)
            if len(selected) >= limit:
                break

        return selected, skipped_duplicates

    def print_result(self, result):
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
                f"{index}. score={ranked.score} slot={self._to_slot_text(ranked)} reasons={reasons}"
            )

        skipped_duplicates = result.get("skipped_duplicate_datetimes", [])
        if skipped_duplicates:
            print("Skipped duplicate datetime candidates:")
            for skipped in skipped_duplicates:
                print(
                    f"- {skipped.get('date')} {skipped.get('time_range')} {skipped.get('raw_text') or ''}".strip()
                )

        selection_result = result.get("selection_result", {})
        if selection_result:
            print("Lottery page selection result:")
            for slot_text, selected in selection_result.items():
                print(f"- {slot_text}: {'selected' if selected else 'not selected'}")
        if result.get("submission_requested"):
            submission_result = result.get("submission_result", {})
            print("Lottery submission result:")
            print(
                "- requested={requested} submitted={submitted} completed={completed}".format(
                    requested=submission_result.get("requested_count", 0),
                    submitted=submission_result.get("submitted_count", 0),
                    completed=submission_result.get("completed", False),
                )
            )
        else:
            print("Lottery submission was not executed.")

    def save_result(self, result, output_dir):
        """Persist a minimal JSON summary for manual review."""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        payload = {
            "credential_source": result.get("credential_source"),
            "user_id": result.get("user_id"),
            "account_label": result.get("account_label"),
            "source_csv": result.get("source_csv"),
            "total_slots": result.get("total_slots", 0),
            "selected_candidates": [
                {
                    "score": ranked.score,
                    "reasons": ranked.reasons,
                    "slot": {
                        "date": ranked.slot.date,
                        "time_range": ranked.slot.time_range,
                        "court_name": ranked.slot.court_name,
                        "applied_count": ranked.slot.applied_count,
                        "raw_text": ranked.slot.raw_text,
                    },
                }
                for ranked in result.get("selected_candidates", [])
            ],
            "skipped_duplicate_datetimes": result.get("skipped_duplicate_datetimes", []),
            "selection_result": result.get("selection_result", {}),
            "submission_requested": result.get("submission_requested", False),
            "confirmation_response": result.get("confirmation_response"),
            "submission_result": result.get("submission_result", {}),
            "submitted": result.get("submitted", False),
        }
        output_file = output_path / "lottery_entry_workflow_result.json"
        output_file.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return output_file

    def _to_slot_text(self, ranked_slot: RankedSlot) -> str:
        slot = ranked_slot.slot
        if slot.raw_text:
            return slot.raw_text
        ymd = slot.date.replace("-", "") if slot.date else ""
        if slot.time_range and "-" in slot.time_range:
            start, end = slot.time_range.split("-", 1)
            compact_time = f"{start.replace(':', '')}-{end.replace(':', '')}"
        else:
            compact_time = slot.time_range or ""
        parts = [part for part in (ymd, compact_time, slot.court_name) if part]
        return " ".join(parts)

    def _read_env_file(self):
        env_values = {}
        for env_path in (Path(".env"), Path("court_reserv/.env")):
            if not env_path.exists():
                continue
            for line in env_path.read_text(encoding="utf-8").splitlines():
                stripped = line.strip()
                if not stripped or stripped.startswith("#") or "=" not in stripped:
                    continue
                key, value = stripped.split("=", 1)
                env_values[key.strip()] = value.strip().strip("'\"")
        return env_values
