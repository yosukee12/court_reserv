# -*- coding: utf-8 -*-
"""CLI-oriented lottery entry workflow helpers."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path

from court_reserv.config import get_default_credentials


class LotteryEntryWorkflowService:
    """Drive lottery entry slot collection, planning, and optional submission."""

    def __init__(
        self,
        config,
        browser_session,
        login_service,
        navigation_service,
        lottery_service,
        id_manager_service,
        slot_collector,
        slot_adapter=None,
        slot_ranking_service=None,
        logger=None,
    ):
        self.config = config
        self.browser_session = browser_session
        self.login_service = login_service
        self.navigation_service = navigation_service
        self.lottery_service = lottery_service
        self.id_manager_service = id_manager_service
        self.slot_collector = slot_collector
        self.slot_adapter = slot_adapter
        self.slot_ranking_service = slot_ranking_service
        self.logger = logger or logging.getLogger(__name__)

    def resolve_accounts(self, id_csv=None, account_id=None):
        """Resolve accounts by issue-defined priority."""
        if id_csv:
            id_dict = self.id_manager_service.load_accounts(id_csv)
            if not id_dict:
                raise ValueError(f"No accounts found in CSV: {id_csv}")

            target_ids = [account_id] if account_id else list(id_dict.keys())
            accounts = []
            for user_id in target_ids:
                if user_id not in id_dict:
                    raise ValueError(f"Account ID not found in CSV: {user_id}")
                values = id_dict[user_id]
                password = values[2] if len(values) >= 3 else ""
                if not password:
                    raise ValueError(f"Password is empty for account: {user_id}")
                accounts.append(
                    {
                        "user_id": user_id,
                        "password": password,
                        "source": "id_csv",
                        "account_label": values[0] if values else "",
                    }
                )
            return accounts

        config_user_id = self.config.get("AUTH", "USER_ID", fallback="").strip()
        config_password = self.config.get("AUTH", "PASSWORD", fallback="").strip()
        if config_user_id and config_password:
            return [
                {
                    "user_id": config_user_id,
                    "password": config_password,
                    "source": "config.local.ini",
                    "account_label": "",
                }
            ]

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
            return [
                {
                    "user_id": user_id,
                    "password": password,
                    "source": source,
                    "account_label": "",
                }
            ]

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
        del source_csv, search_dirs
        accounts = self.resolve_accounts(id_csv=id_csv, account_id=account_id)
        target_weekdays = (
            preference.lottery_target_weekdays
            if preference.lottery_target_weekdays
            else ["土"]
        )
        result = {
            "account_source": accounts[0]["source"] if accounts else None,
            "target_weekdays": target_weekdays,
            "accounts": [],
        }

        for account in accounts:
            account_result = self._run_for_account(
                account=account,
                preference=preference,
                max_select=max_select,
                display_result_callback=display_result_callback,
                confirm_submit_callback=confirm_submit_callback,
            )
            result["accounts"].append(account_result)

        return result

    def _run_for_account(
        self,
        account,
        preference,
        max_select=2,
        display_result_callback=None,
        confirm_submit_callback=None,
    ):
        account_result = {
            "status": "completed",
            "error": None,
            "credential_source": account["source"],
            "user_id": account["user_id"],
            "masked_user_id": self.mask_user_id(account["user_id"]),
            "account_label": account.get("account_label", ""),
            "target_weekdays": preference.lottery_target_weekdays or ["土"],
            "collected_slots": [],
            "week_search": {
                "max_weeks": 8,
                "weeks_explored": 0,
                "weekly_slot_counts": [],
                "stopped_reason": None,
            },
            "planned_slots": [],
            "missing_slots": [],
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

        driver = self.browser_session.create_driver()
        try:
            if not self.login_service.login(driver, account["user_id"], account["password"]):
                account_result["status"] = "login_failed"
                account_result["error"] = "login_failed"
                if display_result_callback is not None:
                    display_result_callback(account_result)
                return account_result

            self.navigation_service.go_to_lottery_entry(driver)
            self.navigation_service.select_lottery_tennis_park(driver)

            account_entries = self._get_account_entries(account, preference)
            collect_result = self.slot_collector.collect_slots_for_entries(
                driver,
                target_entries=account_entries,
                target_weekdays=preference.lottery_target_weekdays or ["土"],
                max_weeks=8,
            )
            collected_slots = collect_result.get("slots", [])
            account_result["collected_slots"] = [
                self._serialize_slot(slot) for slot in collected_slots
            ]
            account_result["week_search"] = {
                "max_weeks": collect_result.get("max_weeks", 8),
                "weeks_explored": collect_result.get("weeks_explored", 0),
                "weekly_slot_counts": [
                    {
                        "week_index": item.get("week_index"),
                        "slot_count": item.get("slot_count"),
                        "dates": item.get("dates", []),
                        "matched_target_count": item.get("matched_target_count"),
                    }
                    for item in collect_result.get("weekly_summaries", [])
                ],
                "stopped_reason": collect_result.get("stopped_reason"),
            }

            planned_slots, missing_slots = self._build_account_plan(
                account=account,
                preference=preference,
                collected_slots=collected_slots,
                entries=account_entries,
                max_select=max_select,
            )
            account_result["planned_slots"] = [
                self._serialize_slot(slot) for slot in planned_slots
            ]
            account_result["missing_slots"] = missing_slots

            if display_result_callback is not None:
                display_result_callback(account_result)

            if not planned_slots:
                return account_result

            selection_result = self.lottery_service.auto_select_and_submit_slots(
                driver,
                [slot.raw_text for slot in planned_slots if slot.raw_text],
                submit=False,
            )
            account_result["selection_result"] = selection_result

            if confirm_submit_callback is not None:
                confirmation_response = confirm_submit_callback(account_result)
                account_result["confirmation_response"] = confirmation_response
                account_result["submission_requested"] = (
                    str(confirmation_response).strip().lower() == "yes"
                )
                if account_result["submission_requested"]:
                    selected_count = sum(
                        1 for selected in selection_result.values() if selected
                    )
                    submission_result = self.lottery_service.submit_selected_slots(
                        driver,
                        success_count=selected_count,
                    )
                    account_result["submission_result"] = submission_result
                    account_result["submitted"] = (
                        submission_result.get("submitted_count", 0) > 0
                    )
            return account_result
        except Exception as exc:
            self.logger.exception(
                "Lottery entry workflow failed for %s", account["user_id"]
            )
            account_result["status"] = "error"
            account_result["error"] = str(exc)
            if display_result_callback is not None:
                display_result_callback(account_result)
            return account_result
        finally:
            self.browser_session.safe_close(driver)

    def _build_account_plan(
        self,
        account,
        preference,
        collected_slots,
        entries=None,
        max_select=2,
    ):
        del account
        entries = entries if isinstance(entries, list) else []

        desired_limit = min(
            int(max_select or preference.lottery_max_entries_per_account or 2),
            int(preference.lottery_max_entries_per_account or 2),
            2,
        )

        unique_keys = set()
        planned_slots = []
        missing_slots = []

        for entry in entries:
            if not isinstance(entry, dict):
                continue
            date = str(entry.get("date", "")).strip()
            time_range = str(entry.get("time_range", "")).strip()
            facility = str(entry.get("facility", "")).strip()
            key = (date, time_range)
            if not date or not time_range or key in unique_keys:
                continue
            unique_keys.add(key)
            if len(planned_slots) >= desired_limit:
                missing_slots.append(
                    {
                        "date": date,
                        "time_range": time_range,
                        "facility": facility,
                        "warning": "skipped because max 2 entries per account is enforced",
                    }
                )
                continue

            matched_slot = self._match_slot(collected_slots, date, time_range, facility)
            if matched_slot is None:
                missing_slots.append(
                    {
                        "date": date,
                        "time_range": time_range,
                        "facility": facility,
                        "warning": "slot not found within explored lottery entry weeks",
                    }
                )
                continue
            planned_slots.append(matched_slot)

        return planned_slots, missing_slots

    def _get_account_entries(self, account, preference):
        entries = preference.lottery_account_overrides.get(
            account["user_id"],
            preference.lottery_default_entries,
        )
        if not isinstance(entries, list):
            return []
        return entries

    def _match_slot(self, collected_slots, date, time_range, facility):
        for slot in collected_slots:
            if slot.date != date or slot.time_range != time_range:
                continue
            if facility:
                searchable = " ".join(
                    part for part in (slot.park_name, slot.facility_name) if part
                )
                if facility not in searchable:
                    continue
            return slot
        return None

    def print_result(self, result):
        print(f"Account source: {result.get('account_source')}")
        print(f"Target weekdays: {', '.join(result.get('target_weekdays', []))}")
        for account_result in result.get("accounts", []):
            account_part = account_result.get("masked_user_id")
            if account_result.get("account_label"):
                account_part = f"{account_part} ({account_result.get('account_label')})"
            print(f"Account: {account_part}")
            print(f"- status={account_result.get('status')}")
            if account_result.get("error"):
                print(f"- error={account_result.get('error')}")
            print(f"- collected_slots={len(account_result.get('collected_slots', []))}")
            week_search = account_result.get("week_search", {})
            if week_search:
                print(
                    "- week search: explored={explored}/{max_weeks} stopped_reason={reason}".format(
                        explored=week_search.get("weeks_explored", 0),
                        max_weeks=week_search.get("max_weeks", 0),
                        reason=week_search.get("stopped_reason"),
                    )
                )
                for week_info in week_search.get("weekly_slot_counts", []):
                    print(
                        "  week {week}: slots={slots} dates={dates}".format(
                            week=week_info.get("week_index"),
                            slots=week_info.get("slot_count"),
                            dates=",".join(week_info.get("dates", [])),
                        )
                    )

            planned_slots = account_result.get("planned_slots", [])
            if planned_slots:
                print("- planned slots:")
                for slot in planned_slots:
                    print(
                        "  {date} {weekday} {time_range} facility={facility} applied={applied}".format(
                            date=slot.get("date"),
                            weekday=slot.get("weekday"),
                            time_range=slot.get("time_range"),
                            facility=" ".join(
                                part
                                for part in (
                                    slot.get("park_name", ""),
                                    slot.get("facility_name", ""),
                                )
                                if part
                            ).strip(),
                            applied=slot.get("applied_count"),
                        )
                    )
            else:
                print("- no planned slots")

            missing_slots = account_result.get("missing_slots", [])
            if missing_slots:
                print("- warnings:")
                for warning in missing_slots:
                    print(
                        "  {date} {time_range} facility={facility} warning={warning_text}".format(
                            date=warning.get("date"),
                            time_range=warning.get("time_range"),
                            facility=warning.get("facility", ""),
                            warning_text=warning.get("warning"),
                        )
                    )

            selection_result = account_result.get("selection_result", {})
            if selection_result:
                print("- selection result:")
                for slot_text, selected in selection_result.items():
                    print(f"  {slot_text}: {'selected' if selected else 'not selected'}")

            if account_result.get("submission_requested"):
                submission_result = account_result.get("submission_result", {})
                print(
                    "- submission: requested={requested} submitted={submitted} completed={completed} recovery_triggered={recovery_triggered} recovery_attempts={recovery_attempts}".format(
                        requested=submission_result.get("requested_count", 0),
                        submitted=submission_result.get("submitted_count", 0),
                        completed=submission_result.get("completed", False),
                        recovery_triggered=submission_result.get(
                            "recovery_triggered", False
                        ),
                        recovery_attempts=submission_result.get(
                            "recovery_attempts", 0
                        ),
                    )
                )
                debug_files = submission_result.get("debug_files", [])
                if debug_files:
                    print("- submission debug files:")
                    for path in debug_files:
                        print(f"  {path}")
            else:
                print("- submission was not executed")

    def save_result(self, result, output_dir):
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        output_file = output_path / "lottery_entry_workflow_result.json"
        output_file.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return output_file

    def mask_user_id(self, user_id):
        text = str(user_id)
        if len(text) <= 4:
            return "*" * len(text)
        return f"{text[:2]}***{text[-2:]}"

    def _serialize_slot(self, slot):
        return {
            "park_name": slot.park_name,
            "facility_name": slot.facility_name,
            "date": slot.date,
            "weekday": slot.weekday,
            "time_range": slot.time_range,
            "start_time": slot.start_time,
            "end_time": slot.end_time,
            "field_number": slot.field_number,
            "available_count": slot.available_count,
            "applied_count": slot.applied_count,
            "raw_text": slot.raw_text,
        }

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
