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
        output_dir=None,
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
            "session_strategy": "created_per_account",
        }

        for index, account in enumerate(accounts):
            account_result = self._run_account_with_retry(
                account=account,
                preference=preference,
                max_select=max_select,
                display_result_callback=display_result_callback,
                confirm_submit_callback=confirm_submit_callback,
                account_index=index + 1,
            )
            result["accounts"].append(account_result)
            if output_dir is not None:
                self.save_result(result, output_dir)
        return result

    def _run_account_with_retry(
        self,
        account,
        preference,
        max_select,
        display_result_callback,
        confirm_submit_callback,
        account_index,
    ):
        retryable = {
            "login_or_navigation_not_ready",
            "login_post_page_not_ready",
            "login_form_not_ready",
            "login_form_not_found",
            "doAction_not_defined",
            "doLotEntry_not_ready",
            "lottery_park_selection_not_ready",
        }
        last_result = None
        for attempt in range(2):
            driver = self.browser_session.create_driver()
            session_reused = False
            session_mode = "created_per_account"
            masked_user_id = self.mask_user_id(account["user_id"])
            self.logger.info(
                "Starting account index=%s user=%s session_mode=%s session_reused=%s",
                account_index,
                masked_user_id,
                session_mode,
                session_reused,
            )
            try:
                last_result = self._run_for_account(
                    driver=driver,
                    account=account,
                    preference=preference,
                    max_select=max_select,
                    display_result_callback=display_result_callback,
                    confirm_submit_callback=confirm_submit_callback,
                    session_reused=session_reused,
                    session_mode=session_mode,
                    account_index=account_index,
                    attempt_index=attempt + 1,
                )
            finally:
                self.browser_session.safe_quit(driver)
            if not last_result or last_result.get("status") not in retryable or attempt >= 1:
                return last_result
            if last_result.get("retry_reason") == "login_alert":
                return last_result
            self.logger.info(
                "Retrying account index=%s reason=%s",
                account_index,
                last_result.get("retry_reason"),
            )
        return last_result

    def _run_for_account(
        self,
        driver,
        account,
        preference,
        max_select=2,
        display_result_callback=None,
        confirm_submit_callback=None,
        session_reused=False,
        session_mode="created",
        account_index=1,
        attempt_index=1,
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
            "session_reused": bool(session_reused),
            "session_mode": session_mode,
            "account_index": account_index,
            "login_attempts": attempt_index,
            "login_retry_used": attempt_index > 1,
            "navigation_ready": False,
            "navigation_state": {},
            "login_state": {},
            "last_page_state": {},
            "retry_reason": None,
            "entries": [],
        }
        try:
            if not self.login_service.login(driver, account["user_id"], account["password"]):
                account_result["login_state"] = dict(
                    getattr(self.login_service, "last_login_state", {}) or {}
                )
                account_result["last_page_state"] = self.login_service.inspect_page_state(driver)
                login_error = getattr(self.login_service, "last_login_error", None) or "login_failed"
                account_result["status"] = login_error
                account_result["error"] = login_error
                account_result["retry_reason"] = login_error
                if display_result_callback is not None:
                    display_result_callback(account_result)
                return account_result

            account_result["login_state"] = dict(
                getattr(self.login_service, "last_login_state", {}) or {}
            )
            navigation_ready = self.navigation_service.wait_until_navigation_ready(driver)
            account_result["navigation_ready"] = bool(navigation_ready)
            account_result["navigation_state"] = self.login_service.inspect_page_state(driver)
            self.logger.info(
                "before go_to_lottery_entry navigation_ready=%s has_doAction=%s has_lottery_action=%s retry_count=%s",
                navigation_ready,
                account_result["navigation_state"].get("has_doAction"),
                account_result["navigation_state"].get("has_lottery_action"),
                attempt_index - 1,
            )
            if not self.navigation_service.go_to_lottery_entry(driver):
                account_result["status"] = "login_or_navigation_not_ready"
                account_result["error"] = "login_or_navigation_not_ready"
                account_result["retry_reason"] = "login_or_navigation_not_ready"
                account_result["last_page_state"] = self.login_service.inspect_page_state(driver)
                if display_result_callback is not None:
                    display_result_callback(account_result)
                return account_result
            try:
                self.navigation_service.select_lottery_tennis_park(driver)
            except Exception as exc:
                message = str(exc)
                status = (
                    "doLotEntry_not_ready"
                    if "doLotEntry" in message
                    else "lottery_park_selection_not_ready"
                )
                account_result["status"] = status
                account_result["error"] = message
                account_result["retry_reason"] = status
                account_result["last_page_state"] = self.login_service.inspect_page_state(driver)
                if display_result_callback is not None:
                    display_result_callback(account_result)
                return account_result

            account_entries = self._get_account_entries(account, preference)
            collect_result = self.slot_collector.collect_slots_for_entries(
                driver,
                target_entries=account_entries,
                target_weekdays=preference.lottery_target_weekdays or ["土"],
                max_weeks=preference.lottery_search_weeks or 8,
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

            submission_entries = []
            for entry_index, slot in enumerate(planned_slots[:2], start=1):
                entry_result = self._build_entry_result(
                    account_result=account_result,
                    slot=slot,
                    entry_index=entry_index,
                    account_index=account_index,
                )
                account_result["entries"].append(entry_result)
                skip_ensure_before_select = False
                self.logger.info(
                    "Starting entry account_index=%s entry_index=%s apply_no=%s date=%s weekday=%s time_range=%s facility=%s current_entry_count=%s",
                    account_index,
                    entry_index,
                    entry_result["apply_no"],
                    slot.date,
                    slot.weekday,
                    slot.time_range,
                    entry_result["slot"]["facility"],
                    entry_result["slot"]["current_entry_count"],
                )

                if entry_index > 1 and any(
                    item.get("status") == "completed" for item in submission_entries
                ):
                    continue_result = {
                        "success": False,
                        "display_no": "",
                        "fallback_used": False,
                        "title": "",
                    }
                    try:
                        continue_result = self.lottery_service.continue_after_lottery_completion(
                            driver
                        )
                    except Exception:
                        self.logger.warning(
                            "Continue-after-completion failed; reloading lottery entry page."
                        )
                    if continue_result.get("success"):
                        self.logger.info(
                            "continue_after_completion executed=true success=true displayNo=%s title=%s url=%s",
                            continue_result.get("display_no"),
                            continue_result.get("title"),
                            continue_result.get("current_url"),
                        )
                    else:
                        continue_result["fallback_used"] = True
                        continue_result["fallback_result"] = self._ensure_lottery_entry_page(
                            driver, force_navigation=True
                        )
                        self.logger.info(
                            "continue_after_completion executed=true success=false fallback_used=true displayNo=%s title=%s url=%s",
                            continue_result.get("display_no"),
                            continue_result.get("title"),
                            continue_result.get("current_url"),
                        )
                    park_selection_state = self._ensure_lottery_tennis_park_selection(
                        driver,
                        account_index=account_index,
                        entry_index=entry_index,
                    )
                    continue_result["park_selection_state"] = park_selection_state
                    entry_validation = self.lottery_service.build_park_facility_validation(
                        slot,
                        selection_state=park_selection_state,
                    )
                    entry_result["validation"] = entry_validation
                    entry_result["continue_after_completion"] = continue_result
                    if not entry_validation.get("park_facility_match", True):
                        entry_result["search_after_continue"] = {
                            "executed": False,
                            "weeks_explored": 0,
                            "matched": False,
                            "stopped_reason": entry_validation.get("status"),
                            "park_selection_state": park_selection_state,
                        }
                        entry_result["status"] = entry_validation.get(
                            "status", "park_mismatch"
                        )
                        entry_result["error_message"] = (
                            entry_validation.get("mismatch_reason")
                            or "park/facility mismatch after continue"
                        )
                        entry_result.setdefault("debug_files", [])
                        entry_result["debug_files"].extend(
                            self.lottery_service._save_submission_debug(
                                driver,
                                prefix=f"lottery_park_mismatch_account{account_index}_entry{entry_index}",
                            )
                        )
                        submission_entries.append(entry_result)
                        break
                    refreshed = self._refresh_slot_for_entry(
                        driver=driver,
                        slot=slot,
                        target_weekdays=preference.lottery_target_weekdays or ["土"],
                        max_weeks=preference.lottery_search_weeks or 8,
                    )
                    entry_result["search_after_continue"] = {
                        "weeks_explored": refreshed.get("weeks_explored", 0),
                        "matched": bool(refreshed.get("slot")),
                        "stopped_reason": refreshed.get("stopped_reason"),
                        "same_date_count": refreshed.get("same_date_count", 0),
                        "same_time_count": refreshed.get("same_time_count", 0),
                        "park_selection_state": refreshed.get("park_selection_state", {}),
                        "facility_reselected": refreshed.get("facility_reselected", False),
                        "retry_used": refreshed.get("retry_used", False),
                    }
                    self.logger.info(
                        "search_after_continue executed=true weeks_explored=%s matched=%s stopped_reason=%s same_date_count=%s same_time_count=%s ensure_skipped_before_select=%s facility_reselected=%s retry_used=%s",
                        refreshed.get("weeks_explored", 0),
                        bool(refreshed.get("slot")),
                        refreshed.get("stopped_reason"),
                        refreshed.get("same_date_count", 0),
                        refreshed.get("same_time_count", 0),
                        bool(refreshed.get("slot")),
                        refreshed.get("facility_reselected", False),
                        refreshed.get("retry_used", False),
                    )
                    if refreshed.get("slot") is not None:
                        slot = refreshed["slot"]
                        skip_ensure_before_select = True
                        entry_result["slot"] = self._build_entry_result(
                            account_result=account_result,
                            slot=slot,
                            entry_index=entry_index,
                            account_index=account_index,
                        )["slot"]
                    else:
                        entry_result["ensure_result"] = self._ensure_lottery_entry_page(
                            driver
                        )
                        entry_result["error_message"] = (
                            "search_after_continue did not match target slot"
                        )
                        entry_result["status"] = "search_unmatched"
                        submission_entries.append(entry_result)
                        break

                if entry_index == 1:
                    entry_result["ensure_result"] = self._ensure_lottery_entry_page(driver)
                elif not skip_ensure_before_select:
                    entry_result["ensure_result"] = self._ensure_lottery_entry_page(driver)
                else:
                    entry_result["ensure_result"] = {
                        "called": False,
                        "skipped": True,
                        "reason": "matched_after_refresh",
                    }
                entry_result["ensure_skipped_before_select"] = skip_ensure_before_select
                self.logger.info(
                    "ensure executed=%s skipped=%s result=%s",
                    entry_result["ensure_result"].get("called", False),
                    entry_result["ensure_skipped_before_select"],
                    json.dumps(entry_result["ensure_result"], ensure_ascii=False),
                )

                confirmation_response = ""
                if confirm_submit_callback is not None:
                    confirmation_response = confirm_submit_callback(
                        account_result,
                        entry_result,
                    )
                entry_result["confirmation_response"] = confirmation_response
                account_result["confirmation_response"] = confirmation_response
                response_key = str(confirmation_response).strip().lower()
                if response_key == "yes":
                    account_result["submission_requested"] = True
                    selected = self.lottery_service.select_single_slot(driver, slot.raw_text)
                    account_result["selection_result"][slot.raw_text] = selected
                    entry_result["selected"] = bool(selected)
                    entry_result["select_result"] = self.lottery_service.capture_lottery_selection_state(
                        driver,
                        slot.raw_text,
                        selected=bool(selected),
                    )
                    select_validation = self.lottery_service.build_park_facility_validation(
                        slot,
                        selection_state=entry_result["select_result"],
                    )
                    entry_result["validation"] = select_validation
                    self.logger.info(
                        "select_single_slot success=%s raw_text=%s selectFieldCnt=%s selected_slots_count=%s checked_input_count=%s checked_selectUseYMD=%s checked_selectStime=%s checked_selectEtime=%s checked_selectField=%s headers=%s displayNo=%s selection_applied_to_form=%s expected_park=%s expected_facility=%s actual_park=%s actual_facility=%s current_bname=%s current_iname=%s clicked_element_outer_html=%s target_cell_outer_html=%s",
                        bool(selected),
                        slot.raw_text,
                        entry_result["select_result"].get("selectFieldCnt"),
                        entry_result["select_result"].get("selected_slots_count"),
                        entry_result["select_result"].get("checked_input_count"),
                        entry_result["select_result"].get("checked_selectUseYMD"),
                        entry_result["select_result"].get("checked_selectStime"),
                        entry_result["select_result"].get("checked_selectEtime"),
                        entry_result["select_result"].get("checked_selectField"),
                        entry_result["select_result"].get("headers"),
                        entry_result["select_result"].get("current_display_no"),
                        entry_result["select_result"].get("selection_applied_to_form"),
                        select_validation.get("expected_park_name"),
                        select_validation.get("expected_facility_name"),
                        select_validation.get("actual_slot_park_name"),
                        select_validation.get("actual_slot_facility_name"),
                        select_validation.get("current_bname_text"),
                        select_validation.get("current_iname_text"),
                        entry_result["select_result"].get("clicked_element_outer_html"),
                        entry_result["select_result"].get("target_cell_outer_html"),
                    )
                    if not select_validation.get("park_facility_match", True):
                        entry_result["status"] = select_validation.get("status", "park_mismatch")
                        entry_result["error_message"] = (
                            select_validation.get("mismatch_reason")
                            or "park/facility mismatch after selection"
                        )
                        entry_result.setdefault("debug_files", [])
                        entry_result["debug_files"].extend(
                            self.lottery_service._save_submission_debug(
                                driver,
                                prefix=f"lottery_park_mismatch_account{account_index}_entry{entry_index}",
                            )
                        )
                        submission_entries.append(entry_result)
                        break
                    if str(entry_result["select_result"].get("selectFieldCnt", "")) != "1":
                        entry_result["status"] = "stopped"
                        entry_result["error_message"] = "selectFieldCnt is not 1"
                        submission_entries.append(entry_result)
                        break
                    if not entry_result["select_result"].get("selection_applied_to_form"):
                        entry_result.setdefault("debug_files", [])
                        entry_result["debug_files"] = self.lottery_service.save_before_apply_debug(
                            driver,
                            account_index=account_index,
                            entry_index=entry_index,
                        )
                        entry_result["status"] = "slot_selection_not_applied"
                        entry_result["error_message"] = (
                            "slot selection was not applied to the submit form"
                        )
                        submission_entries.append(entry_result)
                        break
                    if not selected:
                        entry_result["status"] = "selection_failed"
                        entry_result["error_message"] = "select_single_slot failed"
                        submission_entries.append(entry_result)
                        break
                    submission_result = self.lottery_service.submit_single_selected_slot(
                        driver,
                        apply_no=entry_result["apply_no"],
                        expected_slot=entry_result["slot"],
                        account_index=account_index,
                        entry_index=entry_index,
                        select_result=entry_result["select_result"],
                        manual_final_submit=bool(
                            getattr(preference, "lottery_manual_final_submit", False)
                        ),
                        manual_preconfirm_submit=bool(
                            getattr(preference, "lottery_manual_preconfirm_submit", False)
                        ),
                    )
                    entry_result["validation"] = submission_result.get(
                        "validation", entry_result.get("validation", {})
                    )
                    entry_result["confirm_page"] = submission_result.get("confirm_page", {})
                    entry_result["submit_result"] = submission_result.get("submit_result", {})
                    entry_result["recaptcha_recovery"] = submission_result.get(
                        "recaptcha_recovery", {}
                    )
                    entry_result["submission_result"] = submission_result
                    entry_result["error_message"] = submission_result.get("error_message")
                    entry_result["status"] = submission_result.get(
                        "status",
                        "completed"
                        if submission_result.get("completed")
                        else "stopped"
                        if submission_result.get("stopped")
                        else "submission_failed",
                    )
                    entry_result["submitted"] = bool(
                        submission_result.get("submitted_count", 0)
                    )
                    self.logger.info(
                        "submit_result alert_text=%s completion_detected=%s recaptcha_detected=%s recovery_attempted=%s recovery_retry_count=%s entry_status=%s",
                        submission_result.get("submit_result", {}).get("alert_text"),
                        submission_result.get("submit_result", {}).get("completion_detected"),
                        submission_result.get("submit_result", {}).get("recaptcha_detected"),
                        submission_result.get("recaptcha_recovery", {}).get("attempted"),
                        submission_result.get("recaptcha_recovery", {}).get("retry_count"),
                        entry_result["status"],
                    )
                    if entry_result["status"] != "completed":
                        submission_entries.append(entry_result)
                        break
                elif response_key == "dry-run":
                    entry_result["status"] = "dry_run"
                    self.logger.info(
                        "Dry run: submission skipped for %s %s",
                        account_result["masked_user_id"],
                        entry_result["apply_label"],
                    )
                else:
                    entry_result["status"] = "stopped_by_user"
                submission_entries.append(entry_result)

            account_result["submission_result"] = self._summarize_submission_entries(
                submission_entries
            )
            account_result["submitted"] = (
                account_result["submission_result"].get("submitted_count", 0) > 0
            )
            if submission_entries and all(
                entry.get("status") == "completed" for entry in submission_entries
            ):
                account_result["status"] = "completed"
            elif any(
                entry.get("status")
                in {
                    "submission_failed",
                    "selection_failed",
                    "slot_selection_not_applied",
                    "confirm_page_not_reached",
                    "apply_option_missing",
                    "confirm_mismatch",
                    "park_mismatch",
                    "facility_mismatch",
                    "confirm_park_or_facility_mismatch",
                    "submission_incomplete",
                }
                for entry in submission_entries
            ):
                account_result["status"] = "partial_error"
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
                ).strip()
                park_name = str(getattr(slot, "park_name", "") or "").strip()
                facility_name = str(getattr(slot, "facility_name", "") or "").strip()
                if (
                    searchable
                    and facility not in searchable
                    and searchable not in facility
                    and (not park_name or facility not in park_name)
                    and (not facility_name or facility not in facility_name)
                    and (not park_name or park_name not in facility)
                    and (not facility_name or facility_name not in facility)
                ):
                    continue
            return slot
        return None

    def _build_entry_result(self, account_result, slot, entry_index, account_index):
        facility = " ".join(
            part for part in (slot.park_name, slot.facility_name) if part
        ).strip()
        return {
            "account_index": account_index,
            "entry_index": entry_index,
            "account": account_result.get("masked_user_id"),
            "account_label": account_result.get("account_label", ""),
            "apply_no": f"{entry_index}-1",
            "apply_label": f"{entry_index}件目",
            "selected": False,
            "submitted": False,
            "status": "pending",
            "error_message": None,
            "confirmation_response": None,
            "ensure_result": {},
            "continue_after_completion": {
                "executed": entry_index > 1,
            },
            "search_after_continue": {
                "executed": entry_index > 1,
            },
            "ensure_skipped_before_select": False,
            "select_result": {},
            "confirm_page": {},
            "validation": {},
            "submit_result": {},
            "recaptcha_recovery": {},
            "submission_result": {},
            "slot": {
                "date": slot.date,
                "weekday": slot.weekday,
                "time_range": slot.time_range,
                "facility": facility,
                "current_entry_count": slot.applied_count,
                "park_name": slot.park_name,
                "facility_name": slot.facility_name,
                "raw_text": slot.raw_text,
            },
        }

    def _ensure_lottery_entry_page(self, driver, force_navigation=False):
        result = {
            "called": True,
            "forced": bool(force_navigation),
            "success": False,
            "display_no": "",
            "title": "",
            "current_url": "",
            "has_usedate_table": False,
        }
        try:
            page_state = self.navigation_service.inspect_page_state(driver)
            result.update(page_state)
            if (
                not force_navigation
                and "利用時間設定画面" in page_state.get("title", "")
                and page_state.get("display_no") == "plwba4000"
                and page_state.get("has_usedate_table")
            ):
                result["success"] = True
                return result
        except Exception:
            pass
        if not self.navigation_service.go_to_lottery_entry(driver):
            result["success"] = False
            result["error"] = "login_or_navigation_not_ready"
            return result
        self.navigation_service.select_lottery_tennis_park(driver)
        try:
            page_state = self.navigation_service.inspect_page_state(driver)
            result.update(page_state)
            result["success"] = bool(
                "利用時間設定画面" in page_state.get("title", "")
                and page_state.get("display_no") == "plwba4000"
                and page_state.get("has_usedate_table")
            )
        except Exception:
            pass
        return result

    def _ensure_lottery_tennis_park_selection(
        self,
        driver,
        account_index=None,
        entry_index=None,
        desired_park_value="1301270",
        desired_park_name="府中の森公園",
        desired_court_value="12700020",
        desired_court_name="テニス（人工芝）",
    ):
        state = {
            "checked": False,
            "reselected": False,
            "success": False,
            "reason": "",
            "current": {},
            "expected": {
                "bname_value": desired_park_value,
                "bname_text": desired_park_name,
                "iname_value": desired_court_value,
                "iname_text": desired_court_name,
            },
        }
        try:
            current = self.lottery_service.capture_lottery_selection_state(
                driver,
                "",
                selected=False,
                include_last_info=False,
            )
        except Exception:
            current = {}
        state["current"] = current
        state["checked"] = True
        current_bname = str(current.get("current_bname_value", "") or "").strip()
        current_bname_text = str(current.get("current_bname_text", "") or "").strip()
        current_iname = str(current.get("current_iname_value", "") or "").strip()
        current_iname_text = str(current.get("current_iname_text", "") or "").strip()
        current_bld_hidden = str(current.get("selectBldGrpCd", "") or "").strip()
        current_inst_hidden = str(current.get("selectInstGrpCd", "") or "").strip()
        expected_ok = (
            current_bname == desired_park_value
            and current_iname == desired_court_value
            and current_bld_hidden == desired_park_value
            and current_inst_hidden == desired_court_value
        )
        state["success"] = bool(expected_ok)
        state.update(
            {
                "current_bname_value": current_bname,
                "current_bname_text": current_bname_text,
                "current_iname_value": current_iname,
                "current_iname_text": current_iname_text,
                "selectBldGrpCd": current_bld_hidden,
                "selectInstGrpCd": current_inst_hidden,
            }
        )
        self.logger.info(
            "lottery park selection check account_index=%s entry_index=%s current_bname=%s current_bname_text=%s current_iname=%s current_iname_text=%s selectBldGrpCd=%s selectInstGrpCd=%s expected_ok=%s",
            account_index,
            entry_index,
            current_bname,
            current_bname_text,
            current_iname,
            current_iname_text,
            current_bld_hidden,
            current_inst_hidden,
            expected_ok,
        )
        return state

    def _refresh_slot_for_entry(self, driver, slot, target_weekdays, max_weeks):
        entry = {
            "date": slot.date,
            "time_range": slot.time_range,
            "facility": " ".join(
                part for part in (slot.park_name, slot.facility_name) if part
            ).strip(),
        }
        reselect_state = self._ensure_lottery_tennis_park_selection(driver)
        collect_result = self.slot_collector.collect_slots_for_entries(
            driver,
            target_entries=[entry],
            target_weekdays=target_weekdays,
            max_weeks=max_weeks,
        )
        matched_slot = self._match_slot(
            collect_result.get("slots", []),
            slot.date,
            slot.time_range,
            entry["facility"],
        )
        if matched_slot is None:
            same_date_slots = [
                self._serialize_slot(item)
                for item in collect_result.get("slots", [])
                if item.date == slot.date
            ]
            same_time_slots = [
                self._serialize_slot(item)
                for item in collect_result.get("slots", [])
                if item.date == slot.date and item.time_range == slot.time_range
            ]
            if matched_slot is None:
                self.logger.warning(
                    "refresh_slot_for_entry unmatched target date=%s time_range=%s facility=%s same_date_count=%s same_time_count=%s same_time_slots=%s",
                    slot.date,
                    slot.time_range,
                    entry["facility"],
                    len(same_date_slots),
                    len(same_time_slots),
                    json.dumps(same_time_slots, ensure_ascii=False),
                )
        return {
            "slot": matched_slot,
            "weeks_explored": collect_result.get("weeks_explored", 0),
            "stopped_reason": collect_result.get("stopped_reason"),
            "same_date_count": len(
                [
                    item
                    for item in collect_result.get("slots", [])
                    if item.date == slot.date
                ]
            ),
            "same_time_count": len(
                [
                    item
                    for item in collect_result.get("slots", [])
                    if item.date == slot.date and item.time_range == slot.time_range
                ]
            ),
            "park_selection_state": reselect_state,
            "facility_reselected": False,
            "retry_used": False,
        }

    def _summarize_submission_entries(self, entries):
        debug_files = []
        states = []
        recovery_attempts = 0
        recovery_triggered = False
        recovery_completed = False
        submitted_count = 0
        requested_count = len(entries)
        for entry in entries:
            submission_result = entry.get("submission_result", {}) or {}
            debug_files.extend(submission_result.get("debug_files", []))
            states.extend(submission_result.get("states", []))
            recovery_attempts += submission_result.get("recovery_attempts", 0)
            recovery_triggered = recovery_triggered or submission_result.get(
                "recovery_triggered", False
            )
            recovery_completed = recovery_completed or submission_result.get(
                "recovery_completed", False
            )
            submitted_count += submission_result.get("submitted_count", 0)
        return {
            "requested_count": requested_count,
            "submitted_count": submitted_count,
            "completed": requested_count > 0
            and all(entry.get("status") == "completed" for entry in entries),
            "recovery_triggered": recovery_triggered,
            "recovery_completed": recovery_completed,
            "recovery_attempts": recovery_attempts,
            "states": states,
            "debug_files": debug_files,
        }

    def print_result(self, result):
        print(f"Account source: {result.get('account_source')}")
        print(f"Target weekdays: {', '.join(result.get('target_weekdays', []))}")
        for account_result in result.get("accounts", []):
            account_part = account_result.get("masked_user_id")
            if account_result.get("account_label"):
                account_part = f"{account_part} ({account_result.get('account_label')})"
            print(f"Account: {account_part}")
            print(f"- status={account_result.get('status')}")
            print(
                "- session_mode={mode} reused={reused}".format(
                    mode=account_result.get("session_mode"),
                    reused=account_result.get("session_reused"),
                )
            )
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
                            date=self.lottery_service._slot_value(slot, "date", ""),
                            weekday=self.lottery_service._slot_value(slot, "weekday", ""),
                            time_range=self.lottery_service._slot_value(slot, "time_range", ""),
                            facility=" ".join(
                                part
                                for part in (
                                    self.lottery_service._slot_value(slot, "park_name", ""),
                                    self.lottery_service._slot_value(slot, "facility_name", ""),
                                )
                                if part
                            ).strip(),
                            applied=self.lottery_service._slot_value(slot, "applied_count", ""),
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

            entries = account_result.get("entries", [])
            if entries:
                print("- entries:")
                for entry in entries:
                    slot = entry.get("slot", {})
                    print(
                        "  {apply_label} {date} {weekday} {time_range} facility={facility} applied={applied} status={status}".format(
                            apply_label=entry.get("apply_label"),
                            date=self.lottery_service._slot_value(slot, "date", ""),
                            weekday=self.lottery_service._slot_value(slot, "weekday", ""),
                            time_range=self.lottery_service._slot_value(slot, "time_range", ""),
                            facility=self.lottery_service._slot_value(slot, "facility", ""),
                            applied=self.lottery_service._slot_value(slot, "current_entry_count", ""),
                            status=entry.get("status"),
                        )
                    )

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
