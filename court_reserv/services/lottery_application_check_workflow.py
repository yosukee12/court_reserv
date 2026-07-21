# -*- coding: utf-8 -*-
"""Lottery application status workflow helpers."""

from __future__ import annotations

import logging
import os
import re
import time
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup as bs

from court_reserv.config import get_default_credentials, get_output_base_path


class LotteryApplicationCheckWorkflowService:
    """Fetch lottery application status rows and keep tennis entries only."""

    def __init__(
        self,
        config,
        browser_session,
        login_service,
        navigation_service,
        id_manager_service,
        logger=None,
        sleep_func=time.sleep,
    ):
        self.config = config
        self.browser_session = browser_session
        self.login_service = login_service
        self.navigation_service = navigation_service
        self.id_manager_service = id_manager_service
        self.logger = logger or logging.getLogger(__name__)
        self.sleep_func = sleep_func

    def resolve_accounts(self, id_csv=None, account_id=None):
        """Resolve accounts by ID CSV, config.local.ini, or environment."""
        if id_csv:
            id_dict = self.id_manager_service.load_accounts(id_csv)
            if not id_dict:
                raise ValueError(f"No accounts found in CSV: {id_csv}")

            if account_id:
                if account_id not in id_dict:
                    raise ValueError(f"Account ID not found in CSV: {account_id}")
                target_ids = [account_id]
            else:
                target_ids = list(id_dict.keys())

            accounts = []
            for user_id in target_ids:
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
                        "values": list(values),
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
                    "values": ["", "", config_password],
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
                    "values": ["", "", password],
                }
            ]

        raise ValueError(
            "Credentials were not found. Configure an ID CSV, config.local.ini, or .env."
        )

    def resolve_output_dir(self, id_csv=None, output_dir=None):
        if output_dir:
            return Path(output_dir)
        if id_csv:
            return Path(id_csv).resolve().parent
        return get_output_base_path() / "lottery_automation"

    def run(self, id_csv=None, account_id=None):
        accounts = self.resolve_accounts(id_csv=id_csv, account_id=account_id)
        result = {
            "status": "completed",
            "account_source": accounts[0]["source"] if accounts else None,
            "accounts_checked": len(accounts),
            "rows": [],
            "output_id_dict": {},
            "summary": {
                "accounts_with_rows": 0,
                "tennis_rows": 0,
                "excluded_baseball_rows": 0,
                "excluded_other_rows": 0,
            },
            "account_summaries": [],
        }

        driver = self.browser_session.create_driver() if accounts else None
        try:
            for account in accounts:
                account_summary = self._fetch_account_rows(driver, account)
                result["account_summaries"].append(account_summary)
                result["rows"].extend(account_summary.get("rows", []))
                if account_summary.get("rows"):
                    result["summary"]["accounts_with_rows"] += 1
        finally:
            self.browser_session.safe_close(driver)

        for row in result["rows"]:
            result["summary"]["tennis_rows"] += 1

        for account_summary in result["account_summaries"]:
            result["summary"]["excluded_baseball_rows"] += account_summary.get(
                "excluded_baseball_rows", 0
            )
            result["summary"]["excluded_other_rows"] += account_summary.get(
                "excluded_other_rows", 0
            )
            user_id = account_summary.get("user_id", "")
            if user_id:
                result["output_id_dict"][user_id] = account_summary.get(
                    "output_values", []
                )

        return result

    def _fetch_account_rows(self, driver, account):
        summary = {
            "user_id": account["user_id"],
            "masked_user_id": self.mask_user_id(account["user_id"]),
            "account_label": account.get("account_label", ""),
            "status": "completed",
            "rows": [],
            "excluded_baseball_rows": 0,
            "excluded_other_rows": 0,
            "output_values": self._build_output_values(account, []),
            "error": None,
        }

        try:
            if not self.login_service.login(driver, account["user_id"], account["password"]):
                summary["status"] = "login_failed"
                summary["error"] = "login_failed"
                return summary

            before_state = self._inspect_application_page_state(driver)
            self._log_page_state("before_cancel_list_navigation", before_state)
            self.navigation_service.go_to_lottery_cancel_list(driver)
            if not self._wait_for_lottery_application_page(driver):
                summary["status"] = "page_wait_failed"
                summary["error"] = "lottery_application_page_wait_failed"
                self._save_debug_html(
                    driver, account["user_id"], prefix="lottery_application_wait_failed"
                )
                return summary
            after_state = self._inspect_application_page_state(driver)
            self._log_page_state("after_cancel_list_navigation", after_state)

            html = driver.page_source
            rows, excluded_rows, parse_state = self.parse_application_rows(
                html,
                masked_user_id=summary["masked_user_id"],
                account_label=summary["account_label"],
            )
            summary["rows"] = rows[:2]
            summary["excluded_baseball_rows"] = excluded_rows.get("baseball", 0)
            summary["excluded_other_rows"] = excluded_rows.get("other", 0)
            summary["output_values"] = self._build_output_values(account, summary["rows"])
            self._log_parse_summary(parse_state, rows, excluded_rows)

            if not rows:
                summary["status"] = "no_tennis_rows"
                self._save_debug_html(driver, account["user_id"], prefix="lottery_application_empty")
        except Exception as exc:
            self.logger.exception(
                "Failed to fetch lottery application rows for %s", account["user_id"]
            )
            summary["status"] = "error"
            summary["error"] = str(exc)
            summary["output_values"] = self._build_output_values(account, [])
            self._save_debug_html(driver, account["user_id"], prefix="lottery_application_error")
        finally:
            try:
                self.navigation_service.logout(driver)
            except Exception:
                pass

        return summary

    def parse_application_rows(self, html, masked_user_id, account_label=""):
        """Parse application rows and keep tennis only."""
        soup = bs(html, "html.parser")
        tables = soup.find_all("table")
        target_table = self._find_application_table(tables)
        row_elements = target_table.find_all("tr") if target_table else soup.find_all("tr")

        bs_day_matches = [
            self._normalize_text(elem)
            for elem in soup.find_all(string=re.compile("月.*日(.*)"))
        ]
        bs_time_matches = [
            self._normalize_text(elem)
            for elem in soup.find_all(string=re.compile("時.*分"))
        ]

        rows = []
        excluded_rows = {"baseball": 0, "other": 0}
        seen = set()
        candidate_row_count = 0
        for index, row in enumerate(row_elements, start=1):
            text = self._normalize_text(row.get_text(" ", strip=True))
            if not text or not any(
                token in text
                for token in ("申込み番号：", "状態：", "分類：", "公園・施設：", "利用日：", "時刻：")
            ):
                continue
            candidate_row_count += 1
            parsed, exclusion_reason = self._parse_application_row(
                row,
                masked_user_id=masked_user_id,
                account_label=account_label,
                index=index,
            )
            if parsed is None:
                if exclusion_reason in excluded_rows:
                    excluded_rows[exclusion_reason] += 1
                else:
                    excluded_rows["other"] += 1
                continue

            key = (
                parsed["account"],
                parsed["application_no"],
                parsed["date"],
                parsed["time_range"],
                parsed["park_facility"],
                parsed["sport"],
                parsed["status"],
            )
            if key in seen:
                continue
            seen.add(key)
            rows.append(parsed)

        state = {
            "table_exists": bool(target_table),
            "table_count": len(tables),
            "row_count": candidate_row_count,
            "bs_day_count": len(bs_day_matches),
            "bs_time_count": len(bs_time_matches),
            "candidate_count": candidate_row_count
            + excluded_rows.get("baseball", 0)
            + excluded_rows.get("other", 0),
        }
        return rows, excluded_rows, state

    def _find_application_table(self, tables):
        for table in tables:
            headers = [
                self._normalize_text(th.get_text(" ", strip=True))
                for th in table.find_all("th")
            ]
            if {"申込み", "状態", "分類", "公園・施設", "利用日", "時刻"}.issubset(
                set(headers)
            ):
                return table
        return None

    def _parse_application_row(self, row, masked_user_id, account_label, index):
        text = self._normalize_text(row.get_text(" ", strip=True))
        if not text:
            return None, "other"

        application_no = self._extract_number(text, r"申込み番号：\s*(\d+)")
        status = self._extract_text_after_label(text, "状態：")
        sport = self._extract_text_after_label(text, "分類：")
        park_facility = self._extract_text_after_label(text, "公園・施設：")
        date = self._extract_date(text)
        time_range = self._extract_time_range(text)
        cancel_label = "取消" if "取消" in text else ""

        if "野球" in text:
            self.logger.info(
                "Skipping baseball lottery application row account=%s index=%s text=%s",
                masked_user_id,
                index,
                text,
            )
            return None, "baseball"
        if "テニス" not in text:
            self.logger.info(
                "Skipping non-tennis lottery application row account=%s index=%s text=%s",
                masked_user_id,
                index,
                text,
            )
            return None, "other"

        return (
            {
                "account": masked_user_id,
                "account_label": account_label,
                "application_no": application_no,
                "status": status,
                "sport": sport,
                "park_facility": park_facility,
                "date": date,
                "time_range": time_range,
                "cancel_label": cancel_label,
                "raw_text": text,
            },
            None,
        )

    def print_result(self, result):
        print(f"Account source: {result.get('account_source')}")
        print(f"Accounts checked: {result.get('accounts_checked', 0)}")
        print(
            "Summary: accounts_with_rows={accounts_with_rows} tennis_rows={tennis_rows} excluded_baseball_rows={excluded_baseball_rows} excluded_other_rows={excluded_other_rows}".format(
                accounts_with_rows=result.get("summary", {}).get(
                    "accounts_with_rows", 0
                ),
                tennis_rows=result.get("summary", {}).get("tennis_rows", 0),
                excluded_baseball_rows=result.get("summary", {}).get(
                    "excluded_baseball_rows", 0
                ),
                excluded_other_rows=result.get("summary", {}).get(
                    "excluded_other_rows", 0
                ),
            )
        )
        account_summaries = result.get("account_summaries", [])
        if account_summaries:
            print("Account summaries:")
            for account_summary in account_summaries:
                account_part = account_summary.get("masked_user_id")
                if account_summary.get("account_label"):
                    account_part = (
                        f"{account_part} ({account_summary.get('account_label')})"
                    )
                print(
                    f"- account={account_part} status={account_summary.get('status')} "
                    f"rows={len(account_summary.get('rows', []))} "
                    f"excluded_baseball={account_summary.get('excluded_baseball_rows', 0)} "
                    f"excluded_other={account_summary.get('excluded_other_rows', 0)}"
                )
                if account_summary.get("error"):
                    print(f"  error={account_summary.get('error')}")

        if not result.get("rows"):
            print("No tennis lottery application rows found.")
            return

        print("Lottery application rows:")
        for index, row in enumerate(result["rows"], start=1):
            account_part = row.get("account")
            if row.get("account_label"):
                account_part = f"{account_part} ({row.get('account_label')})"
            print(
                f"{index}. account={account_part} no={row.get('application_no')} "
                f"status={row.get('status')} sport={row.get('sport')} "
                f"date={row.get('date')} time={row.get('time_range')} "
                f"facility={row.get('park_facility')}"
            )

    def save_result(self, result, output_dir=None, id_csv=None):
        output_path = self.resolve_output_dir(id_csv=id_csv, output_dir=output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = output_path / f"check_lottery_status_{timestamp}.csv"
        output_id_dict = result.get("output_id_dict", {})
        if output_id_dict:
            self.id_manager_service.save_accounts(output_id_dict, csv_path)
        else:
            csv_path.write_text("", encoding="utf-8")
        return csv_path

    def mask_user_id(self, user_id):
        text = str(user_id)
        if len(text) <= 4:
            return "*" * len(text)
        return f"{text[:2]}***{text[-2:]}"

    def _build_output_values(self, account, rows):
        values = list(account.get("values", []))
        if len(values) < 3:
            values.extend([""] * (3 - len(values)))
        values = values[:3]
        for row in rows[:2]:
            values.append(self._format_application_value(row))
        while len(values) < 5:
            values.append("")
        return values[:5]

    def _format_application_value(self, row):
        date = self._normalize_text(row.get("date", ""))
        time_range = self._normalize_text(row.get("time_range", ""))
        if date and time_range:
            return f"{date} {time_range}"
        return date or time_range

    def _inspect_application_page_state(self, driver):
        try:
            return driver.execute_script(
                """
                const displayNo = document.querySelector('input[name="displayNo"]');
                const tables = document.querySelectorAll("table");
                const rows = document.querySelectorAll("table tr");
                return {
                  current_url: window.location.href || "",
                  title: document.title || "",
                  display_no: displayNo ? (displayNo.value || "") : "",
                  has_form1: !!document.form1,
                  has_doAction: typeof doAction === "function",
                  table_exists: tables.length > 0,
                  table_count: tables.length,
                  row_count: rows.length,
                  ready_state: document.readyState || "",
                };
                """
            )
        except Exception as exc:
            return {"error": str(exc)}

    def _log_page_state(self, prefix, state):
        state = state if isinstance(state, dict) else {}
        self.logger.info(
            "%s current_url=%s title=%s displayNo=%s table_exists=%s table_count=%s row_count=%s ready_state=%s",
            prefix,
            state.get("current_url", ""),
            state.get("title", ""),
            state.get("display_no", ""),
            state.get("table_exists", False),
            state.get("table_count", 0),
            state.get("row_count", 0),
            state.get("ready_state", ""),
        )

    def _log_parse_summary(self, state, rows, excluded_rows):
        state = state if isinstance(state, dict) else {}
        self.logger.info(
            "parse_summary table_exists=%s row_count=%s bs_day_count=%s bs_time_count=%s extracted_after_count=%s excluded_after_count=%s",
            state.get("table_exists", False),
            state.get("row_count", 0),
            state.get("bs_day_count", 0),
            state.get("bs_time_count", 0),
            len(rows),
            excluded_rows.get("baseball", 0) + excluded_rows.get("other", 0),
        )

    def _save_debug_html(self, driver, user_id, prefix="lottery_application_empty"):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_user_id = re.sub(r"[^0-9A-Za-z_-]+", "_", str(user_id) or "unknown")
            filename = f"{prefix}_{safe_user_id}_{timestamp}.html"
            saver = getattr(self.navigation_service, "save_debug_html", None)
            if callable(saver):
                return saver(driver, filename)
        except Exception:
            self.logger.exception("Failed to save debug HTML for lottery application check")
        return None

    def _extract_text_after_label(self, text, label):
        match = re.search(
            re.escape(label)
            + r"\s*(.*?)(?=(?:申込み番号：|状態：|分類：|公園・施設：|利用日：|時刻：|取消|$))",
            text,
        )
        return self._normalize_text(match.group(1) if match else "")

    def _extract_number(self, text, pattern):
        match = re.search(pattern, text)
        return match.group(1) if match else ""

    def _extract_date(self, text):
        match = re.search(r"(\d{1,2}月\d{1,2}日(?:\([^)]+\))?)", text)
        return match.group(1) if match else ""

    def _extract_time_range(self, text):
        match = re.search(
            r"(\d{1,2}時\d{0,2}分\s*[～~-]\s*\d{1,2}時\d{0,2}分)",
            text,
        )
        return self._normalize_text(match.group(1)) if match else ""

    def _wait_for_lottery_application_page(self, driver, timeout=10):
        try:
            return bool(
                self.browser_session.get_wait(driver, timeout).until(
                    lambda d: "抽選申込みの確認" in getattr(d, "title", "")
                    or "抽選申込の確認・取消画面" in getattr(d, "title", "")
                    or "gLotWTransLotCancelListAction"
                    in (getattr(d, "page_source", "") or "")
                )
            )
        except Exception:
            return False

    def _normalize_text(self, text):
        return re.sub(r"\s+", " ", text or "").strip()

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
