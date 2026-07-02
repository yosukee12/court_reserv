# -*- coding: utf-8 -*-
"""CLI-oriented lottery result workflow helpers."""

from __future__ import annotations

import csv
import json
import logging
import os
import re
from pathlib import Path

from bs4 import BeautifulSoup as bs

from court_reserv.config import get_default_credentials


class LotteryResultWorkflowService:
    """Drive lottery result retrieval without reservation confirmation."""

    def __init__(
        self,
        config,
        browser_session,
        login_service,
        navigation_service,
        lottery_service,
        id_manager_service,
        logger=None,
    ):
        self.config = config
        self.browser_session = browser_session
        self.login_service = login_service
        self.navigation_service = navigation_service
        self.lottery_service = lottery_service
        self.id_manager_service = id_manager_service
        self.logger = logger or logging.getLogger(__name__)

    def resolve_accounts(self, id_csv=None, account_id=None):
        """Resolve accounts by issue-defined priority."""
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

    def run(self, id_csv=None, account_id=None):
        accounts = self.resolve_accounts(id_csv=id_csv, account_id=account_id)
        result = {
            "status": "completed",
            "account_source": accounts[0]["source"] if accounts else None,
            "accounts_checked": len(accounts),
            "results": [],
            "summary": {"won": 0, "lost": 0, "unknown": 0},
            "account_summaries": [],
        }

        driver = self.browser_session.create_driver()
        try:
            for account in accounts:
                account_summary = self._fetch_account_results(driver, account)
                result["account_summaries"].append(account_summary)
                result["results"].extend(account_summary.get("results", []))
        finally:
            self.browser_session.safe_close(driver)

        for row in result["results"]:
            status = row.get("result", "unknown")
            if status not in result["summary"]:
                status = "unknown"
            result["summary"][status] += 1

        return result

    def _fetch_account_results(self, driver, account):
        summary = {
            "user_id": account["user_id"],
            "masked_user_id": self.mask_user_id(account["user_id"]),
            "account_label": account.get("account_label", ""),
            "status": "completed",
            "results": [],
            "error": None,
        }

        try:
            if not self.login_service.login(driver, account["user_id"], account["password"]):
                summary["status"] = "login_failed"
                summary["error"] = "login_failed"
                return summary

            self.navigation_service.go_to_lottery_result_list(driver)
            html = driver.page_source
            rows = self.parse_result_rows(
                html,
                masked_user_id=summary["masked_user_id"],
                account_label=summary["account_label"],
            )
            summary["results"] = rows
            if not rows:
                summary["status"] = "no_results"
        except Exception as exc:
            self.logger.exception(
                "Failed to fetch lottery results for %s", account["user_id"]
            )
            summary["status"] = "error"
            summary["error"] = str(exc)
        finally:
            try:
                self.navigation_service.logout(driver)
            except Exception:
                pass

        return summary

    def parse_result_rows(self, html, masked_user_id, account_label=""):
        """Parse lottery result rows from the current result page."""
        soup = bs(html, "html.parser")
        results = []
        seen = set()

        for row in soup.select("tr"):
            text = self._normalize_text(row.get_text(" ", strip=True))
            if not text:
                continue
            if not any(keyword in text for keyword in ("当選", "落選", "補欠")):
                continue

            parsed = self._build_result_record(text, masked_user_id, account_label)
            key = (
                parsed["account"],
                parsed["date"],
                parsed["time_range"],
                parsed["facility"],
                parsed["result"],
                parsed["raw_text"],
            )
            if key in seen:
                continue
            seen.add(key)
            results.append(parsed)

        if results:
            return results

        found_day_list = [
            self._normalize_text(elem.text)
            for elem in soup.find_all("span", string=re.compile("月.*日(.*)"))
        ]
        found_time_list = [
            self._normalize_text(elem.text)
            for elem in soup.find_all(string=re.compile("時.*分.*時.*分"))
        ]
        fallback_results = []
        for index, day_text in enumerate(found_day_list):
            time_text = found_time_list[index] if index < len(found_time_list) else ""
            raw_text = self._normalize_text(f"{day_text} {time_text}".strip())
            fallback_results.append(
                {
                    "account": masked_user_id,
                    "account_label": account_label,
                    "date": self._extract_date(raw_text),
                    "time_range": self._extract_time_range(raw_text),
                    "facility": "",
                    "result": "won",
                    "raw_text": raw_text,
                }
            )
        return fallback_results

    def _build_result_record(self, text, masked_user_id, account_label):
        date = self._extract_date(text)
        time_range = self._extract_time_range(text)
        result = self._classify_result(text)
        facility = self._extract_facility(text, date, time_range, result)
        return {
            "account": masked_user_id,
            "account_label": account_label,
            "date": date,
            "time_range": time_range,
            "facility": facility,
            "result": result,
            "raw_text": text,
        }

    def _extract_date(self, text):
        match = re.search(r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})", text)
        if match:
            return match.group(1).replace("/", "-")
        match = re.search(r"(\d{1,2}月\d{1,2}日)", text)
        if match:
            return match.group(1)
        return ""

    def _extract_time_range(self, text):
        match = re.search(r"(\d{1,2}:\d{2}\s*[～~-]\s*\d{1,2}:\d{2})", text)
        if match:
            return match.group(1).replace(" ", "")
        match = re.search(r"(\d{1,2}時\d{0,2}分\s*[～~-]\s*\d{1,2}時\d{0,2}分)", text)
        if match:
            return self._normalize_text(match.group(1))
        return ""

    def _extract_facility(self, text, date, time_range, result):
        facility = text
        for token in (date, time_range, "当選", "落選", "補欠", "不明"):
            if token:
                facility = facility.replace(token, " ")
        facility = re.sub(r"\s+", " ", facility).strip(" -:/")
        return facility

    def _classify_result(self, text):
        if "当選" in text:
            return "won"
        if "落選" in text:
            return "lost"
        return "unknown"

    def print_result(self, result):
        print(f"Account source: {result.get('account_source')}")
        print(f"Accounts checked: {result.get('accounts_checked', 0)}")
        print(
            "Summary: won={won} lost={lost} unknown={unknown}".format(
                won=result.get("summary", {}).get("won", 0),
                lost=result.get("summary", {}).get("lost", 0),
                unknown=result.get("summary", {}).get("unknown", 0),
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
                    f"- account={account_part} status={account_summary.get('status')}"
                )
                if account_summary.get("error"):
                    print(f"  error={account_summary.get('error')}")

        if not result.get("results"):
            print("No lottery result rows found.")
            return

        print("Lottery result rows:")
        for index, row in enumerate(result["results"], start=1):
            account_part = row.get("account")
            if row.get("account_label"):
                account_part = f"{account_part} ({row.get('account_label')})"
            print(
                f"{index}. account={account_part} result={row.get('result')} "
                f"date={row.get('date')} time={row.get('time_range')} "
                f"facility={row.get('facility')}"
            )

    def save_result(self, result, output_dir):
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        json_path = output_path / "lottery_result_workflow_result.json"
        csv_path = output_path / "lottery_result_workflow_result.csv"

        json_path.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        with open(csv_path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=[
                    "account",
                    "account_label",
                    "date",
                    "time_range",
                    "facility",
                    "result",
                    "raw_text",
                ],
            )
            writer.writeheader()
            for row in result.get("results", []):
                writer.writerow(row)

        return json_path, csv_path

    def mask_user_id(self, user_id):
        text = str(user_id)
        if len(text) <= 4:
            return "*" * len(text)
        return f"{text[:2]}***{text[-2:]}"

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
