# -*- coding: utf-8 -*-
"""Open each account's reservation page for user-guided cancellation decisions."""

from __future__ import annotations

import logging
import re
import time
import csv
from pathlib import Path

from bs4 import BeautifulSoup as bs
from selenium.common.exceptions import UnexpectedAlertPresentException


class ReservationStatusWorkflowService:
    """Inspect reservation pages without automatically cancelling anything."""

    def __init__(
        self,
        reservation_service,
        login_service,
        browser_session,
        navigation_service,
        logger=None,
        sleep_func=time.sleep,
    ):
        self.reservation_service = reservation_service
        self.login_service = login_service
        self.browser_session = browser_session
        self.navigation_service = navigation_service
        self.logger = logger or logging.getLogger(__name__)
        self.sleep_func = sleep_func

    def run(self, accounts, decision_callback=None):
        """Open reservation pages one account at a time.

        ``decision_callback`` returns ``True`` when the user intends to cancel
        manually in the browser, and ``False`` to continue to the next account.
        This service never finds or clicks a cancellation control.
        """
        driver = self.browser_session.create_driver()
        results = {}
        try:
            for account in accounts:
                user_id = account["user_id"]
                result = {
                    "status": "unknown",
                    "account_label": account.get("account_label", ""),
                    "reservation_page": {},
                }
                try:
                    if not self.login_service.login(
                        driver, user_id, account["password"]
                    ):
                        result["status"] = "login_failed"
                        results[user_id] = result
                        continue

                    self.navigation_service.go_to_reservation_list(driver)
                    self.sleep_func(1)
                    result["reservation_page"] = self.inspect_page(driver)
                    result["status"] = "opened"
                    results[user_id] = result

                    # 予約がないIDではキャンセル確認を表示せず、次のIDへ進む。
                    if (
                        decision_callback is not None
                        and result["reservation_page"].get("reservation_count", 0) > 0
                    ):
                        should_cancel = bool(decision_callback(user_id, result))
                        result["cancel_decision"] = (
                            "manual_cancel" if should_cancel else "skip"
                        )
                    self._logout(driver)
                except UnexpectedAlertPresentException:
                    self._accept_alert(driver)
                    result["status"] = "no_reservation_or_alert"
                    results[user_id] = result
                    self._logout(driver)
                except Exception:
                    self.logger.exception("ID:%s reservation status check failed", user_id)
                    result["status"] = "error"
                    results[user_id] = result
                    self._logout(driver)
        finally:
            self.browser_session.safe_close(driver)
        return results

    def verify_accounts(self, accounts):
        """Re-open every reservation page and return the current reservation state."""
        driver = self.browser_session.create_driver()
        results = {}
        try:
            for account in accounts:
                user_id = account["user_id"]
                result = {
                    "status": "unknown",
                    "account_label": account.get("account_label", ""),
                    "reservation_page": {},
                }
                try:
                    if not self.login_service.login(
                        driver, user_id, account["password"]
                    ):
                        result["status"] = "login_failed"
                    else:
                        self.navigation_service.go_to_reservation_list(driver)
                        self.sleep_func(1)
                        page = self.inspect_page(driver)
                        result["reservation_page"] = page
                        result["reservation_count"] = page["reservation_count"]
                        result["status"] = "verified"
                    results[user_id] = result
                    self._logout(driver)
                except UnexpectedAlertPresentException:
                    self._accept_alert(driver)
                    result["status"] = "no_reservation_or_alert"
                    result["reservation_count"] = 0
                    results[user_id] = result
                    self._logout(driver)
                except Exception:
                    self.logger.exception(
                        "ID:%s reservation status verification failed", user_id
                    )
                    result["status"] = "error"
                    results[user_id] = result
                    self._logout(driver)
        finally:
            self.browser_session.safe_close(driver)
        return results

    @staticmethod
    def build_verification_results(results):
        """Convert the initial inspection results into CSV verification results."""
        verification = {}
        for user_id, result in results.items():
            status = result.get("status")
            page = result.get("reservation_page", {})
            verification_status = {
                "status": "verified" if status == "opened" else status,
                "account_label": result.get("account_label", ""),
                "reservation_page": page,
            }
            if status in {"opened", "no_reservation_or_alert"}:
                verification_status["reservation_count"] = page.get(
                    "reservation_count", 0
                )
            verification[user_id] = verification_status
        return verification

    @staticmethod
    def save_remaining_accounts_csv(
        input_csv_path,
        cancelled_ids,
        output_csv_path,
        reservation_details=None,
    ):
        """Write remaining accounts and append their current reservation dates/times."""
        cancelled = {str(user_id) for user_id in cancelled_ids}
        include_reservation_details = reservation_details is not None
        reservation_details = reservation_details or {}
        with open(input_csv_path, newline="", encoding="utf-8-sig") as source:
            rows = list(csv.reader(source))
        remaining = []
        for row in rows:
            if not row or row[0] in cancelled:
                continue
            reservations = reservation_details.get(row[0], [])
            reservation_text = " / ".join(
                " ".join(
                    part
                    for part in (reservation.get("date", ""), reservation.get("time", ""))
                    if part
                )
                for reservation in reservations
                if reservation.get("date") or reservation.get("time")
            )
            remaining.append(row + [reservation_text] if include_reservation_details else row)
        output_path = Path(output_csv_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", newline="", encoding="utf-8-sig") as target:
            csv.writer(target).writerows(remaining)
        return output_path

    def inspect_page(self, driver):
        """Return a sanitized summary while leaving the page open for the user."""
        soup = bs(driver.page_source, "html.parser")
        body_text = " ".join(soup.get_text(" ", strip=True).split())
        reservation_rows = soup.select("#rsvacceptlist tbody tr")
        tennis_rows = [row for row in reservation_rows if self._is_tennis_row(row)]
        reservations = [self._extract_reservation_details(row) for row in tennis_rows]
        controls = []
        for element in soup.find_all(["a", "button", "input"]):
            text = " ".join(element.get_text(" ", strip=True).split())
            value = (element.get("value") or "").strip()
            label = text or value
            if label and re.search("予約|キャンセル|取消", label):
                controls.append(
                    {
                        "tag": element.name,
                        "label": label,
                        "id": element.get("id", ""),
                        "name": element.get("name", ""),
                        "type": element.get("type", ""),
                    }
                )
        return {
            "title": driver.title,
            "url": driver.current_url,
            "reservation_count": len(tennis_rows),
            "all_reservation_count": len(reservation_rows),
            "reservations": reservations,
            "body_text": body_text[:4000],
            "controls": controls,
        }

    @staticmethod
    def _is_tennis_row(row):
        cells = row.find_all("td", recursive=False)
        facility_text = cells[3].get_text(" ", strip=True) if len(cells) > 3 else row.get_text(" ", strip=True)
        return "テニス" in facility_text

    @staticmethod
    def _extract_reservation_details(row):
        cells = row.find_all("td", recursive=False)
        cell_texts = [" ".join(cell.get_text(" ", strip=True).split()) for cell in cells]
        date = ""
        reservation_time = ""
        for text in cell_texts:
            if not date:
                match = re.search(
                    r"(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:\s*[（(][^）)]*[）)])?",
                    text,
                )
                if match:
                    date = match.group(0).strip()
            if not reservation_time:
                match = re.search(
                    r"\d{1,2}時\d{2}分\s*[～〜-]\s*\d{1,2}時\d{2}分",
                    text,
                )
                if match:
                    reservation_time = match.group(0).strip()
        return {"date": date, "time": reservation_time}

    def _accept_alert(self, driver):
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass

    def _logout(self, driver):
        try:
            self.navigation_service.logout(driver)
            self.sleep_func(0.5)
        except Exception:
            pass
