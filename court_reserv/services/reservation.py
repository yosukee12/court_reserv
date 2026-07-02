# -*- coding: utf-8 -*-
"""Reservation business logic extracted from the legacy UI class."""

import re
import time

from bs4 import BeautifulSoup as bs
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException


class ReservationService:
    """Encapsulate reservation confirmation and reservation-check flows."""

    def __init__(
        self,
        config,
        browser_session,
        navigation_service,
        logger,
        get_id_dict_from_csv,
        output_id_dict,
        sleep_func=time.sleep,
    ):
        self.config = config
        self.browser_session = browser_session
        self.navigation_service = navigation_service
        self.logger = logger
        self.get_id_dict_from_csv = get_id_dict_from_csv
        self.output_id_dict = output_id_dict
        self.sleep_func = sleep_func

    def _get_wait(self, driver, timeout=10):
        return self.browser_session.get_wait(driver, timeout)

    def _legacy_login(self, driver, user_id, password):
        driver.get(self.config["URL"]["TOP_URL"])
        self.navigation_service.do_action(driver, "gRsvWTransUserLoginAction")
        driver.find_element(By.NAME, "userId").send_keys(user_id)
        driver.find_element(By.NAME, "password").send_keys(password)
        self.sleep_func(0.5)
        self.navigation_service.execute_script(
            driver,
            "javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);",
        )

    def confirm_accounts(self, accounts, login_service=None):
        """Confirm reservations for selected accounts using the legacy flow."""
        result_dict = {}
        driver = self.browser_session.create_driver()
        try:
            for account in accounts:
                user_id = account["user_id"]
                password = account["password"]
                display_name = account.get("account_label", "")
                user_values = [display_name, "", password]
                try:
                    if login_service is not None:
                        if not login_service.login(driver, user_id, password):
                            self.logger.warning("ID:%s login failed", user_id)
                            result_dict[user_id] = {
                                "status": "login_failed",
                                "confirmed": [],
                            }
                            continue
                    else:
                        self._legacy_login(driver, user_id, password)
                except UnexpectedAlertPresentException:
                    print("ID:" + user_id + " 期限切れ")
                    self.logger.warning("ID:" + user_id + " 期限切れ")
                    result_dict[user_id] = {
                        "status": "login_failed",
                        "confirmed": [],
                    }
                    continue

                result_entry = {"status": "completed", "confirmed": []}
                if "ホーム画面" in driver.title:
                    try:
                        self.navigation_service.go_to_lottery_result_list(driver)
                        self.sleep_func(0.5)
                        soup = bs(driver.page_source, "html.parser")
                        found_day_list = [
                            elem.text
                            for elem in soup.find_all(
                                "span", string=re.compile("月.*日(.*)")
                            )
                        ]
                        found_time_list = [
                            elem.text
                            for elem in soup.find_all(
                                string=re.compile("時.*分～.*時.*分")
                            )
                        ]
                        if len(found_day_list) == 1:
                            self._get_wait(driver, 240).until(
                                EC.alert_is_present(),
                                "Timed out waiting for PA creation confirmation popup to appear.",
                            )
                            alert = driver.switch_to.alert
                            alert.accept()
                            confirmed_label = found_day_list[0] + " " + found_time_list[0]
                            print("ID:" + user_id + " 確定日→ " + confirmed_label)
                            result_entry["confirmed"].append(confirmed_label)
                            self.logger.info(
                                "ID:%s 予約確定完了→ %s", user_id, confirmed_label
                            )
                        elif len(found_day_list) == 2:
                            for i in range(2):
                                self._get_wait(driver, 240).until(
                                    EC.alert_is_present(),
                                    "Timed out waiting for PA creation confirmation popup to appear.",
                                )
                                alert = driver.switch_to.alert
                                alert.accept()
                                confirmed_label = (
                                    found_day_list[i] + " " + found_time_list[i]
                                )
                                print("ID:" + user_id + " 確定日→ " + confirmed_label)
                                result_entry["confirmed"].append(confirmed_label)
                                self.logger.info(
                                    "ID:%s 予約確定完了→ %s",
                                    user_id,
                                    confirmed_label,
                                )
                        else:
                            result_entry["status"] = "no_won_results"
                    except UnexpectedAlertPresentException:
                        print("ID:" + user_id + " 申込みなし")
                        result_entry["status"] = "no_won_results"
                    except Exception:
                        self.logger.exception("ID:%s reservation confirmation failed", user_id)
                        result_entry["status"] = "error"

                result_dict[user_id] = result_entry
                self.sleep_func(1)
                try:
                    self.navigation_service.logout(driver)
                except Exception:
                    pass
                self.sleep_func(1)
        finally:
            self.browser_session.safe_close(driver)
        return result_dict

    def determine_reserv(self, input_csv_path="", output_csv_path=""):
        """Run the legacy reservation confirmation flow."""
        print(input_csv_path)
        id_dict = self.get_id_dict_from_csv(input_csv_path)
        result_dict = {}
        driver = self.browser_session.create_driver()
        try:
            for user_id, user_values in id_dict.items():
                try:
                    self._legacy_login(driver, user_id, user_values[2])
                except UnexpectedAlertPresentException:
                    print("ID:" + user_id + " 期限切れ")
                    self.logger.warning("ID:" + user_id + " 期限切れ")
                    continue

                if "ホーム画面" in driver.title:
                    try:
                        self.navigation_service.go_to_lottery_result_list(driver)
                        self.sleep_func(0.5)
                        soup = bs(driver.page_source, "html.parser")
                        found_day_list = [
                            elem.text
                            for elem in soup.find_all(
                                "span", string=re.compile("月.*日(.*)")
                            )
                        ]
                        found_time_list = [
                            elem.text
                            for elem in soup.find_all(
                                string=re.compile("時.*分～.*時.*分")
                            )
                        ]
                        if len(found_day_list) == 1:
                            self._get_wait(driver, 240).until(
                                EC.alert_is_present(),
                                "Timed out waiting for PA creation confirmation popup to appear.",
                            )
                            alert = driver.switch_to.alert
                            alert.accept()
                            print(
                                "ID:"
                                + user_id
                                + " 確定日→ "
                                + found_day_list[0]
                                + " "
                                + found_time_list[0]
                            )
                            result_dict[user_id] = [
                                user_values[0],
                                user_values[1],
                                user_values[2],
                                found_day_list[0] + " " + found_time_list[0],
                            ]
                            self.logger.info(
                                "ID:"
                                + user_id
                                + " 予約確定完了→ "
                                + found_day_list[0]
                                + " "
                                + found_time_list[0]
                            )
                        elif len(found_day_list) == 2:
                            for i in range(2):
                                self._get_wait(driver, 240).until(
                                    EC.alert_is_present(),
                                    "Timed out waiting for PA creation confirmation popup to appear.",
                                )
                                alert = driver.switch_to.alert
                                alert.accept()
                                if i == 0:
                                    print(
                                        "ID:"
                                        + user_id
                                        + " 確定日→ "
                                        + found_day_list[0]
                                        + " "
                                        + found_time_list[0]
                                    )
                                    self.logger.info(
                                        "ID:"
                                        + user_id
                                        + " 予約確定完了→ "
                                        + found_day_list[0]
                                        + " "
                                        + found_time_list[0]
                                    )
                                elif i == 1:
                                    print(
                                        "ID:"
                                        + user_id
                                        + " 確定日→ "
                                        + found_day_list[1]
                                        + " "
                                        + found_time_list[1]
                                    )
                                    result_dict[user_id] = [
                                        user_values[0],
                                        user_values[1],
                                        user_values[2],
                                        found_day_list[0] + " " + found_time_list[0],
                                        found_day_list[1] + " " + found_time_list[1],
                                    ]
                                    self.logger.info(
                                        "ID:"
                                        + user_id
                                        + " 予約確定完了→ "
                                        + found_day_list[1]
                                        + " "
                                        + found_time_list[1]
                                    )
                    except UnexpectedAlertPresentException:
                        print("ID:" + user_id + " 申込みなし")
                        result_dict[user_id] = [
                            user_values[0],
                            user_values[1],
                            user_values[2],
                            "",
                            "",
                        ]
                        continue

                self.sleep_func(1)
                self.navigation_service.logout(driver)
                self.sleep_func(1)
        finally:
            self.browser_session.safe_close(driver)

        if output_csv_path != "":
            self.output_id_dict(result_dict, output_csv_path)

        return result_dict

    def check_reserv(self, id_dict, output_csv_path=""):
        """Run the legacy reservation-check flow."""
        result_dict = {}
        driver = self.browser_session.create_driver()
        try:
            for user_id, user_values in id_dict.items():
                try:
                    self._legacy_login(driver, user_id, user_values[2])
                except UnexpectedAlertPresentException:
                    print("ID:" + user_id + " 期限切れ")
                    self.logger.warning("ID:" + user_id + " 期限切れ")
                    continue

                if "ホーム画面" in driver.title:
                    try:
                        self.navigation_service.go_to_reservation_list(driver)
                        self.sleep_func(3)
                    except UnexpectedAlertPresentException:
                        print("ID:" + user_id + " 申込みなし")
                        result_dict[user_id] = [
                            user_values[0],
                            user_values[1],
                            user_values[2],
                            "",
                            "",
                        ]
                        continue
                self.navigation_service.logout(driver)
                self.sleep_func(1)
        finally:
            self.browser_session.safe_close(driver)

        if output_csv_path != "":
            self.output_id_dict(result_dict, output_csv_path)

        return result_dict
