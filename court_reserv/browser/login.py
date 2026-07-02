# -*- coding: utf-8 -*-
"""Login helpers for legacy Selenium flows."""

import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    JavascriptException,
    TimeoutException,
    UnexpectedAlertPresentException,
)


class LoginService:
    """Encapsulate legacy login flow and manual captcha waiting."""

    def __init__(
        self,
        top_url,
        wait_factory,
        logger,
        show_info,
        ask_yes_no,
        sleep_func=time.sleep,
    ):
        self.top_url = top_url
        self.wait_factory = wait_factory
        self.logger = logger
        self.show_info = show_info
        self.ask_yes_no = ask_yes_no
        self.sleep_func = sleep_func

    def login(self, driver, user_id, password):
        """Run the legacy login flow and return True on success."""
        driver.get(self.top_url)
        try:
            try:
                driver.execute_script(
                    "javascript:doAction(document.form1, gRsvWTransUserLoginAction);"
                )
            except JavascriptException:
                pass

            try:
                user_el = driver.find_element(By.NAME, "userId")
            except Exception:
                try:
                    driver.switch_to.frame("pawae1002")
                    user_el = driver.find_element(By.NAME, "userId")
                except Exception:
                    self.logger.warning("Login form not found for user %s", user_id)
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
                    return False

            user_el.send_keys(user_id)
            driver.find_element(By.NAME, "password").send_keys(password)
            self.sleep_func(0.5)
            try:
                driver.execute_script(
                    "javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);"
                )
            except JavascriptException:
                try:
                    driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
                except Exception:
                    pass

            try:
                self.wait_factory(driver, 2).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                self.logger.warning("ID:%s login alert: %s", user_id, alert_text)
                return False
            except TimeoutException:
                try:
                    if self.detect_captcha(driver):
                        return bool(self.wait_for_manual_captcha(driver))
                except Exception:
                    pass
                return True
        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                self.logger.warning(
                    "ID:%s unexpected alert during login: %s",
                    user_id,
                    alert_text,
                )
            except Exception:
                pass
            return False

    def detect_captcha(self, driver):
        """Detect whether a reCAPTCHA-like widget appears on the page."""
        if not driver:
            return False
        try:
            els = driver.find_elements(
                By.CSS_SELECTOR, "iframe[src*='recaptcha'], div.g-recaptcha"
            )
            if els and len(els) > 0:
                return True
            page_source = driver.page_source
            if "g-recaptcha" in page_source or "recaptcha" in page_source:
                return True
        except Exception:
            return False
        return False

    def wait_for_manual_captcha(self, driver, timeout_minutes=5):
        """Wait for user-driven captcha completion without bypassing it."""
        try:
            self.show_info(
                "reCAPTCHA 検出",
                "ページ上にreCAPTCHAが検出されました。ブラウザで手動で認証してください。完了したら「OK」を押してください。",
            )
        except Exception:
            print(
                "reCAPTCHA detected: please solve it manually in the browser and then continue."
            )
        start = time.time()
        timeout = timeout_minutes * 60
        while True:
            try:
                if not self.detect_captcha(driver):
                    return True
            except Exception:
                return False
            if time.time() - start > timeout:
                try:
                    cont = self.ask_yes_no(
                        "reCAPTCHA まだ有効",
                        "reCAPTCHAがまだ残っています。続行して再試行しますか？(OK=再確認 / キャンセル=中断)",
                    )
                except Exception:
                    cont = False
                if not cont:
                    return False
                start = time.time()
            self.sleep_func(1)
