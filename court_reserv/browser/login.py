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
        self.last_login_state = {}
        self.last_login_error = None

    def login(self, driver, user_id, password):
        """Run the legacy login flow and return True on success."""
        self.last_login_state = {}
        self.last_login_error = None
        driver.get(self.top_url)
        try:
            if not self._wait_for_top_page_ready(driver):
                state = self.inspect_page_state(driver)
                if self._should_retry_login_page(state):
                    if not self._retry_login_page_until_ready(driver, user_id):
                        self.last_login_error = "login_page_not_ready"
                        self.last_login_state = self.inspect_page_state(driver)
                        self.logger.warning(
                            "Login page not ready for user %s: %s",
                            user_id,
                            self.last_login_state,
                        )
                        return False
                else:
                    self.last_login_error = "login_page_not_ready"
                    self.last_login_state = state
                    self.logger.warning(
                        "Login page not ready for user %s: %s",
                        user_id,
                        state,
                    )
                    return False
            before_state = self.inspect_page_state(driver)
            self.last_login_state = {"before_login": before_state}

            user_el = None
            password_el = None
            for attempt in range(3):
                try:
                    if before_state.get("has_doAction") and before_state.get("has_login_action"):
                        driver.execute_script(
                            "javascript:doAction(document.form1, gRsvWTransUserLoginAction);"
                        )
                except JavascriptException:
                    pass
                try:
                    user_el, password_el = self._wait_for_login_form(driver, timeout=5)
                    break
                except Exception:
                    self.logger.info(
                        "Login form not ready yet for user %s attempt=%s state=%s",
                        user_id,
                        attempt + 1,
                        self.inspect_page_state(driver),
                    )
                    self.sleep_func(0.5)
                    before_state = self.inspect_page_state(driver)
            if user_el is None or password_el is None:
                self.last_login_error = (
                    "login_form_not_found"
                    if not before_state.get("has_userId") and not before_state.get("has_password")
                    else "login_form_not_ready"
                )
                self.last_login_state["login_form_state"] = self.inspect_page_state(driver)
                self.logger.warning("Login form not found for user %s", user_id)
                return False

            form_state = self.inspect_page_state(driver)
            self.last_login_state["login_form_state"] = form_state
            self.logger.info("Login form ready: %s", form_state)

            try:
                user_el.clear()
            except Exception:
                pass
            try:
                password_el.clear()
            except Exception:
                pass
            user_el.send_keys(user_id)
            password_el.send_keys(password)
            self.sleep_func(0.5)
            self._submit_login(driver, password_el)
            self.logger.info("Login submitted: %s", self.inspect_page_state(driver))

            try:
                self.wait_factory(driver, 2).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                self.logger.warning("ID:%s login alert: %s", user_id, alert_text)
                self.last_login_error = "login_alert"
                self.last_login_state["after_login_submit"] = self.inspect_page_state(driver)
                return False
            except TimeoutException:
                current_state = self.inspect_page_state(driver)
                if current_state.get("has_login_form"):
                    self._retry_login_submit(driver, password_el)
                    self.logger.info(
                        "Login re-submitted after login form remained: %s",
                        self.inspect_page_state(driver),
                    )
                try:
                    if self.detect_captcha(driver):
                        return bool(self.wait_for_manual_captcha(driver))
                except Exception:
                    pass
                if self._wait_for_post_login_ready(driver):
                    post_state = self.inspect_page_state(driver)
                    self.last_login_state["after_login_submit"] = post_state
                    self.logger.info("Login post page ready: %s", post_state)
                    return True
                self.last_login_error = "login_post_page_not_ready"
                self.last_login_state["after_login_submit"] = self.inspect_page_state(driver)
                return False
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
            self.last_login_error = "unexpected_login_alert"
            return False

    def _should_retry_login_page(self, state):
        title = state.get("title", "")
        url = state.get("url", "")
        has_login_form = bool(state.get("has_login_form", False))
        has_user_id = bool(state.get("has_userId", False))
        has_password = bool(state.get("has_password", False))
        return (
            title == "施設予約システムからのお知らせ"
            or (
                "/web/index.jsp" in url
                and not has_login_form
                and not has_user_id
                and not has_password
            )
        )

    def _retry_login_page_until_ready(self, driver, user_id, max_retries=10):
        for retry_count in range(1, max_retries + 1):
            state = self.inspect_page_state(driver)
            self.logger.warning(
                "login notice page detected retry_count=%s title=%s url=%s",
                retry_count,
                state.get("title", ""),
                state.get("url", ""),
            )
            try:
                if retry_count % 2 == 1:
                    driver.refresh()
                else:
                    driver.get(self._build_login_retry_url(driver))
            except Exception:
                try:
                    driver.get(self._build_login_retry_url(driver))
                except Exception:
                    pass
            self.sleep_func(1.5)
            if self._wait_for_top_page_ready(driver):
                ready_state = self.inspect_page_state(driver)
                self.logger.info(
                    "final login page ready retry_count=%s title=%s url=%s",
                    retry_count,
                    ready_state.get("title", ""),
                    ready_state.get("url", ""),
                )
                return True
        final_state = self.inspect_page_state(driver)
        self.logger.warning(
            "final login page failed retry_count=%s title=%s url=%s",
            max_retries,
            final_state.get("title", ""),
            final_state.get("url", ""),
        )
        return False

    def _build_login_retry_url(self, driver):
        candidate = getattr(driver, "current_url", "") or self.top_url or ""
        for source in (candidate, self.top_url or ""):
            if not source:
                continue
            if "rsvWTransUserLoginAction.do" in source:
                return source
            if "/web/index.jsp" in source:
                return source.split("/web/index.jsp", 1)[0] + "/web/rsvWTransUserLoginAction.do"
        return self.top_url

    def _submit_login(self, driver, password_el):
        try:
            if driver.execute_script("return typeof submitLogin === 'function';"):
                driver.execute_script(
                    "javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);"
                )
                return
        except JavascriptException:
            pass
        except Exception:
            pass
        try:
            password_el.send_keys(Keys.RETURN)
            return
        except Exception:
            pass
        self._click_login_submit(driver)

    def _retry_login_submit(self, driver, password_el):
        self.sleep_func(0.5)
        try:
            password_el.send_keys(Keys.RETURN)
            return
        except Exception:
            pass
        self._click_login_submit(driver)

    def _click_login_submit(self, driver):
        for selector in (
            "button[type='submit']",
            "input[type='submit']",
            "#btn-go",
        ):
            try:
                element = driver.find_element(By.CSS_SELECTOR, selector)
                element.click()
                return True
            except Exception:
                continue
        try:
            driver.execute_script("document.form1.submit();")
            return True
        except Exception:
            return False

    def inspect_page_state(self, driver):
        try:
            return driver.execute_script(
                """
                const hasLoginInputs = !!document.querySelector('input[name="userId"]') && !!document.querySelector('input[name="password"]');
                const hasUserMenu = !!document.querySelector('a[href*="UserAttestationEndAction"], a[href*="logout"], .user-info, #btn-barcode');
                return {
                  readyState: document.readyState || "",
                  title: document.title || "",
                  url: window.location.href,
                  has_form1: !!document.form1,
                  has_doAction: typeof doAction === "function",
                  has_login_action: typeof gRsvWTransUserLoginAction !== "undefined",
                  has_lottery_action: typeof gLotWOpeLotSearchAction !== "undefined",
                  has_userId: !!document.querySelector('input[name="userId"]'),
                  has_password: !!document.querySelector('input[name="password"]'),
                  has_login_form: hasLoginInputs,
                  has_user_menu: hasUserMenu
                };
                """
            )
        except Exception:
            return {
                "readyState": "",
                "title": getattr(driver, "title", ""),
                "url": getattr(driver, "current_url", ""),
                "has_form1": False,
                "has_doAction": False,
                "has_login_action": False,
                "has_lottery_action": False,
                "has_userId": False,
                "has_password": False,
                "has_login_form": False,
                "has_user_menu": False,
            }

    def _wait_for_top_page_ready(self, driver, retries=3):
        for _ in range(max(retries, 1)):
            state = self.inspect_page_state(driver)
            if state.get("readyState") == "complete" and (
                (state.get("has_form1") and state.get("has_doAction"))
                or state.get("has_login_form")
            ):
                return True
            self.sleep_func(0.5)
        return False

    def _wait_for_login_form(self, driver, timeout=5):
        self.wait_factory(driver, timeout).until(
            lambda d: self.inspect_page_state(d).get("has_login_form")
        )
        return (
            driver.find_element(By.NAME, "userId"),
            driver.find_element(By.NAME, "password"),
        )

    def _wait_for_post_login_ready(self, driver, timeout=10):
        try:
            self.wait_factory(driver, timeout).until(
                lambda d: self._is_post_login_ready(d)
            )
            return True
        except Exception:
            return False

    def _is_post_login_ready(self, driver):
        state = self.inspect_page_state(driver)
        return (
            state.get("readyState") == "complete"
            and state.get("has_form1")
            and state.get("has_doAction")
            and state.get("has_lottery_action")
            and not state.get("has_login_form")
            and (
                state.get("has_user_menu")
                or "ログイン" not in state.get("title", "")
            )
        )

    def detect_captcha(self, driver):
        """Detect whether a visible manual reCAPTCHA widget appears on the page."""
        if not driver:
            return False
        try:
            state = self.inspect_recaptcha_state(driver)
            if state and state.get("requires_manual"):
                try:
                    self.logger.info("Visible reCAPTCHA detected: %s", state)
                except Exception:
                    pass
                return True
        except Exception:
            return False
        return False

    def inspect_recaptcha_state(self, driver):
        if not driver:
            return {
                "readyState": "",
                "title": "",
                "url": "",
                "visible_matches": [],
                "requires_manual": False,
                "response_fields": [],
                "response_length": 0,
                "response_present": False,
            }
        try:
            return driver.execute_script(
                """
                const isVisible = (el) => {
                  if (!el) return false;
                  const style = window.getComputedStyle(el);
                  const rect = el.getBoundingClientRect();
                  return style.display !== 'none'
                    && style.visibility !== 'hidden'
                  && rect.width > 0
                    && rect.height > 0;
                };
                const responseFields = Array.from(document.querySelectorAll(
                  "textarea[name='g-recaptcha-response'], textarea[id^='g-recaptcha-response'], textarea#recaptchaToken"
                )).map((el) => ({
                  id: el.id || '',
                  name: el.name || '',
                  value_length: (el.value || '').length,
                  value_present: !!(el.value || '').trim(),
                }));
                const selectors = [
                  "iframe[src*='recaptcha'][title*='reCAPTCHA']",
                  "iframe[src*='recaptcha'][title*='challenge']",
                  "iframe[src*='recaptcha'][src*='/bframe']",
                  "div.g-recaptcha",
                  "iframe[title*='セキュリティ']",
                ];
                const visibleMatches = [];
                for (const selector of selectors) {
                  for (const el of Array.from(document.querySelectorAll(selector))) {
                    const src = el.getAttribute('src') || '';
                    const dataSize = el.getAttribute('data-size') || '';
                    const isInvisibleAnchor =
                      src.indexOf('/anchor') >= 0 && src.indexOf('size=invisible') >= 0;
                    const isInvisibleWidget =
                      el.classList.contains('g-recaptcha') && dataSize === 'invisible';
                    const isNormalCheckbox =
                      src.indexOf('/anchor') >= 0 && src.indexOf('size=normal') >= 0;
                    const isNormalWidget =
                      el.classList.contains('g-recaptcha') && dataSize !== 'invisible';
                    const isChallenge =
                      src.indexOf('/bframe') >= 0 ||
                      (el.getAttribute('title') || '').toLowerCase().indexOf('challenge') >= 0 ||
                      (el.getAttribute('title') || '').indexOf('セキュリティ') >= 0;
                    if (
                      isVisible(el)
                      && !isInvisibleAnchor
                      && !isInvisibleWidget
                      && (isChallenge || isNormalCheckbox || isNormalWidget)
                    ) {
                      visibleMatches.push({
                        selector,
                        tag: el.tagName,
                        title: el.getAttribute('title') || '',
                        name: el.getAttribute('name') || '',
                        id: el.getAttribute('id') || '',
                        src,
                      });
                    }
                  }
                }
                return {
                  readyState: document.readyState || '',
                  title: document.title || '',
                  url: window.location.href || '',
                  visible_matches: visibleMatches,
                  requires_manual: visibleMatches.length > 0,
                  response_fields: responseFields,
                  response_length: responseFields.reduce((max, item) => Math.max(max, item.value_length || 0), 0),
                  response_present: responseFields.some((item) => item.value_present),
                };
                """
            )
        except Exception:
            return {
                "readyState": "",
                "title": getattr(driver, "title", ""),
                "url": getattr(driver, "current_url", ""),
                "visible_matches": [],
                "requires_manual": False,
                "response_fields": [],
                "response_length": 0,
                "response_present": False,
            }

    def get_recaptcha_response_info(self, driver):
        state = self.inspect_recaptcha_state(driver)
        return {
            "response_length": state.get("response_length", 0),
            "response_present": state.get("response_present", False),
            "response_fields": state.get("response_fields", []),
            "title": state.get("title", ""),
            "url": state.get("url", ""),
        }

    def wait_for_manual_captcha(self, driver, timeout_minutes=5, return_state=False):
        """Prompt for manual captcha solving and return immediately after acknowledgement."""
        state = self.inspect_recaptcha_state(driver)
        try:
            self.show_info(
                "reCAPTCHA 検出",
                "ページ上にreCAPTCHAが検出されました。ブラウザで手動で認証してください。完了したら「OK」を押してください。",
            )
        except Exception:
            print(
                "reCAPTCHA detected: please solve it manually in the browser and then continue."
            )
        response_info = self.get_recaptcha_response_info(driver)
        result = {
            "prompt_shown": True,
            "confirmed": True,
            "response_length": response_info.get("response_length", 0),
            "response_present": response_info.get("response_present", False),
            "response_fields": response_info.get("response_fields", []),
            "title": response_info.get("title", state.get("title", "")),
            "url": response_info.get("url", state.get("url", "")),
            "requires_manual": state.get("requires_manual", False),
            "visible_matches": state.get("visible_matches", []),
        }
        self.logger.info(
            "recaptcha manual prompt shown: %s",
            result,
        )
        if return_state:
            return result
        return True
