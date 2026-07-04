# -*- coding: utf-8 -*-
"""Navigation helpers for legacy Selenium flows."""

from __future__ import annotations

import json
import logging
import time
from datetime import datetime
from pathlib import Path

from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select


class NavigationService:
    """Encapsulate repeated JavaScript execution and page transitions."""

    def __init__(
        self,
        wait_factory,
        sleep_func=time.sleep,
        logger=None,
        get_debug_output_dir=None,
    ):
        self.wait_factory = wait_factory
        self.sleep_func = sleep_func
        self.logger = logger or logging.getLogger(__name__)
        self.get_debug_output_dir = get_debug_output_dir

    def execute_script(self, driver, script, *args):
        return driver.execute_script(script, *args)

    def do_action(self, driver, action_name):
        return self.execute_script(
            driver,
            f"javascript:doAction(document.form1, {action_name});",
        )

    def logout(self, driver):
        return self.do_action(driver, "gRsvWTransUserAttestationEndAction")

    def go_to_lottery_entry(self, driver):
        if not self.wait_until_navigation_ready(driver):
            self.logger.warning(
                "Lottery navigation not ready: document/form1/doAction/action are unavailable."
            )
            return False
        try:
            self.do_action(driver, "gLotWOpeLotSearchAction")
        except Exception:
            return False
        return self.wait_until_lottery_search_ready(driver)

    def select_lottery_tennis_park(
        self,
        driver,
        park_value="1301270",
        court_value="12700020",
        court_text="テニス（人工芝）",
    ):
        if not self._wait_until_lottery_entry_selector_ready(driver):
            debug_html_path = self._save_debug_html(
                driver,
                prefix="lottery_entry_selector_not_ready",
            )
            state = self._inspect_lottery_entry_selector_state(driver)
            state["debug_html_path"] = str(debug_html_path) if debug_html_path else None
            message = (
                "Lottery entry selector is not ready before doLotEntry. Context: "
                f"{json.dumps(state, ensure_ascii=False)}"
            )
            self.logger.error(message)
            raise TimeoutException(message)
        self.execute_script(driver, "javascript:doLotEntry('130');")
        wait = self.wait_factory(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "bname")))
        wait.until(EC.presence_of_element_located((By.ID, "iname")))
        self.logger.info(
            "initial iname options=%s",
            json.dumps(self._describe_select(driver, "iname").get("options", []), ensure_ascii=False),
        )
        initial_settle = self._wait_for_lottery_entry_initial_settle(driver)
        self.logger.info(
            "after initial settle hidden values=%s",
            json.dumps(initial_settle, ensure_ascii=False),
        )
        self._log_select_diagnostics(driver, "before_lottery_park_selection")
        Select(driver.find_element(By.ID, "bname")).select_by_value(park_value)
        after_bname_hidden = self._inspect_lottery_hidden_selection_state(driver)
        self.logger.info(
            "after bname select hidden values=%s",
            json.dumps(after_bname_hidden, ensure_ascii=False),
        )
        self.execute_script(driver, "changeBname(document.form1);")
        retry_count = 0
        resolved_court_value = court_value
        selected_by = ""
        selected_state = {}
        selected_option = {}
        onchange_value = ""
        js_info = {}
        hidden_state = {}
        final_iname_options = []
        while retry_count <= 3:
            try:
                resolved_court_value = self._wait_for_lottery_iname_options(
                    driver,
                    option_value=court_value,
                    option_text=court_text,
                )
            except TimeoutException as exc:
                debug_context = self._build_select_debug_context(driver)
                debug_html_path = self._save_debug_html(
                    driver,
                    prefix="lottery_tennis_park_selection_failure",
                )
                debug_context["debug_html_path"] = (
                    str(debug_html_path) if debug_html_path else None
                )
                debug_context["retry_count"] = retry_count
                message = (
                    "Lottery tennis park selection timed out while waiting for "
                    f"iname option value '{court_value}'. Context: "
                    f"{json.dumps(debug_context, ensure_ascii=False)}"
                )
                self.logger.error(message)
                raise TimeoutException(message) from exc
            final_iname_options = self._describe_select(driver, "iname").get("options", [])
            self.logger.info(
                "after changeBname iname options=%s retry_count=%s",
                json.dumps(final_iname_options, ensure_ascii=False),
                retry_count,
            )
            iname_element = driver.find_element(By.ID, "iname")
            selected_by = self._select_lottery_court_option(
                iname_element,
                preferred_value=resolved_court_value,
                preferred_text=court_text,
            )
            selected_state = self._describe_select(driver, "iname")
            selected_option = self._get_selected_option_state(driver, "iname")
            onchange_value = iname_element.get_attribute("onchange")
            js_info = self._inspect_lottery_iname_javascript(driver)
            hidden_state = self._inspect_lottery_hidden_selection_state(driver)
            self.logger.info(
                "after iname select hidden values=%s retry_count=%s",
                json.dumps(hidden_state, ensure_ascii=False),
                retry_count,
            )
            if self._lottery_hidden_selection_matches(
                hidden_state,
                park_value=park_value,
                court_value=resolved_court_value,
            ):
                break
            self.sleep_func(0.5)
            self.execute_script(driver, "changeBname(document.form1);")
            retry_count += 1
        if not self._lottery_hidden_selection_matches(
            hidden_state,
            park_value=park_value,
            court_value=resolved_court_value,
        ):
            self._sync_lottery_hidden_selection(
                driver,
                park_value=park_value,
                court_value=resolved_court_value,
            )
            hidden_state = self._inspect_lottery_hidden_selection_state(driver)
        if not self._lottery_hidden_selection_matches(
            hidden_state,
            park_value=park_value,
            court_value=resolved_court_value,
        ):
            message = (
                "Lottery hidden selection mismatch after final sync. Context: "
                f"{json.dumps(hidden_state, ensure_ascii=False)}"
            )
            self.logger.error(message)
            raise TimeoutException(message)
        self.logger.info(
            "after_iname_select: selected_by=%s resolved_value=%s value=%s text=%s bname_value=%s bname_text=%s iname_value=%s iname_text=%s selectBldGrpCd=%s selectInstGrpCd=%s url=%s title=%s onchange=%s js=%s retry_count=%s",
            selected_by,
            resolved_court_value,
            selected_option.get("value"),
            selected_option.get("text"),
            selected_state.get("value"),
            selected_state.get("text"),
            hidden_state.get("bname_value"),
            hidden_state.get("bname_text"),
            hidden_state.get("iname_value"),
            hidden_state.get("iname_text"),
            hidden_state.get("selectBldGrpCd"),
            hidden_state.get("selectInstGrpCd"),
            driver.current_url,
            driver.title,
            onchange_value,
            json.dumps(js_info, ensure_ascii=False),
            retry_count,
        )
        self.logger.info(
            "after_iname_select options=%s final_bname_value=%s final_bname_text=%s final_iname_value=%s final_iname_text=%s final_selectBldGrpCd=%s final_selectInstGrpCd=%s",
            json.dumps(selected_state.get("options", []), ensure_ascii=False),
            hidden_state.get("bname_value"),
            hidden_state.get("bname_text"),
            hidden_state.get("iname_value"),
            hidden_state.get("iname_text"),
            hidden_state.get("selectBldGrpCd"),
            hidden_state.get("selectInstGrpCd"),
        )
        self._trigger_iname_change(driver, iname_element, onchange_value)
        self._wait_for_lottery_entry_calendar_with_retry(
            driver,
            preferred_value=resolved_court_value,
            preferred_text=court_text,
            onchange_value=onchange_value,
        )
        self._save_named_debug_html(
            driver,
            "lottery_entry_after_iname_selection.html",
        )
        self._save_dom_summary(
            driver,
            "lottery_entry_after_iname_selection_dom_summary.json",
        )
        self._log_select_diagnostics(driver, "after_lottery_iname_change")

    def _wait_for_lottery_entry_initial_settle(self, driver, timeout=10):
        deadline = time.time() + max(timeout, 1)
        last_snapshot = None
        last_hidden = None
        stable_count = 0
        while time.time() < deadline:
            try:
                loading_visible = self.execute_script(
                    driver,
                    """
                    const loading = document.getElementById("usedate-loading");
                    return !!loading &&
                      loading.style.display !== "none" &&
                      getComputedStyle(loading).display !== "none";
                    """,
                )
            except Exception:
                loading_visible = False
            iname_state = self._describe_select(driver, "iname")
            snapshot = tuple((opt.get("value"), opt.get("text")) for opt in iname_state.get("options", []))
            hidden_state = self._inspect_lottery_hidden_selection_state(driver)
            self.logger.info(
                "initial settle snapshot=%s hidden=%s loading_visible=%s",
                json.dumps(iname_state.get("options", []), ensure_ascii=False),
                json.dumps(hidden_state, ensure_ascii=False),
                loading_visible,
            )
            if not loading_visible and snapshot == last_snapshot and hidden_state == last_hidden:
                stable_count += 1
                if stable_count >= 2:
                    return hidden_state
            else:
                stable_count = 0
            last_snapshot = snapshot
            last_hidden = hidden_state
            self.sleep_func(0.3)
        return last_hidden or {}

    def _wait_until_lottery_entry_selector_ready(self, driver, timeout=10):
        deadline = time.time() + max(timeout, 1)
        last_state = {}
        while time.time() < deadline:
            last_state = self._inspect_lottery_entry_selector_state(driver)
            if (
                last_state.get("ready_state") == "complete"
                and last_state.get("has_form1")
                and last_state.get("has_doLotEntry")
            ):
                self.logger.info(
                    "lottery entry selector ready: %s",
                    json.dumps(last_state, ensure_ascii=False),
                )
                return True
            self.sleep_func(0.5)
        self.logger.warning(
            "lottery entry selector not ready: %s",
            json.dumps(last_state, ensure_ascii=False),
        )
        return False

    def _inspect_lottery_entry_selector_state(self, driver):
        try:
            return self.execute_script(
                driver,
                """
                return {
                  ready_state: document.readyState || "",
                  title: document.title || "",
                  current_url: window.location.href,
                  display_no: (document.querySelector('input[name="displayNo"]') || {}).value || "",
                  has_form1: !!document.form1,
                  has_doAction: typeof doAction === "function",
                  has_doLotEntry: typeof doLotEntry === "function",
                  has_lottery_action: typeof gLotWOpeLotSearchAction !== "undefined",
                  has_bname: !!document.getElementById("bname"),
                  has_iname: !!document.getElementById("iname"),
                };
                """,
            )
        except Exception as exc:
            return {"error": str(exc)}

    def _inspect_lottery_hidden_selection_state(self, driver):
        try:
            return self.execute_script(
                driver,
                """
                const getValue = (name) => {
                  const el =
                    document.querySelector(`input[name="${name}"]`) ||
                    document.getElementById(name);
                  return el ? (el.value || "") : "";
                };
                const bname = document.getElementById("bname");
                const iname = document.getElementById("iname");
                return {
                  bname_value: bname ? (bname.value || "") : "",
                  bname_text:
                    bname && bname.selectedIndex >= 0 && bname.options[bname.selectedIndex]
                      ? (bname.options[bname.selectedIndex].textContent || "").trim()
                      : "",
                  iname_value: iname ? (iname.value || "") : "",
                  iname_text:
                    iname && iname.selectedIndex >= 0 && iname.options[iname.selectedIndex]
                      ? (iname.options[iname.selectedIndex].textContent || "").trim()
                      : "",
                  selectBldGrpCd: getValue("selectBldGrpCd"),
                  selectInstGrpCd: getValue("selectInstGrpCd"),
                };
                """,
            )
        except Exception as exc:
            return {"error": str(exc)}

    def _lottery_hidden_selection_matches(self, hidden_state, park_value, court_value):
        hidden_state = hidden_state if isinstance(hidden_state, dict) else {}
        bname_value = str(hidden_state.get("bname_value", "")).strip()
        iname_value = str(hidden_state.get("iname_value", "")).strip()
        select_bld = str(hidden_state.get("selectBldGrpCd", "")).strip()
        select_inst = str(hidden_state.get("selectInstGrpCd", "")).strip()
        expected_park_value = str(park_value).strip()
        expected_court_value = str(court_value).strip()
        return (
            bname_value == expected_park_value
            and iname_value == expected_court_value
            and select_bld == expected_park_value
            and select_inst == expected_court_value
        )

    def _sync_lottery_hidden_selection(self, driver, park_value, court_value):
        try:
            self.execute_script(
                driver,
                """
                const setValue = (name, value) => {
                  const el =
                    document.querySelector(`input[name="${name}"]`) ||
                    document.getElementById(name);
                  if (el) {
                    el.value = value;
                    el.setAttribute("value", value);
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                  }
                };
                setValue("selectBldGrpCd", arguments[0]);
                setValue("selectInstGrpCd", arguments[1]);
                const bname = document.getElementById("bname");
                const iname = document.getElementById("iname");
                if (bname) {
                  bname.value = arguments[0];
                  bname.dispatchEvent(new Event("change", { bubbles: true }));
                }
                if (iname) {
                  iname.value = arguments[1];
                  iname.dispatchEvent(new Event("change", { bubbles: true }));
                }
                return true;
                """,
                park_value,
                court_value,
            )
            return True
        except Exception:
            return False

    def go_to_lottery_cancel_list(self, driver):
        return self.do_action(driver, "gLotWTransLotCancelListAction")

    def go_to_lottery_result_list(self, driver):
        return self.do_action(driver, "gLotWTransLotElectListAction")

    def go_to_reservation_list(self, driver):
        return self.do_action(driver, "gRsvWGetCancelRsvDataAction")

    def go_to_temp_apply(self, driver):
        result = {
            "success": False,
            "method": "",
            "pre_click": {},
            "post_click": {},
        }
        try:
            result["pre_click"] = self.execute_script(
                driver,
                """
                const button = document.getElementById("btn-go");
                const displayNo = document.querySelector('input[name="displayNo"]');
                return {
                  selectFieldCnt: document.getElementById("selectFieldCnt")
                    ? document.getElementById("selectFieldCnt").value || ""
                    : "",
                  display_no: displayNo ? (displayNo.value || "") : "",
                  title: document.title || "",
                  current_url: window.location.href,
                  btn_go: button
                    ? {
                        displayed: !!(button.offsetWidth || button.offsetHeight || button.getClientRects().length),
                        enabled: !button.disabled,
                        onclick: button.getAttribute("onclick") || "",
                      }
                    : null,
                };
                """,
            )
        except Exception:
            result["pre_click"] = {}

        def _capture_post_click(alert_text=""):
            try:
                state = self.inspect_page_state(driver)
            except Exception:
                state = {}
            state["has_apply"] = False
            state["alert_text"] = alert_text
            try:
                state["has_apply"] = bool(driver.find_elements(By.ID, "apply"))
            except Exception:
                pass
            return state

        def _consume_alert_text():
            try:
                alert = driver.switch_to.alert
                text = alert.text
                alert.accept()
                return text
            except Exception:
                return ""

        def _wait_for_confirmation():
            try:
                self.wait_factory(driver, 5).until(
                    lambda d: (
                        self.inspect_page_state(d).get("display_no") == "plwca1000"
                        or "申込内容確認画面" in getattr(d, "title", "")
                        or bool(d.find_elements(By.ID, "apply"))
                    )
                )
                return True
            except Exception:
                return False

        try:
            button = driver.find_element(By.ID, "btn-go")
            self.execute_script(
                driver,
                "arguments[0].scrollIntoView({block: 'center'});",
                button,
            )
            button.click()
            result["method"] = "native_click"
            alert_text = _consume_alert_text()
            if _wait_for_confirmation():
                result["success"] = True
                result["post_click"] = _capture_post_click(alert_text=alert_text)
                return result
        except Exception:
            pass

        try:
            button = driver.find_element(By.ID, "btn-go")
            self.execute_script(driver, "arguments[0].click();", button)
            result["method"] = "javascript_click"
            alert_text = _consume_alert_text()
            if _wait_for_confirmation():
                result["success"] = True
                result["post_click"] = _capture_post_click(alert_text=alert_text)
                return result
        except Exception:
            pass

        try:
            self.execute_script(
                driver,
                "javascript:doApplay(document.form1, gLotWInstTempLotApplyAction);",
            )
            result["method"] = "doApplay"
            alert_text = _consume_alert_text()
            result["success"] = _wait_for_confirmation()
            result["post_click"] = _capture_post_click(alert_text=alert_text)
            return result
        except Exception:
            result["post_click"] = _capture_post_click(alert_text="")
            return result

    def clear_lottery_selection(self, driver):
        try:
            button = driver.find_element(
                By.XPATH,
                "//button[contains(@onclick, 'clearUsedRowColumeAll')]",
            )
            self.execute_script(
                driver,
                "arguments[0].scrollIntoView({block: 'center'});",
                button,
            )
            button.click()
            return True
        except Exception:
            pass
        try:
            self.execute_script(driver, "clearUsedRowColumeAll(document.form1);")
            return True
        except Exception:
            return False

    def continue_lottery_entry_from_complete(self, driver):
        result = {
            "success": False,
            "fallback_used": False,
            "display_no": "",
            "title": "",
            "has_usedate_table": False,
        }
        try:
            button = driver.find_element(By.ID, "btn-light")
            self.execute_script(
                driver,
                "arguments[0].scrollIntoView({block: 'center'});",
                button,
            )
            button.click()
        except Exception:
            self.execute_script(
                driver,
                "javascript:doAction(document.form1, gWOpeTransLotInstSrchVacantAction);",
            )
        try:
            self._wait_for_lottery_entry_calendar(driver)
        except Exception:
            pass
        page_state = self.inspect_page_state(driver)
        result.update(page_state)
        result["success"] = bool(
            page_state.get("has_usedate_table")
            and page_state.get("display_no") == "plwba4000"
        )
        return result

    def inspect_page_state(self, driver):
        return self.execute_script(
            driver,
            """
            const displayNo = document.querySelector('input[name="displayNo"]');
            return {
              title: document.title || "",
              current_url: window.location.href,
              display_no: displayNo ? (displayNo.value || "") : "",
              has_usedate_table: !!document.getElementById("usedate-table"),
              has_btn_go: !!document.getElementById("btn-go"),
            };
            """,
        )

    def wait_until_navigation_ready(self, driver, timeout=10):
        try:
            self.wait_factory(driver, timeout).until(
                lambda d: self.execute_script(
                    d,
                    """
                    return (
                      document.readyState === "complete" &&
                      typeof doAction === "function" &&
                      !!document.form1 &&
                      typeof gLotWOpeLotSearchAction !== "undefined"
                    );
                    """,
                )
            )
            return True
        except Exception:
            return False

    def wait_until_lottery_search_ready(self, driver, timeout=10):
        try:
            self.wait_factory(driver, timeout).until(
                lambda d: self.execute_script(
                    d,
                    """
                    return (
                      document.readyState === "complete" &&
                      !!document.form1 &&
                      (
                        (document.title || "").indexOf("抽選分類一覧画面") >= 0 ||
                        window.location.href.indexOf("lotWOpeLotSearchAction") >= 0 ||
                        typeof doLotEntry === "function"
                      )
                    );
                    """,
                )
            )
            return True
        except Exception:
            return False

    def go_to_lottery_next_week(self, driver):
        current_headers = self._get_lottery_header_dates(driver)
        updated_headers = current_headers
        methods = ("native_click", "javascript_click", "doNextWeek")
        used_method = ""
        for method in methods:
            try:
                if method == "native_click":
                    button = driver.find_element(By.ID, "next-week")
                    self.execute_script(
                        driver,
                        "arguments[0].scrollIntoView({block: 'center'});",
                        button,
                    )
                    button.click()
                elif method == "javascript_click":
                    button = driver.find_element(By.ID, "next-week")
                    self.execute_script(driver, "arguments[0].click();", button)
                else:
                    self.execute_script(
                        driver,
                        "doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
                    )
                used_method = method
                self._wait_for_lottery_entry_calendar(driver)
                updated_headers = self._wait_for_header_change(
                    driver,
                    previous_headers=current_headers,
                    fallback_headers=updated_headers,
                )
                if updated_headers != current_headers:
                    break
            except Exception:
                try:
                    updated_headers = self._wait_for_header_change(
                        driver,
                        previous_headers=current_headers,
                        fallback_headers=updated_headers,
                        timeout=2,
                    )
                    if updated_headers != current_headers:
                        used_method = method
                        break
                except Exception:
                    pass
                continue
        self.logger.info(
            "go_to_lottery_next_week before=%s after=%s changed=%s method=%s",
            current_headers,
            updated_headers,
            updated_headers != current_headers,
            used_method,
        )
        return {
            "before_dates": current_headers,
            "after_dates": updated_headers,
            "changed": updated_headers != current_headers,
            "method": used_method,
        }

    def go_to_lottery_previous_week(self, driver):
        current_headers = self._get_lottery_header_dates(driver)
        updated_headers = current_headers
        methods = ("native_click", "javascript_click", "doPrevWeek")
        used_method = ""
        for method in methods:
            try:
                if method == "native_click":
                    button = driver.find_element(By.ID, "last-week")
                    self.execute_script(
                        driver,
                        "arguments[0].scrollIntoView({block: 'center'});",
                        button,
                    )
                    button.click()
                elif method == "javascript_click":
                    button = driver.find_element(By.ID, "last-week")
                    self.execute_script(driver, "arguments[0].click();", button)
                else:
                    self.execute_script(
                        driver,
                        "doPrevWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
                    )
                used_method = method
                self._wait_for_lottery_entry_calendar(driver)
                updated_headers = self._wait_for_header_change(
                    driver,
                    previous_headers=current_headers,
                    fallback_headers=updated_headers,
                )
                if updated_headers != current_headers:
                    break
            except Exception:
                try:
                    updated_headers = self._wait_for_header_change(
                        driver,
                        previous_headers=current_headers,
                        fallback_headers=updated_headers,
                        timeout=2,
                    )
                    if updated_headers != current_headers:
                        used_method = method
                        break
                except Exception:
                    pass
                continue
        self.logger.info(
            "go_to_lottery_previous_week before=%s after=%s changed=%s method=%s",
            current_headers,
            updated_headers,
            updated_headers != current_headers,
            used_method,
        )
        return {
            "before_dates": current_headers,
            "after_dates": updated_headers,
            "changed": updated_headers != current_headers,
            "method": used_method,
        }

    def go_to_vacant_search(self, driver):
        self.execute_script(
            driver,
            "javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formWTransInstSrchVacantAction : document.formWTransInstSrchVacantAction ), gRsvWTransInstSrchVacantAction);",
        )
        self.execute_script(
            driver,
            "javascript:doComplexSearchAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction);",
        )

    def select_weekly_vacant_conditions(self, driver):
        self.execute_script(
            driver,
            "javaScript: sendSelectWeekNum2((_dom == 3) ? document.layers['disp'].document.form1: document.form1, gRsvWTransInstSrchPpsAction);",
        )
        self.execute_script(
            driver,
            "javascript:doTransInstSrchMultipleAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction, '1000', '1030');",
        )
        self.execute_script(
            driver,
            "javascript:sendSelectWeekNum((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWGetInstSrchInfAction);",
        )

    def _has_option_value(self, driver, select_id, option_value):
        select_state = self._describe_select(driver, select_id)
        options = select_state.get("options", [])
        return any(option.get("value") == option_value for option in options)

    def _has_option_value_or_text(self, driver, select_id, option_value, option_text):
        select_state = self._describe_select(driver, select_id)
        options = select_state.get("options", [])
        return any(
            option.get("value") == option_value
            for option in options
        )

    def _wait_for_lottery_iname_options(self, driver, option_value, option_text):
        wait = self.wait_factory(driver, 10)
        wait.until(
            lambda d: self._has_option_value_or_text(
                d,
                select_id="iname",
                option_value=option_value,
                option_text=option_text,
            )
        )
        deadline = time.time() + 5
        last_snapshot = None
        stable_count = 0
        resolved_value = option_value
        while time.time() < deadline:
            select_state = self._describe_select(driver, "iname")
            snapshot = tuple(
                (opt.get("value"), opt.get("text"))
                for opt in select_state.get("options", [])
            )
            matched_value = None
            for opt_value, opt_text in snapshot:
                if opt_value == option_value:
                    matched_value = opt_value
                    break
            if matched_value:
                resolved_value = matched_value
                if snapshot == last_snapshot:
                    stable_count += 1
                else:
                    stable_count = 1
                if stable_count >= 3:
                    self.logger.info(
                        "lottery iname options stabilized: resolved_value=%s options=%s",
                        resolved_value,
                        json.dumps(select_state.get("options", []), ensure_ascii=False),
                    )
                    return resolved_value
            last_snapshot = snapshot
            self.sleep_func(0.3)
        self.logger.info(
            "lottery iname options fallback without full stabilization: resolved_value=%s",
            resolved_value,
        )
        if resolved_value != option_value:
            raise TimeoutException(
                f"Lottery iname option value not found: expected={option_value}"
            )
        return resolved_value

    def _select_lottery_court_option(self, select_element, preferred_value, preferred_text):
        select = Select(select_element)
        for option in select.options:
            if option.get_attribute("value") == preferred_value:
                select.select_by_value(preferred_value)
                return f"value:{preferred_value}"
        raise NoSuchElementException(
            f"Lottery court option not found: value={preferred_value} text={preferred_text}"
        )

    def _describe_select(self, driver, select_id):
        try:
            state = self.execute_script(
                driver,
                """
                const select = document.getElementById(arguments[0]);
                if (!select) {
                  return {
                    id: arguments[0],
                    exists: false,
                    value: null,
                    options: [],
                  };
                }
                return {
                  id: arguments[0],
                  exists: true,
                  value: select.value || "",
                  options: Array.from(select.options || []).map((option) => ({
                    value: option.value || "",
                    text: (option.textContent || "").trim(),
                  })),
                };
                """,
                select_id,
            )
            if isinstance(state, dict):
                return state
        except Exception:
            pass
        return {
            "id": select_id,
            "exists": False,
            "value": None,
            "options": [],
        }

    def _build_select_debug_context(self, driver):
        context = {
            "current_url": driver.current_url,
            "title": driver.title,
            "bname": self._describe_select(driver, "bname"),
            "bname_selected": self._get_selected_option_state(driver, "bname"),
            "iname": self._describe_select(driver, "iname"),
            "iname_selected": self._get_selected_option_state(driver, "iname"),
        }
        return context

    def inspect_lottery_park_selection(self, driver):
        return self._build_select_debug_context(driver)

    def _log_select_diagnostics(self, driver, stage):
        context = self._build_select_debug_context(driver)
        self.logger.info(
            "%s: url=%s title=%s bname_exists=%s iname_exists=%s",
            stage,
            context["current_url"],
            context["title"],
            context["bname"]["exists"],
            context["iname"]["exists"],
        )
        self.logger.info(
            "%s bname options=%s",
            stage,
            json.dumps(context["bname"]["options"], ensure_ascii=False),
        )
        self.logger.info(
            "%s iname options=%s",
            stage,
            json.dumps(context["iname"]["options"], ensure_ascii=False),
        )

    def _save_debug_html(self, driver, prefix):
        if self.get_debug_output_dir is None:
            return None
        debug_dir = Path(self.get_debug_output_dir())
        debug_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        debug_path = debug_dir / f"{prefix}_{timestamp}.html"
        debug_path.write_text(driver.page_source, encoding="utf-8")
        return debug_path

    def _save_named_debug_html(self, driver, filename):
        if self.get_debug_output_dir is None:
            return None
        debug_dir = Path(self.get_debug_output_dir())
        debug_dir.mkdir(parents=True, exist_ok=True)
        debug_path = debug_dir / filename
        debug_path.write_text(driver.page_source, encoding="utf-8")
        return debug_path

    def _save_dom_summary(self, driver, filename):
        if self.get_debug_output_dir is None:
            return None
        debug_dir = Path(self.get_debug_output_dir())
        debug_dir.mkdir(parents=True, exist_ok=True)
        debug_path = debug_dir / filename
        summary = self.execute_script(
            driver,
            """
            const pick = (selector, mapper) =>
              Array.from(document.querySelectorAll(selector)).slice(0, 100).map(mapper);
            const text = (value) => (value || "").replace(/\\s+/g, " ").trim();
            return {
              current_url: window.location.href,
              title: document.title,
              forms: pick("form", (el) => ({
                name: el.getAttribute("name") || "",
                id: el.id || "",
                action: el.getAttribute("action") || "",
                method: el.getAttribute("method") || "",
              })),
              selects: pick("select", (el) => ({
                id: el.id || "",
                name: el.getAttribute("name") || "",
                value: el.value || "",
                onchange: el.getAttribute("onchange") || "",
                option_count: el.options ? el.options.length : 0,
                selected_text:
                  el.selectedIndex >= 0 && el.options[el.selectedIndex]
                    ? text(el.options[el.selectedIndex].textContent)
                    : "",
              })),
              buttons: pick("button", (el) => ({
                id: el.id || "",
                type: el.getAttribute("type") || "",
                onclick: el.getAttribute("onclick") || "",
                text: text(el.textContent),
                disabled: !!el.disabled,
              })),
              inputs: pick("input", (el) => ({
                id: el.id || "",
                name: el.getAttribute("name") || "",
                type: el.getAttribute("type") || "",
                value: el.value || "",
                checked: !!el.checked,
              })),
              links: pick("a", (el) => ({
                id: el.id || "",
                href: el.getAttribute("href") || "",
                onclick: el.getAttribute("onclick") || "",
                text: text(el.textContent),
              })),
              scripts: pick("script", (el) => ({
                src: el.getAttribute("src") || "",
                inline_preview: text(el.textContent).slice(0, 300),
              })),
            };
            """,
        )
        debug_path.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return debug_path

    def save_debug_html(self, driver, filename):
        return self._save_named_debug_html(driver, filename)

    def save_dom_summary(self, driver, filename):
        return self._save_dom_summary(driver, filename)

    def wait_for_lottery_entry_calendar(self, driver):
        return self._wait_for_lottery_entry_calendar(driver)

    def _get_selected_option_state(self, driver, select_id):
        try:
            element = driver.find_element(By.ID, select_id)
        except NoSuchElementException:
            return {"id": select_id, "exists": False, "value": None, "text": None}
        select = Select(element)
        if not select.options:
            return {"id": select_id, "exists": True, "value": "", "text": ""}
        selected_option = select.first_selected_option
        return {
            "id": select_id,
            "exists": True,
            "value": selected_option.get_attribute("value"),
            "text": selected_option.text.strip(),
        }

    def _inspect_lottery_iname_javascript(self, driver):
        return self.execute_script(
            driver,
            """
            const iname = document.getElementById("iname");
            return {
              onchange: iname ? (iname.getAttribute("onchange") || "") : "",
              has_changeIname: typeof window.changeIname === "function",
              has_changeBname: typeof window.changeBname === "function",
              has_sendInstGrCd: typeof window.sendInstGrCd === "function",
              ajax_action:
                typeof window.gLotWTransLotInstSrchVacantAjaxAction !== "undefined"
                  ? String(window.gLotWTransLotInstSrchVacantAjaxAction)
                  : "",
            };
            """,
        )

    def _trigger_iname_change(self, driver, iname_element, onchange_value):
        self.execute_script(
            driver,
            """
            arguments[0].dispatchEvent(new Event("change", { bubbles: true }));
            """,
            iname_element,
        )
        if onchange_value:
            script = onchange_value.strip()
            if script.lower().startswith("javascript:"):
                script = script[len("javascript:") :]
            self.execute_script(driver, script)

    def _wait_for_lottery_entry_calendar_with_retry(
        self,
        driver,
        preferred_value,
        preferred_text,
        onchange_value,
    ):
        try:
            self._wait_for_lottery_entry_calendar(driver)
            return
        except TimeoutException as exc:
            selected_option = self._get_selected_option_state(driver, "iname")
            select_state = self._describe_select(driver, "iname")
            has_target = any(
                option.get("value") == preferred_value
                or option.get("text") == preferred_text
                for option in select_state.get("options", [])
            )
            selected_matches = (
                selected_option.get("value") == preferred_value
                or selected_option.get("text") == preferred_text
            )
            if not has_target or selected_matches:
                raise
            self.logger.info(
                "lottery iname selection reset detected; retrying selection once. selected=%s options=%s",
                json.dumps(selected_option, ensure_ascii=False),
                json.dumps(select_state.get("options", []), ensure_ascii=False),
            )
            iname_element = driver.find_element(By.ID, "iname")
            selected_by = self._select_lottery_court_option(
                iname_element,
                preferred_value=preferred_value,
                preferred_text=preferred_text,
            )
            self.logger.info("lottery iname retry selected_by=%s", selected_by)
            self._trigger_iname_change(driver, iname_element, onchange_value)
            self._wait_for_lottery_entry_calendar(driver)

    def _wait_for_lottery_entry_calendar(self, driver):
        wait = self.wait_factory(driver, 15)
        try:
            wait.until(
                lambda d: self.execute_script(
                    d,
                    """
                    const loading = document.getElementById("usedate-loading");
                    const loadingVisible =
                      !!loading &&
                      loading.style.display !== "none" &&
                      getComputedStyle(loading).display !== "none";
                    const headerCount = document.querySelectorAll(
                      '#usedate-table thead input[name="selectUseYMD"]'
                    ).length;
                    const bodyRowCount = document.querySelectorAll(
                      "#usedate-table tbody tr"
                    ).length;
                    return !loadingVisible && (headerCount > 0 || bodyRowCount > 0);
                    """,
                )
            )
        except TimeoutException as exc:
            debug_context = self._build_select_debug_context(driver)
            debug_context["calendar_state"] = self.execute_script(
                driver,
                """
                const loading = document.getElementById("usedate-loading");
                return {
                  loading_display: loading ? loading.style.display || "" : "",
                  computed_loading_display: loading ? getComputedStyle(loading).display : "",
                  header_count: document.querySelectorAll(
                    '#usedate-table thead input[name="selectUseYMD"]'
                  ).length,
                  body_row_count: document.querySelectorAll("#usedate-table tbody tr").length,
                };
                """,
            )
            self._save_named_debug_html(
                driver,
                "lottery_entry_after_iname_selection.html",
            )
            self._save_dom_summary(
                driver,
                "lottery_entry_after_iname_selection_dom_summary.json",
            )
            message = (
                "Lottery entry calendar did not finish loading after iname selection. "
                f"Context: {json.dumps(debug_context, ensure_ascii=False)}"
            )
            self.logger.error(message)
            raise TimeoutException(message) from exc

    def _get_lottery_header_dates(self, driver):
        return self.execute_script(
            driver,
            """
            return Array.from(
              document.querySelectorAll('#usedate-table thead input[name="selectUseYMD"]')
            ).map((input) => input.value || "");
            """,
        )

    def _wait_for_header_change(self, driver, previous_headers, fallback_headers, timeout=5):
        deadline = time.time() + max(timeout, 1)
        latest_headers = fallback_headers
        while time.time() < deadline:
            try:
                latest_headers = self._get_lottery_header_dates(driver)
            except Exception:
                latest_headers = fallback_headers
            if latest_headers != previous_headers:
                return latest_headers
            self.sleep_func(0.2)
        return latest_headers
