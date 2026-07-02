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
        return self.do_action(driver, "gLotWOpeLotSearchAction")

    def select_lottery_tennis_park(
        self,
        driver,
        park_value="1301270",
        court_value="12700020",
    ):
        self.execute_script(driver, "javascript:doLotEntry('130');")
        wait = self.wait_factory(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "bname")))
        wait.until(EC.presence_of_element_located((By.ID, "iname")))
        self._log_select_diagnostics(driver, "before_lottery_park_selection")
        Select(driver.find_element(By.ID, "bname")).select_by_value(park_value)
        self.execute_script(driver, "changeBname(document.form1);")
        try:
            wait.until(
                lambda d: self._has_option_value(
                    d,
                    select_id="iname",
                    option_value=court_value,
                )
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
            message = (
                "Lottery tennis park selection timed out while waiting for "
                f"iname option '{court_value}'. Context: "
                f"{json.dumps(debug_context, ensure_ascii=False)}"
            )
            self.logger.error(message)
            raise TimeoutException(message) from exc
        self._log_select_diagnostics(driver, "after_lottery_park_selection")
        iname_element = driver.find_element(By.ID, "iname")
        Select(iname_element).select_by_value(court_value)
        selected_state = self._describe_select(driver, "iname")
        selected_option = self._get_selected_option_state(driver, "iname")
        onchange_value = iname_element.get_attribute("onchange")
        js_info = self._inspect_lottery_iname_javascript(driver)
        self.logger.info(
            "after_iname_select: value=%s text=%s url=%s title=%s onchange=%s js=%s",
            selected_option.get("value"),
            selected_option.get("text"),
            driver.current_url,
            driver.title,
            onchange_value,
            json.dumps(js_info, ensure_ascii=False),
        )
        self.logger.info(
            "after_iname_select options=%s",
            json.dumps(selected_state.get("options", []), ensure_ascii=False),
        )
        self._trigger_iname_change(driver, iname_element, onchange_value)
        self._wait_for_lottery_entry_calendar(driver)
        self._save_named_debug_html(
            driver,
            "lottery_entry_after_iname_selection.html",
        )
        self._save_dom_summary(
            driver,
            "lottery_entry_after_iname_selection_dom_summary.json",
        )
        self._log_select_diagnostics(driver, "after_lottery_iname_change")

    def go_to_lottery_cancel_list(self, driver):
        return self.do_action(driver, "gLotWTransLotCancelListAction")

    def go_to_lottery_result_list(self, driver):
        return self.do_action(driver, "gLotWTransLotElectListAction")

    def go_to_reservation_list(self, driver):
        return self.do_action(driver, "gRsvWGetCancelRsvDataAction")

    def go_to_temp_apply(self, driver):
        return self.execute_script(
            driver,
            "javascript:doApplay(document.form1, gLotWInstTempLotApplyAction);",
        )

    def go_to_lottery_next_week(self, driver):
        current_headers = self._get_lottery_header_dates(driver)
        try:
            driver.find_element(By.ID, "next-week").click()
        except Exception:
            self.execute_script(
                driver,
                "doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
            )
        self._wait_for_lottery_entry_calendar(driver)
        updated_headers = self._get_lottery_header_dates(driver)
        return {
            "before_dates": current_headers,
            "after_dates": updated_headers,
            "changed": updated_headers != current_headers,
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

    def _describe_select(self, driver, select_id):
        try:
            element = driver.find_element(By.ID, select_id)
        except NoSuchElementException:
            return {
                "id": select_id,
                "exists": False,
                "value": None,
                "options": [],
            }

        select = Select(element)
        options = []
        for option in select.options:
            options.append(
                {
                    "value": option.get_attribute("value"),
                    "text": option.text.strip(),
                }
            )
        return {
            "id": select_id,
            "exists": True,
            "value": element.get_attribute("value"),
            "options": options,
        }

    def _build_select_debug_context(self, driver):
        context = {
            "current_url": driver.current_url,
            "title": driver.title,
            "bname": self._describe_select(driver, "bname"),
            "iname": self._describe_select(driver, "iname"),
            "iname_selected": self._get_selected_option_state(driver, "iname"),
        }
        return context

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
