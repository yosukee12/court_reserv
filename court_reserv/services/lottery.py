# -*- coding: utf-8 -*-
"""Lottery business logic extracted from the legacy UI class."""

import json
import re
import time
from datetime import datetime

from bs4 import BeautifulSoup as bs
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException,
    UnexpectedAlertPresentException,
)


class LotteryService:
    """Encapsulate lottery entry and lottery-check flows."""

    def __init__(
        self,
        config,
        browser_session,
        login_service,
        navigation_service,
        logger,
        show_info,
        output_id_dict,
        sleep_func=time.sleep,
    ):
        self.config = config
        self.browser_session = browser_session
        self.login_service = login_service
        self.navigation_service = navigation_service
        self.logger = logger
        self.show_info = show_info
        self.output_id_dict = output_id_dict
        self.sleep_func = sleep_func
        self._last_slot_select_info = {}

    def _parse_slot_components(self, slot_text):
        parts = str(slot_text or "").split()
        ymd = parts[0] if parts else ""
        stime = ""
        etime = ""
        field = ""
        for part in parts:
            if "-" in part and part.split("-")[0].isdigit():
                stime = part.split("-")[0]
                etime = part.split("-")[1]
            elif part.startswith("fields:"):
                field = part.split(":", 1)[1]
        return {
            "ymd": ymd,
            "stime": stime,
            "etime": etime,
            "field": field,
        }

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

    def _show_manual_submit_prompt(self):
        try:
            self.show_info(
                "手動操作が必要です",
                "申込み実行画面が表示されました。画面上で申込み実行を手動で実施してください。reCAPTCHAが表示された場合はブラウザ上で認証してください。",
            )
        except Exception:
            print(
                "Manual submit required: please click submit in the browser. If reCAPTCHA appears, solve it manually."
            )

    def semiauto_reserv(self, id_dict):
        """Run the legacy semi-auto lottery entry flow."""
        list_count = 1
        driver = self.browser_session.create_driver()
        try:
            for user_id, user_values in id_dict.items():
                reserv_count = 0
                print(
                    "申し込み "
                    + str(list_count)
                    + "人目/"
                    + str(len(id_dict))
                    + "人"
                    + user_values[0]
                )
                try:
                    self._legacy_login(driver, user_id, user_values[2])
                except UnexpectedAlertPresentException:
                    print("ID:" + user_id + " 期限切れ")
                    self.logger.warning("ID:" + user_id + " 期限切れ")
                    continue

                self.logger.info("ID:" + user_id + " ログイン")

                if "ホーム画面" in driver.title:
                    self.navigation_service.go_to_lottery_entry(driver)
                    self.navigation_service.select_lottery_tennis_park(driver)

                    while reserv_count < 2:
                        self.sleep_func(0.5)
                        try:
                            if "東京都スポーツ施設サービス" in driver.title:
                                self.logger.info("ID:" + user_id + " ログアウト")
                                break
                            if "申込内容確認画面" in driver.title:
                                reserv_count += 1
                                soup = bs(driver.page_source, "html.parser")
                                soup.find_all(
                                    "td", string=["年", "月", "日", "時", "分"]
                                )
                                if reserv_count == 1:
                                    self.sleep_func(0.3)
                                    Select(driver.find_element(By.ID, "apply")).select_by_value(
                                        "1-1"
                                    )
                                    self.sleep_func(0.2)
                                elif reserv_count == 2:
                                    self.sleep_func(0.3)
                                    Select(driver.find_element(By.ID, "apply")).select_by_value(
                                        "2-1"
                                    )
                                    self.sleep_func(0.2)

                                try:
                                    if self.login_service.detect_captcha(driver):
                                        self._show_manual_submit_prompt()
                                    else:
                                        print(
                                            "Manual submit required: please click submit in the browser. If reCAPTCHA appears, solve it manually."
                                        )
                                except Exception:
                                    print(
                                        "Manual submit required: please click submit in the browser. If reCAPTCHA appears, solve it manually."
                                    )

                                while "抽選メール送信完了画面" not in driver.title:
                                    try:
                                        try:
                                            if self.login_service.detect_captcha(driver):
                                                solved = self.login_service.wait_for_manual_captcha(
                                                    driver
                                                )
                                                if not solved:
                                                    self.logger.warning(
                                                        "User cancelled captcha solving; aborting this reservation"
                                                    )
                                                    break
                                        except Exception:
                                            pass

                                        self._get_wait(driver, 60).until(
                                            EC.alert_is_present(),
                                            "Timed out waiting for PA creation confirmation popup to appear.",
                                        )
                                        alert = driver.switch_to.alert
                                        alert.accept()
                                        self._get_wait(driver, 1).until(
                                            EC.alert_is_present(),
                                            "Timed out waiting for PA creation confirmation popup to appear.",
                                        )
                                        self.sleep_func(0.3)
                                    except (
                                        TimeoutException,
                                        UnexpectedAlertPresentException,
                                    ):
                                        continue
                                print(
                                    "reserved: ID = "
                                    + user_id
                                    + ", reserv_count = "
                                    + str(reserv_count)
                                )
                        except (
                            TimeoutException,
                            UnexpectedAlertPresentException,
                        ):
                            continue
                list_count += 1
                self.sleep_func(0.5)
                self.navigation_service.logout(driver)
                self.sleep_func(0.5)
        finally:
            self.browser_session.safe_close(driver)

    def full_auto_reserv(self, id_dict, max_attempts=2):
        """Run the legacy full-auto lottery entry flow."""
        list_count = 1
        driver = self.browser_session.create_driver()
        try:
            for user_id, user_values in id_dict.items():
                reserv_count = 0
                print(
                    "自動申し込み "
                    + str(list_count)
                    + "人目/"
                    + str(len(id_dict))
                    + "人"
                    + user_values[0]
                )
                if not self.login_service.login(driver, user_id, user_values[2]):
                    continue
                self.logger.info("ID:%s ログイン", user_id)

                if "ホーム画面" in driver.title:
                    try:
                        self.navigation_service.go_to_lottery_entry(driver)
                        self.navigation_service.select_lottery_tennis_park(driver)
                        while reserv_count < max_attempts:
                            self.sleep_func(0.5)
                            if "申込内容確認画面" in driver.title:
                                reserv_count += 1
                                sel_val = f"{reserv_count}-1"
                                Select(driver.find_element(By.ID, "apply")).select_by_value(
                                    sel_val
                                )
                                self.sleep_func(0.2)
                                try:
                                    try:
                                        if self.login_service.detect_captcha(driver):
                                            ok = self.login_service.wait_for_manual_captcha(
                                                driver
                                            )
                                            if not ok:
                                                self.logger.warning(
                                                    "ID:%s captcha not solved, aborting auto apply",
                                                    user_id,
                                                )
                                                break
                                    except Exception:
                                        pass
                                    self.navigation_service.execute_script(
                                        driver,
                                        "javascript:sendLotApply(document.form1, gLotWInstLotApplyAction, event);",
                                    )
                                except Exception:
                                    self.logger.warning(
                                        "ID:%s 自動申込みスクリプト実行に失敗", user_id
                                    )
                                try:
                                    self._get_wait(driver, 60).until(
                                        EC.alert_is_present()
                                    )
                                    alert = driver.switch_to.alert
                                    alert.accept()
                                except TimeoutException:
                                    self.logger.warning(
                                        "ID:%s 申込み確認ポップアップが表示されませんでした",
                                        user_id,
                                    )
                                try:
                                    self._get_wait(driver, 10).until(
                                        lambda d: "抽選メール送信完了画面" in d.title
                                    )
                                except Exception:
                                    self.logger.info(
                                        "ID:%s 抽選送信完了画面への遷移を確認できませんでした",
                                        user_id,
                                    )
                                print(
                                    "reserved: ID = "
                                    + user_id
                                    + ", reserv_count = "
                                    + str(reserv_count)
                                )
                            elif "東京都スポーツ施設サービス" in driver.title:
                                self.logger.info("ID:%s ログアウト検出", user_id)
                                break
                            else:
                                self.sleep_func(0.5)
                    except Exception:
                        self.logger.exception("ID:%s 自動申込み中に例外", user_id)

                list_count += 1
                self.sleep_func(0.5)
                try:
                    self.navigation_service.logout(driver)
                except Exception:
                    pass
                self.sleep_func(0.5)
        finally:
            self.browser_session.safe_close(driver)

    def auto_select_and_submit_slots(
        self,
        driver,
        selected_slots,
        submit=True,
        wait_alert_seconds=10,
    ):
        """Auto-select lottery slots on the current lottery page."""
        result = {}
        if not driver:
            raise RuntimeError("Driver not started")

        for slot in selected_slots:
            slot_parts = self._parse_slot_components(slot)
            ymd = slot_parts["ymd"]
            stime = slot_parts["stime"]
            etime = slot_parts["etime"]
            field = slot_parts["field"]
            if ymd:
                try:
                    attempts = 0
                    print(f"[debug] Ensure week for {ymd} is displayed")
                    while attempts < 10:
                        attempts += 1
                        headers = self.navigation_service.execute_script(
                            driver,
                            "return Array.from(document.querySelectorAll('#usedate-table thead input[name=\"selectUseYMD\"]')).map(h=>h.value);",
                        )
                        print(f"[debug] attempt {attempts}, headers={headers}")
                        if ymd in headers:
                            print(f"[debug] target {ymd} found in headers")
                            stable_headers = self._wait_for_target_headers_stable(
                                driver, ymd
                            )
                            print(
                                f"[debug] target {ymd} stabilized with headers={stable_headers}"
                            )
                            if ymd in stable_headers:
                                break
                            print(
                                f"[debug] target {ymd} disappeared during stabilization, continue searching"
                            )
                        if headers:
                            try:
                                min_h = min(headers)
                                max_h = max(headers)
                                print(f"[debug] min_h={min_h}, max_h={max_h}")
                                if ymd > max_h:
                                    print("[debug] clicking next-week")
                                    nav_result = self.navigation_service.go_to_lottery_next_week(
                                        driver
                                    )
                                    if not nav_result.get("changed"):
                                        print("[debug] next-week did not change headers")
                                elif ymd < min_h:
                                    print("[debug] clicking last-week")
                                    nav_result = self.navigation_service.go_to_lottery_previous_week(
                                        driver
                                    )
                                    if not nav_result.get("changed"):
                                        print("[debug] last-week did not change headers")
                                else:
                                    print("[debug] not in range, clicking next-week")
                                    nav_result = self.navigation_service.go_to_lottery_next_week(
                                        driver
                                    )
                                    if not nav_result.get("changed"):
                                        print("[debug] next-week did not change headers")
                            except Exception as exc:
                                print(f"[debug] header compare failed: {exc}")
                        self.sleep_func(0.6)
                    print(f"[debug] finished pagination attempts for {ymd}")
                except Exception:
                    pass
            js = """
            (function(ymd, stime, etime, field) {
                var info = {
                  found:false,
                  idx:-1,
                  clicked:false,
                  reason:'',
                  headers:[],
                  matched_header:'',
                  matched_time:'',
                  matched_end_time:'',
                  matched_field:'',
                  current_display_no:'',
                  selected_slots_count:0,
                  checked_after_mark:{}
                };
                const markChecked = (input) => {
                  if (!input) return false;
                  const name = input.getAttribute('name') || '';
                  const value = input.value || '';
                  try {
                    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'checked');
                    if (setter && setter.set) {
                      setter.set.call(input, true);
                    } else {
                      input.checked = true;
                    }
                  } catch (e) {}
                  try { input.defaultChecked = true; } catch (e) {}
                  try { input.setAttribute('checked', 'checked'); } catch (e) {}
                  try {
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                  } catch (e) {}
                  try {
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                  } catch (e) {}
                  if (!input.checked) {
                    try { input.click(); } catch (e) {}
                    try {
                      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'checked');
                      if (setter && setter.set) {
                        setter.set.call(input, true);
                      } else {
                        input.checked = true;
                      }
                    } catch (e) {}
                    try { input.defaultChecked = true; } catch (e) {}
                    try { input.setAttribute('checked', 'checked'); } catch (e) {}
                    try {
                      input.dispatchEvent(new Event('input', { bubbles: true }));
                    } catch (e) {}
                    try {
                      input.dispatchEvent(new Event('change', { bubbles: true }));
                    } catch (e) {}
                  }
                  try {
                    info.checked_after_mark[name + ':' + value] = !!input.checked;
                  } catch (e) {}
                  return !!input.checked;
                };
                try {
                    var displayNo = document.querySelector('input[name="displayNo"]');
                    info.current_display_no = displayNo ? (displayNo.value || '') : '';
                    var headerInputs = Array.from(document.querySelectorAll('#usedate-table thead input[name="selectUseYMD"]'));
                    var headers = headerInputs.map(function(h){ return String(h.value || '').trim(); });
                    info.headers = headers.slice();
                    ymd = String(ymd || '').trim();
                    stime = String(stime || '').trim();
                    etime = String(etime || '').trim();
                    field = String(field || '').trim();
                    for (var i=0;i<headers.length;i++){
                        if (headers[i] === ymd){ info.idx = i; break; }
                    }
                    if (info.idx === -1){ info.reason='ymd not in headers'; return info; }
                    info.matched_header = headers[info.idx];
                    var rows = document.querySelectorAll('#usedate-table tbody tr');
                    for (var r=0;r<rows.length;r++){
                        var tds = rows[r].querySelectorAll('td');
                        var td = tds[info.idx];
                        if (!td) continue;
                        var st = td.querySelector('input[name="selectStime"]');
                        var et = td.querySelector('input[name="selectEtime"]');
                        var fieldInput = td.querySelector('input[name="selectField"]');
                        var koma = td.querySelector('input[name="selectKomaNo"]');
                        if (st){
                            try{
                                if (
                                  String(parseInt(st.value,10)) === String(parseInt(stime,10)) &&
                                  (!etime || (et && String(parseInt(et.value,10)) === String(parseInt(etime,10)))) &&
                                  (!field || (fieldInput && String(fieldInput.value || '') === field))
                                ){
                                    info.found = true;
                                    info.matched_time = String(st.value || '');
                                    info.matched_end_time = et ? String(et.value || '') : '';
                                    info.matched_field = fieldInput ? String(fieldInput.value || '') : '';
                                    try{ td.click(); info.clicked = true; }catch(e){}
                                    try{
                                        var num = td.querySelector('span.font-weight-bold');
                                        if(num){ num.click(); info.clicked = true; }
                                    }catch(e){}
                                    try { markChecked(headerInputs[info.idx]); } catch (e) {}
                                    try { markChecked(koma); } catch (e) {}
                                    try { markChecked(st); } catch (e) {}
                                    try { markChecked(et); } catch (e) {}
                                    try { markChecked(fieldInput); } catch (e) {}
                                    try {
                                      if (td.classList) {
                                        td.classList.add('selected');
                                        td.classList.add('active');
                                      }
                                    } catch (e) {}
                                    try {
                                      const form = document.forms['form1'];
                                      if (form && form.selectFieldCnt) {
                                        form.selectFieldCnt.value = '1';
                                      }
                                    } catch (e) {}
                                    info.clicked_element_tag = st ? (st.tagName || '') : (td.tagName || '');
                                    info.clicked_element_text = [ymd, st ? (st.value || '') : '', et ? (et.value || '') : '', fieldInput ? (fieldInput.value || '') : '']
                                      .filter(Boolean)
                                      .join(' ');
                                    info.clicked_element_outer_html = st ? (st.outerHTML || '') : (td.outerHTML || '');
                                    try{
                                      const checkedInputs = Array.from(document.querySelectorAll('input[type="checkbox"]')).filter(function(el){ return el.checked; });
                                      info.selected_slots_count = Array.from(document.querySelectorAll('input[name="selectStime"]')).filter(function(el){ return el.checked; }).length;
                                      info.checked_input_count = checkedInputs.length;
                                      info.checked_inputs = checkedInputs.map(function(el) {
                                        return {
                                          name: el.getAttribute('name') || '',
                                          value: el.value || ''
                                        };
                                      });
                                      info.checked_selectUseYMD = checkedInputs.filter(function(el){ return el.name === 'selectUseYMD'; }).map(function(el){ return el.value || ''; });
                                      info.checked_selectKomaNo = checkedInputs.filter(function(el){ return el.name === 'selectKomaNo'; }).map(function(el){ return el.value || ''; });
                                      info.checked_selectStime = checkedInputs.filter(function(el){ return el.name === 'selectStime'; }).map(function(el){ return el.value || ''; });
                                      info.checked_selectEtime = checkedInputs.filter(function(el){ return el.name === 'selectEtime'; }).map(function(el){ return el.value || ''; });
                                      info.checked_selectField = checkedInputs.filter(function(el){ return el.name === 'selectField'; }).map(function(el){ return el.value || ''; });
                                      info.selection_applied_to_form =
                                        info.checked_selectUseYMD.indexOf(ymd) >= 0 &&
                                        info.checked_selectStime.indexOf(stime) >= 0 &&
                                        (!etime || info.checked_selectEtime.indexOf(etime) >= 0);
                                      info.target_cell_outer_html = td.outerHTML || '';
                                    }catch(e){}
                                    return info;
                                }
                            }catch(e){ }
                        }
                    }
                    info.reason='no matching stime in cells';
                    return info;
                } catch(e) { info.reason='exception:'+e; return info; }
            })(arguments[0], arguments[1], arguments[2], arguments[3]);
            """
            try:
                clean_js = js.rstrip()
                if clean_js.endswith(";"):
                    clean_js = clean_js[:-1]
                json_js = "return JSON.stringify(" + clean_js + ");"
                info_str = self.navigation_service.execute_script(
                    driver, json_js, ymd, stime, etime, field
                )
                import json as _json

                try:
                    info = _json.loads(info_str) if info_str else None
                except Exception:
                    info = None
                print(f"[debug] select attempt for {slot}: {info}")
                self._last_slot_select_info = info or {}
                ok = bool(info and info.get("selection_applied_to_form", False))
            except Exception as exc:
                print(f"[debug] execute_script failed: {exc}")
                self._last_slot_select_info = {"reason": f"execute_script failed: {exc}"}
                ok = False
            result[slot] = ok

        if submit and any(result.values()):
            self.submit_selected_slots(
                driver,
                success_count=sum(1 for value in result.values() if value),
                wait_alert_seconds=wait_alert_seconds,
            )

        return result

    def select_single_slot(self, driver, slot_text):
        if not slot_text:
            return False
        self.navigation_service.clear_lottery_selection(driver)
        result = self.auto_select_and_submit_slots(
            driver,
            [slot_text],
            submit=False,
        )
        return bool(result.get(slot_text))

    def build_slot_text(self, expected_slot):
        if expected_slot is None:
            return ""
        date = str(self._slot_value(expected_slot, "date", "")).replace("-", "")
        time_range = str(self._slot_value(expected_slot, "time_range", "")).strip()
        field = str(
            self._slot_value(expected_slot, "field_number", "")
            or self._slot_value(expected_slot, "field", "")
        ).strip()
        applied = self._slot_value(expected_slot, "current_entry_count")
        parts = [date]
        if time_range:
            parts.append(time_range.replace(":", ""))
        if field:
            parts.append(f"fields:{field}")
        if applied not in (None, ""):
            parts.append(f"applied:{applied}")
        raw_text = str(self._slot_value(expected_slot, "raw_text", "")).strip()
        return raw_text or " ".join(parts).strip()

    def save_before_apply_debug(self, driver, account_index, entry_index):
        return self._save_submission_debug(
            driver,
            prefix=f"lottery_before_apply_account{account_index}_entry{entry_index}",
        )

    def capture_lottery_selection_state(
        self,
        driver,
        raw_text,
        selected,
        include_last_info=True,
    ):
        state = {
            "raw_text": raw_text,
            "selected": bool(selected),
            "selectFieldCnt": "",
            "selected_slots_count": 0,
            "checked_input_count": 0,
            "checked_inputs": [],
            "checked_selectUseYMD": [],
            "checked_selectKomaNo": [],
            "checked_selectStime": [],
            "checked_selectEtime": [],
            "checked_selectField": [],
            "clicked_element_tag": "",
            "clicked_element_text": "",
            "clicked_element_outer_html": "",
            "target_cell_outer_html": "",
            "selection_applied_to_form": False,
            "current_bname_value": "",
            "current_bname_text": "",
            "current_iname_value": "",
            "current_iname_text": "",
            "selectBldGrpCd": "",
            "selectInstGrpCd": "",
            "before_apply_hidden_values": {},
            "current_park_name": "",
            "current_facility_name": "",
        }
        try:
            captured = self.navigation_service.execute_script(
                driver,
                """
                const checked = (selector) =>
                  Array.from(document.querySelectorAll(selector)).filter((el) => el.checked);
                const checkedInputs = checked('input[type="checkbox"]');
                return {
                  selectFieldCnt: document.getElementById("selectFieldCnt")
                    ? document.getElementById("selectFieldCnt").value || ""
                    : "",
                  current_bname_value: (document.getElementById("bname") || {}).value || "",
                  current_bname_text:
                    (document.getElementById("bname") && document.getElementById("bname").selectedIndex >= 0 && document.getElementById("bname").options[document.getElementById("bname").selectedIndex])
                      ? (document.getElementById("bname").options[document.getElementById("bname").selectedIndex].textContent || "").trim()
                      : "",
                  current_iname_value: (document.getElementById("iname") || {}).value || "",
                  current_iname_text:
                    (document.getElementById("iname") && document.getElementById("iname").selectedIndex >= 0 && document.getElementById("iname").options[document.getElementById("iname").selectedIndex])
                      ? (document.getElementById("iname").options[document.getElementById("iname").selectedIndex].textContent || "").trim()
                      : "",
                  selectBldGrpCd:
                    (document.querySelector('input[name="selectBldGrpCd"]') || document.getElementById("selectBldGrpCd"))
                      ? ((document.querySelector('input[name="selectBldGrpCd"]') || document.getElementById("selectBldGrpCd")).value || "")
                      : "",
                  selectInstGrpCd:
                    (document.querySelector('input[name="selectInstGrpCd"]') || document.getElementById("selectInstGrpCd"))
                      ? ((document.querySelector('input[name="selectInstGrpCd"]') || document.getElementById("selectInstGrpCd")).value || "")
                      : "",
                  selected_slots_count: checked('input[name="selectStime"]').length,
                  checked_input_count: checkedInputs.length,
                  checked_inputs: checkedInputs.map((el) => ({
                    name: el.getAttribute("name") || "",
                    value: el.value || "",
                  })),
                  checked_selectUseYMD: checked('input[name="selectUseYMD"]').map((el) => el.value || ""),
                  checked_selectKomaNo: checked('input[name="selectKomaNo"]').map((el) => el.value || ""),
                  checked_selectStime: checked('input[name="selectStime"]').map((el) => el.value || ""),
                  checked_selectEtime: checked('input[name="selectEtime"]').map((el) => el.value || ""),
                  checked_selectField: checked('input[name="selectField"]').map((el) => el.value || ""),
                };
                """,
            )
            state.update(captured or {})
        except Exception:
            pass
        state["current_park_name"] = state.get("current_bname_text", "")
        state["current_facility_name"] = state.get("current_iname_text", "")
        state["before_apply_hidden_values"] = {
            "selectBldGrpCd": state.get("selectBldGrpCd", ""),
            "selectInstGrpCd": state.get("selectInstGrpCd", ""),
        }
        if include_last_info and self._last_slot_select_info:
            state.update(self._last_slot_select_info)
        return state

    def _slot_value(self, slot, key, default=""):
        if slot is None:
            return default
        if isinstance(slot, dict):
            return slot.get(key, default)
        return getattr(slot, key, default)

    def build_park_facility_validation(
        self,
        expected_slot,
        selection_state=None,
        confirm_page=None,
        validation_source="selection",
    ):
        selection_state = selection_state if isinstance(selection_state, dict) else {}
        confirm_page = confirm_page if isinstance(confirm_page, dict) else {}
        current_state = selection_state.get("current")
        if not isinstance(current_state, dict) or not current_state:
            current_state = selection_state
        expected_park_name = str(self._slot_value(expected_slot, "park_name", "")).strip()
        expected_facility_name = str(self._slot_value(expected_slot, "facility_name", "")).strip()
        expected_date = str(self._slot_value(expected_slot, "date", "")).strip()
        expected_time_range = str(self._slot_value(expected_slot, "time_range", "")).strip()
        expected_field_number = str(
            self._slot_value(expected_slot, "field_number", "")
            or self._slot_value(expected_slot, "field", "")
        ).strip()
        actual_slot_park_name = str(current_state.get("current_bname_text", "")).strip()
        actual_slot_facility_name = str(current_state.get("current_iname_text", "")).strip()
        current_bname_value = str(current_state.get("current_bname_value", "")).strip()
        current_bname_text = str(current_state.get("current_bname_text", "")).strip()
        current_iname_value = str(current_state.get("current_iname_value", "")).strip()
        current_iname_text = str(current_state.get("current_iname_text", "")).strip()
        select_bld_grp_cd = str(current_state.get("selectBldGrpCd", "")).strip()
        select_inst_grp_cd = str(current_state.get("selectInstGrpCd", "")).strip()
        confirm_park_name = str(confirm_page.get("confirm_park_name", "")).strip()
        confirm_facility_name = str(confirm_page.get("confirm_facility_name", "")).strip()
        mismatch_reason = ""
        park_facility_match = True
        if expected_park_name:
            if not actual_slot_park_name:
                park_facility_match = False
                mismatch_reason = (
                    f"park missing: expected={expected_park_name} actual="
                )
            elif expected_park_name not in actual_slot_park_name and actual_slot_park_name not in expected_park_name:
                park_facility_match = False
                mismatch_reason = (
                    f"park mismatch: expected={expected_park_name} actual={actual_slot_park_name}"
                )
        if park_facility_match and expected_facility_name:
            if not actual_slot_facility_name:
                park_facility_match = False
                mismatch_reason = (
                    f"facility missing: expected={expected_facility_name} actual="
                )
            elif expected_facility_name not in actual_slot_facility_name and actual_slot_facility_name not in expected_facility_name:
                park_facility_match = False
                mismatch_reason = (
                    f"facility mismatch: expected={expected_facility_name} actual={actual_slot_facility_name}"
                )
        if confirm_page:
            if expected_park_name and not confirm_park_name:
                park_facility_match = False
                mismatch_reason = f"confirm park missing: expected={expected_park_name}"
            elif (
                expected_park_name
                and confirm_park_name
                and expected_park_name not in confirm_park_name
                and confirm_park_name not in expected_park_name
            ):
                park_facility_match = False
                mismatch_reason = (
                    f"confirm park mismatch: expected={expected_park_name} actual={confirm_park_name}"
                )
            if park_facility_match and expected_facility_name and not confirm_facility_name:
                park_facility_match = False
                mismatch_reason = f"confirm facility missing: expected={expected_facility_name}"
            elif (
                park_facility_match
                and expected_facility_name
                and confirm_facility_name
                and expected_facility_name not in confirm_facility_name
                and confirm_facility_name not in expected_facility_name
            ):
                park_facility_match = False
                mismatch_reason = (
                    f"confirm facility mismatch: expected={expected_facility_name} actual={confirm_facility_name}"
                )
        if park_facility_match:
            if expected_park_name == "府中の森公園" and select_bld_grp_cd != "1301270":
                park_facility_match = False
                mismatch_reason = (
                    f"hidden park mismatch: expected=1301270 actual={select_bld_grp_cd}"
                )
            if park_facility_match and expected_facility_name == "テニス（人工芝）" and select_inst_grp_cd != "12700020":
                park_facility_match = False
                mismatch_reason = (
                    f"hidden facility mismatch: expected=12700020 actual={select_inst_grp_cd}"
                )
        status = "matched" if park_facility_match else "park_mismatch"
        if not park_facility_match and "facility mismatch" in mismatch_reason:
            status = "facility_mismatch"
        if not park_facility_match and "confirm" in mismatch_reason:
            status = "confirm_park_or_facility_mismatch"
        if not park_facility_match and "hidden" in mismatch_reason:
            status = "park_facility_selection_invalid"
        return {
            "validation_source": validation_source,
            "expected_park_name": expected_park_name,
            "expected_facility_name": expected_facility_name,
            "date": expected_date,
            "time_range": expected_time_range,
            "field_number": expected_field_number,
            "actual_slot_park_name": actual_slot_park_name,
            "actual_slot_facility_name": actual_slot_facility_name,
            "current_bname_value": current_bname_value,
            "current_bname_text": current_bname_text,
            "current_iname_value": current_iname_value,
            "current_iname_text": current_iname_text,
            "selectBldGrpCd": select_bld_grp_cd,
            "selectInstGrpCd": select_inst_grp_cd,
            "before_apply_hidden_values": {
                "selectBldGrpCd": select_bld_grp_cd,
                "selectInstGrpCd": select_inst_grp_cd,
            },
            "confirm_park_name": confirm_park_name,
            "confirm_facility_name": confirm_facility_name,
            "park_facility_match": bool(park_facility_match),
            "mismatch_reason": mismatch_reason,
            "status": status,
        }

    def submit_single_selected_slot(
        self,
        driver,
        apply_no,
        expected_slot=None,
        account_index=None,
        entry_index=None,
        select_result=None,
        manual_final_submit=False,
        manual_preconfirm_submit=False,
        wait_alert_seconds=10,
    ):
        summary = {
            "requested_count": 1,
            "submitted_count": 0,
            "completed": False,
            "stopped": False,
            "recovery_triggered": False,
            "recovery_completed": False,
            "recovery_attempts": 0,
            "states": [],
            "debug_files": [],
            "apply_no": apply_no,
            "validation": {},
            "go_to_confirm": {},
            "confirm_page": {},
            "submit_result": {
                "alert_text": "",
                "completion_detected": False,
                "recaptcha_detected": False,
            },
            "manual_final_submit_enabled": bool(manual_final_submit),
            "manual_preconfirm_submit_enabled": bool(manual_preconfirm_submit),
            "recaptcha_recovery": {
                "attempted": False,
                "retry_count": 0,
            },
            "error_message": None,
        }
        if not driver:
            return summary

        prefix_base = "lottery"
        if account_index is not None and entry_index is not None:
            prefix_base = f"lottery_account{account_index}_entry{entry_index}"

        try:
            target_slot_text = self.build_slot_text(expected_slot)
            fresh_select_result = select_result if isinstance(select_result, dict) else {}
            reselected = False
            if not self._selection_ready_for_submit(fresh_select_result, expected_slot):
                fresh_select_result = self.capture_lottery_selection_state(
                    driver,
                    target_slot_text,
                    selected=False,
                    include_last_info=False,
                )
            if not self._selection_ready_for_submit(fresh_select_result, expected_slot):
                reselected = self.select_single_slot(driver, target_slot_text)
                fresh_select_result = self.capture_lottery_selection_state(
                    driver,
                    target_slot_text,
                    selected=bool(reselected),
            )
            summary["select_result_before_apply"] = fresh_select_result
            validation = self.build_park_facility_validation(
                expected_slot,
                selection_state=fresh_select_result,
                validation_source="before_apply",
            )
            summary["validation"] = validation
            summary["debug_files"].extend(
                self._save_submission_debug(
                    driver,
                    prefix=f"lottery_before_apply_account{account_index}_entry{entry_index}",
                )
            )
            self.logger.info(
                "before_btn_go selection_ready=%s reselected=%s target_slot=%s expected_park=%s expected_facility=%s actual_park=%s actual_facility=%s current_bname=%s current_iname=%s selectBldGrpCd=%s selectInstGrpCd=%s validation_source=%s checked_input_count=%s checked_inputs=%s selection_applied_to_form=%s displayNo=%s headers=%s checked_selectUseYMD=%s checked_selectStime=%s checked_selectEtime=%s checked_selectField=%s",
                self._selection_ready_for_submit(fresh_select_result, expected_slot),
                bool(reselected),
                target_slot_text,
                validation.get("expected_park_name"),
                validation.get("expected_facility_name"),
                validation.get("actual_slot_park_name"),
                validation.get("actual_slot_facility_name"),
                validation.get("current_bname_text"),
                validation.get("current_iname_text"),
                validation.get("selectBldGrpCd"),
                validation.get("selectInstGrpCd"),
                validation.get("validation_source"),
                fresh_select_result.get("checked_input_count"),
                json.dumps(fresh_select_result.get("checked_inputs", []), ensure_ascii=False),
                fresh_select_result.get("selection_applied_to_form"),
                fresh_select_result.get("current_display_no"),
                fresh_select_result.get("headers"),
                fresh_select_result.get("checked_selectUseYMD"),
                fresh_select_result.get("checked_selectStime"),
                fresh_select_result.get("checked_selectEtime"),
                fresh_select_result.get("checked_selectField"),
            )
            if not validation.get("park_facility_match", True):
                summary["stopped"] = True
                summary["status"] = validation.get("status", "park_mismatch")
                summary["error_message"] = validation.get("mismatch_reason") or "park/facility mismatch"
                summary["debug_files"].extend(
                    self._save_submission_debug(
                        driver,
                        prefix=f"lottery_park_mismatch_account{account_index}_entry{entry_index}",
                    )
                )
                return summary
            if not self._selection_ready_for_submit(fresh_select_result, expected_slot):
                summary["stopped"] = True
                summary["status"] = "slot_selection_not_applied"
                summary["error_message"] = "slot selection was not applied to submit form"
                return summary
            if manual_preconfirm_submit:
                self.logger.info("manual_preconfirm_submit_enabled=True")
                self.logger.info("waiting_manual_preconfirm_submit")
                self._prompt_manual_preconfirm_submit(driver)
                self.logger.info("manual_preconfirm_submit_confirmed")
                go_to_confirm = {
                    "success": False,
                    "method": "manual_click",
                    "pre_click": self.navigation_service.execute_script(
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
                    ),
                    "post_click": {},
                }
                confirm_reached = self._wait_for_confirmation_page(
                    driver,
                    wait_alert_seconds=wait_alert_seconds,
                )
                try:
                    go_to_confirm["post_click"] = self.navigation_service.inspect_page_state(
                        driver
                    )
                    go_to_confirm["success"] = bool(confirm_reached)
                except Exception:
                    pass
            else:
                go_to_confirm = self.navigation_service.go_to_temp_apply(driver)
                confirm_reached = self._wait_for_confirmation_page(
                    driver,
                    wait_alert_seconds=wait_alert_seconds,
                )
            summary["go_to_confirm"] = go_to_confirm or {}
            pre_click = (go_to_confirm or {}).get("pre_click", {})
            btn_meta = pre_click.get("btn_go") or {}
            self.logger.info(
                "btn_go pre_click selectFieldCnt=%s displayed=%s enabled=%s onclick=%s method=%s",
                pre_click.get("selectFieldCnt"),
                btn_meta.get("displayed"),
                btn_meta.get("enabled"),
                btn_meta.get("onclick"),
                (go_to_confirm or {}).get("method"),
            )
            confirm_page = self.inspect_confirmation_page(driver)
            summary["confirm_page"] = confirm_page
            validation = self.build_park_facility_validation(
                expected_slot,
                selection_state=fresh_select_result,
                confirm_page=confirm_page,
                validation_source="confirm_page",
            )
            summary["validation"] = validation
            self.logger.info(
                "btn_go post_click displayNo=%s title=%s url=%s has_apply=%s alert_text=%s confirm_reached=%s confirm_park_name=%s confirm_facility_name=%s park_facility_match=%s mismatch_reason=%s validation_source=%s",
                (go_to_confirm or {}).get("post_click", {}).get("display_no"),
                (go_to_confirm or {}).get("post_click", {}).get("title"),
                (go_to_confirm or {}).get("post_click", {}).get("current_url"),
                (go_to_confirm or {}).get("post_click", {}).get("has_apply"),
                (go_to_confirm or {}).get("post_click", {}).get("alert_text"),
                confirm_reached,
                validation.get("confirm_park_name"),
                validation.get("confirm_facility_name"),
                validation.get("park_facility_match"),
                validation.get("mismatch_reason"),
                validation.get("validation_source"),
            )
            if not validation.get("park_facility_match", True):
                summary["stopped"] = True
                summary["status"] = validation.get(
                    "status", "confirm_park_or_facility_mismatch"
                )
                summary["error_message"] = (
                    validation.get("mismatch_reason")
                    or "confirmation page park/facility mismatch"
                )
                summary["debug_files"].extend(
                    self._save_submission_debug(
                        driver,
                        prefix=f"lottery_park_mismatch_account{account_index}_entry{entry_index}",
                    )
                )
                return summary
            if manual_preconfirm_submit and confirm_reached:
                self.logger.info("manual_preconfirm_submit_completed")
            summary["debug_files"].extend(
                self._save_submission_debug(
                    driver,
                    prefix=f"lottery_confirm_account{account_index}_entry{entry_index}",
                )
            )
            if not confirm_reached:
                summary["stopped"] = True
                alert_text = ((go_to_confirm or {}).get("post_click") or {}).get("alert_text", "")
                if "利用時間帯を選択して下さい" in alert_text:
                    summary["status"] = "slot_selection_not_applied"
                    summary["error_message"] = alert_text
                else:
                    summary["status"] = "confirm_page_not_reached"
                    summary["error_message"] = "confirmation page was not reached after btn-go"
                summary["debug_files"].extend(
                    self._save_submission_debug(
                        driver,
                        prefix=f"lottery_confirm_not_reached_account{account_index}_entry{entry_index}",
                    )
                )
                return summary
            if not self._validate_confirmation_page(
                summary,
                confirm_page=confirm_page,
                expected_slot=expected_slot,
                expected_apply_no=apply_no,
                driver=driver,
                prefix_base=prefix_base,
            ):
                return summary
            self._restore_apply_selection(driver, apply_no)
            selected_apply = self._get_selected_apply_state(driver)
            summary["confirm_page"]["selected_apply_value"] = selected_apply.get("value")
            summary["confirm_page"]["selected_apply_text"] = selected_apply.get("text")
            self.sleep_func(0.2)
            submission_result = self._submit_with_recovery(
                driver,
                sel_val=apply_no,
                manual_final_submit=manual_final_submit,
                wait_alert_seconds=wait_alert_seconds,
            )
            summary["states"].extend(submission_result.get("states", []))
            summary["debug_files"].extend(submission_result.get("debug_files", []))
            summary["recovery_attempts"] = submission_result.get("recovery_attempts", 0)
            summary["recovery_triggered"] = submission_result.get(
                "recovery_triggered", False
            )
            summary["recovery_completed"] = submission_result.get(
                "recovery_completed", False
            )
            summary["completed"] = submission_result.get("completed", False)
            summary["submit_result"]["alert_text"] = submission_result.get(
                "last_alert_text", ""
            )
            summary["submit_result"]["completion_detected"] = submission_result.get(
                "completed", False
            )
            summary["submit_result"]["recaptcha_detected"] = submission_result.get(
                "recaptcha_detected", False
            )
            summary["submit_result"]["recaptcha_detected_immediately_after_alert"] = submission_result.get(
                "recaptcha_detected_immediately_after_alert", False
            )
            summary["submit_result"]["recaptcha_manual_prompt_shown"] = submission_result.get(
                "recaptcha_manual_prompt_shown", False
            )
            summary["submit_result"]["recaptcha_user_confirmed"] = submission_result.get(
                "recaptcha_user_confirmed", False
            )
            summary["recaptcha_recovery"]["attempted"] = submission_result.get(
                "recovery_triggered", False
            )
            summary["recaptcha_recovery"]["retry_count"] = submission_result.get(
                "recovery_attempts", 0
            )
            summary["recaptcha_recovery"]["response_length"] = submission_result.get(
                "recaptcha_response_length", 0
            )
            summary["recaptcha_recovery"]["response_present"] = submission_result.get(
                "recaptcha_response_present", False
            )
            summary["recaptcha_recovery"]["apply_value_before_restore_after_recaptcha"] = submission_result.get(
                "apply_value_before_restore_after_recaptcha", ""
            )
            summary["recaptcha_recovery"]["apply_restored_after_recaptcha"] = submission_result.get(
                "apply_restored_after_recaptcha", False
            )
            summary["recaptcha_recovery"]["apply_value_after_restore_after_recaptcha"] = submission_result.get(
                "apply_value_after_restore_after_recaptcha", ""
            )
            summary["recaptcha_recovery"]["final_apply_retried_after_recaptcha"] = submission_result.get(
                "final_apply_retried_after_recaptcha", False
            )
            summary["recaptcha_recovery"]["post_alert_monitor_state"] = submission_result.get(
                "post_alert_monitor_state", {}
            )
            summary["submitted_count"] = 1 if summary["completed"] else 0
            if summary["completed"]:
                summary["status"] = "completed"
                summary["debug_files"].extend(
                    self._save_submission_debug(
                        driver,
                        prefix=f"lottery_complete_account{account_index}_entry{entry_index}",
                    )
                )
            else:
                summary["status"] = "submission_incomplete"
                summary["error_message"] = "submission completion was not detected"
            return summary
        except Exception:
            self.logger.exception("Error during single lottery submission: %s", apply_no)
            summary["error_message"] = "exception during single submission"
            summary["status"] = "submission_failed"
            summary["debug_files"].extend(
                self._save_submission_debug(
                    driver,
                    prefix=f"lottery_failure_account{account_index}_entry{entry_index}",
                )
            )
            return summary

    def continue_after_lottery_completion(self, driver):
        return self.navigation_service.continue_lottery_entry_from_complete(driver)

    def submit_selected_slots(
        self,
        driver,
        success_count,
        wait_alert_seconds=10,
    ):
        """Submit already selected lottery slots from the current page."""
        summary = {
            "requested_count": int(success_count),
            "submitted_count": 0,
            "completed": False,
            "recovery_triggered": False,
            "recovery_completed": False,
            "recovery_attempts": 0,
            "states": [],
            "debug_files": [],
        }
        if not driver or success_count <= 0:
            return summary

        try:
            try:
                self.navigation_service.go_to_temp_apply(driver)
            except Exception:
                try:
                    driver.find_element(By.ID, "btn-go").click()
                except Exception:
                    pass

            applied = 0
            for _ in range(success_count):
                try:
                    try:
                        self._get_wait(driver, wait_alert_seconds).until(
                            lambda d: "申込内容確認画面" in d.title
                        )
                    except Exception:
                        self.sleep_func(0.5)

                    applied += 1
                    sel_val = f"{applied}-1"
                    try:
                        self._restore_apply_selection(driver, sel_val)
                    except Exception:
                        debug_context = self._capture_apply_debug_context(driver)
                        self.logger.warning(
                            "apply select not found to set %s: %s",
                            sel_val,
                            json.dumps(debug_context, ensure_ascii=False),
                        )
                        summary["debug_files"].extend(
                            self._save_submission_debug(
                                driver,
                                prefix=f"lottery_apply_select_missing_{sel_val.replace('-', '_')}",
                            )
                        )
                    self.sleep_func(0.2)
                    submission_result = self._submit_with_recovery(
                        driver,
                        sel_val=sel_val,
                        wait_alert_seconds=wait_alert_seconds,
                    )
                    summary["states"].extend(submission_result.get("states", []))
                    summary["debug_files"].extend(
                        submission_result.get("debug_files", [])
                    )
                    summary["recovery_attempts"] += submission_result.get(
                        "recovery_attempts", 0
                    )
                    if submission_result.get("recovery_triggered"):
                        summary["recovery_triggered"] = True
                    if submission_result.get("recovery_completed"):
                        summary["recovery_completed"] = True
                    if not submission_result.get("completed"):
                        break
                    summary["submitted_count"] = applied
                except Exception:
                    self.logger.exception("Error during confirmation/apply loop")
                    debug_files = self._save_submission_debug(
                        driver,
                        prefix=f"lottery_submission_failure_{applied}",
                    )
                    summary["debug_files"].extend(debug_files)
                    break
        except Exception:
            self.logger.exception("Error during submit")

        summary["completed"] = summary["submitted_count"] == summary["requested_count"]
        return summary

    def _submit_with_recovery(
        self,
        driver,
        sel_val,
        manual_final_submit=False,
        wait_alert_seconds=10,
    ):
        result = {
            "completed": False,
            "recovery_triggered": False,
            "recovery_completed": False,
            "recovery_attempts": 0,
            "recaptcha_detected": False,
            "recaptcha_detected_immediately_after_alert": False,
            "recaptcha_manual_prompt_shown": False,
            "recaptcha_user_confirmed": False,
            "recaptcha_response_length": 0,
            "recaptcha_response_present": False,
            "recaptcha_response_length_before_retry": 0,
            "apply_value_before_restore_after_recaptcha": "",
            "apply_restored_after_recaptcha": False,
            "apply_value_after_restore_after_recaptcha": "",
            "final_apply_retried_after_recaptcha": False,
            "completion_detected_after_recaptcha": False,
            "completion_detected_without_recaptcha": False,
            "last_alert_text": "",
            "states": [],
            "debug_files": [],
            "manual_captcha_handled": False,
            "post_alert_monitor_state": {},
        }

        def _log_post_alert_snapshot(label):
            snapshot = {
                "label": label,
                "current_url": "",
                "form_action": "",
                "displayNo": "",
                "loadmsg_visible": False,
                "recaptcha_iframe_count": 0,
                "g_recaptcha_response_length": 0,
            }
            try:
                snapshot.update(
                    self.navigation_service.execute_script(
                        driver,
                        """
                        const form = document.form1 || document.querySelector('form');
                        const displayNo = document.querySelector('input[name="displayNo"]');
                        const loadmsg = document.getElementById('loadmsg');
                        const gToken = document.querySelector("textarea[name='g-recaptcha-response'], textarea#g-recaptcha-response");
                        const iframes = Array.from(document.querySelectorAll("iframe[src*='recaptcha'], iframe[title*='reCAPTCHA'], div.g-recaptcha"));
                        return {
                          current_url: window.location.href || '',
                          form_action: form ? (form.getAttribute('action') || form.action || '') : '',
                          displayNo: displayNo ? (displayNo.value || '') : '',
                          loadmsg_visible: !!loadmsg && loadmsg.style.display !== 'none' && getComputedStyle(loadmsg).display !== 'none',
                          recaptcha_iframe_count: iframes.length,
                          g_recaptcha_response_length: gToken ? (gToken.value || '').length : 0,
                        };
                        """,
                    )
                )
            except Exception:
                pass
            self.logger.info(
                "post_alert_snapshot=%s",
                json.dumps(snapshot, ensure_ascii=False),
            )
            return snapshot

        if manual_final_submit:
            self.logger.info("manual_final_submit_enabled=True")
            self.logger.info("waiting_manual_final_submit")
            self._prompt_manual_final_submit(driver)
            self.logger.info("manual_final_submit_confirmed")
        else:
            self._trigger_final_apply(driver)

        accepted = self._accept_submission_alerts(driver, result, timeout=max(wait_alert_seconds, 60))
        if accepted:
            _log_post_alert_snapshot("accept_alert_immediate")
            self.sleep_func(3)
            _log_post_alert_snapshot("accept_alert_after_3s")
            self.sleep_func(7)
            _log_post_alert_snapshot("accept_alert_after_10s")
            if self._wait_for_completion_after_alert(driver):
                result["completed"] = True
                if manual_final_submit:
                    self.logger.info("manual_final_submit_completed")
                return result

        deadline = time.time() + max(wait_alert_seconds, 1) * 3
        while time.time() < deadline:
            state = self._classify_submission_state(driver)
            result["states"].append(state)

            if state.get("state") == "completed":
                result["completed"] = True
                result["completion_detected_without_recaptcha"] = True
                if manual_final_submit:
                    self.logger.info("manual_final_submit_completed")
                return result

            if state.get("state") == "alert":
                try:
                    alert = driver.switch_to.alert
                    alert_text = alert.text
                    result["last_alert_text"] = alert_text
                    alert.accept()
                    self.logger.info("Accepted confirmation alert: %s", alert_text)
                except Exception:
                    pass
                self.sleep_func(0.3)
                continue

            if state.get("state") == "recaptcha":
                if result.get("manual_captcha_handled"):
                    self.sleep_func(0.3)
                    continue
                result["recaptcha_detected"] = True
                result["recaptcha_detected_immediately_after_alert"] = True
                manual_state = self.login_service.wait_for_manual_captcha(
                    driver,
                    return_state=True,
                )
                result["recaptcha_manual_prompt_shown"] = bool(
                    manual_state.get("prompt_shown", False)
                )
                result["recaptcha_user_confirmed"] = bool(
                    manual_state.get("confirmed", False)
                )
                result["recaptcha_response_length"] = int(
                    manual_state.get("response_length", 0) or 0
                )
                result["recaptcha_response_present"] = bool(
                    manual_state.get("response_present", False)
                )
                self.logger.info(
                    "recaptcha_manual_prompt_shown=%s recaptcha_user_confirmed=%s recaptcha_response_length=%s",
                    result["recaptcha_manual_prompt_shown"],
                    result["recaptcha_user_confirmed"],
                    result["recaptcha_response_length"],
                )
                if not result["recaptcha_user_confirmed"]:
                    self.logger.warning(
                        "User cancelled captcha handling during recovery for %s", sel_val
                    )
                    result["debug_files"].extend(
                        self._save_submission_debug(
                            driver,
                            prefix=f"lottery_submission_recaptcha_cancelled_{sel_val.replace('-', '_')}",
                        )
                    )
                    return result

                result["manual_captcha_handled"] = True
                result["recovery_triggered"] = True
                result["recovery_attempts"] += 1
                try:
                    current_apply = self._get_selected_apply_state(driver)
                except Exception:
                    current_apply = {"value": "", "text": ""}
                result["apply_value_before_restore_after_recaptcha"] = current_apply.get("value", "")
                restore_ok = self._restore_apply_selection(driver, sel_val)
                result["apply_restored_after_recaptcha"] = bool(restore_ok)
                try:
                    restored_apply = self._get_selected_apply_state(driver)
                except Exception:
                    restored_apply = {"value": "", "text": ""}
                result["apply_value_after_restore_after_recaptcha"] = restored_apply.get("value", "")
                self.logger.info(
                    "apply_value_before_restore_after_recaptcha=%s apply_restored_after_recaptcha=%s apply_value_after_restore_after_recaptcha=%s",
                    result["apply_value_before_restore_after_recaptcha"],
                    result["apply_restored_after_recaptcha"],
                    result["apply_value_after_restore_after_recaptcha"],
                )
                if not restore_ok or result["apply_value_after_restore_after_recaptcha"] != sel_val:
                    result["debug_files"].extend(
                        self._save_submission_debug(
                            driver,
                            prefix=f"lottery_submission_apply_restore_failed_after_recaptcha_{sel_val.replace('-', '_')}",
                        )
                    )
                    result["debug_files"].extend(
                        self._save_recaptcha_debug(
                            driver,
                            prefix=f"lottery_submission_apply_restore_failed_after_recaptcha_{sel_val.replace('-', '_')}",
                        )
                    )
                    return result

                result["recaptcha_response_length_before_retry"] = result.get(
                    "recaptcha_response_length", 0
                )
                self._trigger_final_apply(driver)
                result["final_apply_retried_after_recaptcha"] = True
                if self._accept_submission_alerts(
                    driver,
                    result,
                    timeout=max(wait_alert_seconds, 60),
                ):
                    _log_post_alert_snapshot("recaptcha_retry_accept_alert_immediate")
                    if self._wait_for_completion_after_alert(driver):
                        result["completed"] = True
                        result["recovery_completed"] = True
                        result["completion_detected_after_recaptcha"] = True
                        if manual_final_submit:
                            self.logger.info("manual_final_submit_completed")
                        return result
                result["recovery_completed"] = True
                continue

            if state.get("state") in {"confirm", "unknown"}:
                self.sleep_func(0.3)
                continue

            if state.get("state") == "error":
                result["debug_files"].extend(
                    self._save_submission_debug(
                        driver,
                        prefix=f"lottery_submission_error_{sel_val.replace('-', '_')}",
                    )
                )
                return result

            self.sleep_func(0.3)
        result["debug_files"].extend(
            self._save_submission_debug(
                driver,
                prefix=f"lottery_submission_timeout_{sel_val.replace('-', '_')}",
            )
        )
        return result

    def _prompt_manual_final_submit(self, driver):
        try:
            self.show_info(
                "手動最終送信",
                "申込みボタンを手動で押してください。\n押したらOKを押してください。",
            )
        except Exception:
            self.logger.info(
                "手動最終送信: 申込みボタンを手動で押してください。押したらOKを押してください。"
            )
        return {}

    def _prompt_manual_preconfirm_submit(self, driver):
        try:
            self.show_info(
                "手動申込みボタン",
                "申込みボタンを手動で押してください。\n押したらOKを押してください。",
            )
        except Exception:
            self.logger.info(
                "手動申込みボタン: 申込みボタンを手動で押してください。押したらOKを押してください。"
            )
        return {}

    def _accept_submission_alerts(self, driver, result, timeout=60):
        accepted = False
        attempt = 1
        wait_seconds = max(timeout, 1)
        while True:
            try:
                self._get_wait(driver, wait_seconds).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_text = alert.text
                result["last_alert_text"] = alert_text
                result["states"].append(
                    {
                        "state": "alert",
                        "alert_text": alert_text,
                        "attempt": attempt,
                    }
                )
                alert.accept()
                accepted = True
                self.logger.info("Accepted confirmation alert: %s", alert_text)
                self.sleep_func(0.3)
                try:
                    if self.login_service.detect_captcha(driver):
                        break
                except Exception:
                    pass
                attempt += 1
                wait_seconds = 2
            except TimeoutException:
                break
            except UnexpectedAlertPresentException:
                try:
                    alert = driver.switch_to.alert
                    alert_text = alert.text
                    result["last_alert_text"] = alert_text
                    alert.accept()
                    accepted = True
                    self.logger.info("Accepted unexpected confirmation alert: %s", alert_text)
                    attempt += 1
                    wait_seconds = 2
                except Exception:
                    break
            except Exception:
                break
        return accepted

    def _wait_for_completion_after_alert(self, driver, timeout=15):
        try:
            self._get_wait(driver, timeout).until(
                lambda d: self._is_lottery_completion_page(d)
            )
            return True
        except Exception:
            return False

    def _wait_for_completion_or_recaptcha_after_alert(
        self,
        driver,
        result,
        sel_val,
        timeout=15,
        recaptcha_enabled=True,
    ):
        def _capture_monitor_state():
            state = {
                "alert_present": False,
                "alert_text": "",
                "recaptcha_visible": False,
                "response_present": False,
                "response_length": 0,
                "current_state": "unknown",
            }
            try:
                alert = driver.switch_to.alert
                state["alert_present"] = True
                state["alert_text"] = alert.text
            except NoAlertPresentException:
                pass
            except Exception:
                pass
            try:
                recaptcha_state = self.login_service.inspect_recaptcha_state(driver)
            except Exception:
                recaptcha_state = {}
            try:
                response_info = self.login_service.get_recaptcha_response_info(driver)
            except Exception:
                response_info = {}
            state["recaptcha_visible"] = bool(recaptcha_state.get("requires_manual", False))
            state["response_length"] = int(response_info.get("response_length", 0) or 0)
            state["response_present"] = bool(response_info.get("response_present", False))
            if state["alert_present"]:
                state["current_state"] = "alert"
            elif state["recaptcha_visible"]:
                state["current_state"] = "recaptcha"
            elif self._is_lottery_completion_page(driver):
                state["current_state"] = "completed"
            return state

        def _log_monitor_state(state):
            self.logger.info(
                "final_apply_monitor_state=%s",
                json.dumps(state, ensure_ascii=False),
            )

        def _accept_current_alert(state):
            try:
                alert = driver.switch_to.alert
                alert.accept()
                result["last_alert_text"] = state.get("alert_text", "")
                self.logger.info("Accepted confirmation alert: %s", state.get("alert_text", ""))
                return True
            except Exception:
                return False

        def _restore_after_captcha():
            try:
                current_apply = self._get_selected_apply_state(driver)
            except Exception:
                current_apply = {"value": "", "text": ""}
            result["apply_value_before_restore_after_recaptcha"] = current_apply.get("value", "")
            restore_ok = self._restore_apply_selection(driver, sel_val)
            result["apply_restored_after_recaptcha"] = bool(restore_ok)
            try:
                restored_apply = self._get_selected_apply_state(driver)
            except Exception:
                restored_apply = {"value": "", "text": ""}
            result["apply_value_after_restore_after_recaptcha"] = restored_apply.get("value", "")
            self.logger.info(
                "apply_value_before_restore_after_recaptcha=%s apply_restored_after_recaptcha=%s apply_value_after_restore_after_recaptcha=%s",
                result["apply_value_before_restore_after_recaptcha"],
                result["apply_restored_after_recaptcha"],
                result["apply_value_after_restore_after_recaptcha"],
            )
            return bool(restore_ok and result["apply_value_after_restore_after_recaptcha"] == sel_val)

        self.logger.info("final_apply_monitoring_started interval=0.3")
        deadline = time.time() + max(timeout, 1)
        monitor_interval = 0.3
        while time.time() < deadline:
            state = _capture_monitor_state()
            result["states"].append(state)
            result["post_alert_monitor_state"] = state
            _log_monitor_state(state)

            if state.get("current_state") == "completed":
                return "completed"

            if state.get("alert_present"):
                _accept_current_alert(state)
                self.sleep_func(monitor_interval)
                continue

            if recaptcha_enabled and state.get("current_state") == "recaptcha":
                if result.get("manual_captcha_handled"):
                    if state.get("response_present") or int(state.get("response_length", 0) or 0) > 0:
                        self.logger.info(
                            "recaptcha_solved_pending_submit response_length=%s response_present=%s",
                            state.get("response_length", 0),
                            state.get("response_present", False),
                        )
                    self.sleep_func(monitor_interval)
                    continue

                result["recaptcha_detected"] = True
                result["recaptcha_detected_immediately_after_alert"] = True
                manual_state = self.login_service.wait_for_manual_captcha(
                    driver,
                    return_state=True,
                )
                result["recaptcha_manual_prompt_shown"] = bool(
                    manual_state.get("prompt_shown", False)
                )
                result["recaptcha_user_confirmed"] = bool(
                    manual_state.get("confirmed", False)
                )
                result["recaptcha_response_length"] = int(
                    manual_state.get("response_length", 0) or 0
                )
                result["recaptcha_response_present"] = bool(
                    manual_state.get("response_present", False)
                )
                self.logger.info(
                    "recaptcha_manual_prompt_shown=%s recaptcha_user_confirmed=%s recaptcha_response_length=%s",
                    result["recaptcha_manual_prompt_shown"],
                    result["recaptcha_user_confirmed"],
                    result["recaptcha_response_length"],
                )
                if not result["recaptcha_user_confirmed"]:
                    self.logger.warning(
                        "User cancelled captcha handling during recovery for %s", sel_val
                    )
                    result["debug_files"].extend(
                        self._save_submission_debug(
                            driver,
                            prefix=f"lottery_submission_recaptcha_cancelled_{sel_val.replace('-', '_')}",
                        )
                    )
                    return "recaptcha_cancelled"

                result["manual_captcha_handled"] = True
                result["recovery_triggered"] = True
                result["recovery_attempts"] += 1
                if not _restore_after_captcha():
                    result["debug_files"].extend(
                        self._save_submission_debug(
                            driver,
                            prefix=f"lottery_submission_apply_restore_failed_after_recaptcha_{sel_val.replace('-', '_')}",
                        )
                    )
                    result["debug_files"].extend(
                        self._save_recaptcha_debug(
                            driver,
                            prefix=f"lottery_submission_apply_restore_failed_after_recaptcha_{sel_val.replace('-', '_')}",
                        )
                    )
                    return "apply_restore_failed_after_recaptcha"

                self.logger.info(
                    "recaptcha_response_length_before_retry=%s",
                    result.get("recaptcha_response_length", 0),
                )
                self._trigger_final_apply(driver)
                self.sleep_func(monitor_interval)
                continue

            self.sleep_func(monitor_interval)

        return "timeout"

    def _is_lottery_completion_page(self, driver):
        try:
            title = driver.title
        except Exception:
            title = ""
        if "抽選メール送信完了画面" in title or "抽選申込み完了" in title:
            return True
        try:
            return bool(
                self.navigation_service.execute_script(
                    driver,
                    """
                    const displayNo = document.querySelector('input[name="displayNo"]');
                    return (
                      (displayNo && displayNo.value === "plwdc1000")
                      || !!document.getElementById("btn-light")
                      || !!document.querySelector("#fin-lottery")
                    );
                    """,
                )
            )
        except Exception:
            return False

    def _wait_for_confirmation_page(self, driver, wait_alert_seconds=10):
        try:
            self._get_wait(driver, wait_alert_seconds).until(
                lambda d: (
                    "申込内容確認画面" in d.title
                    or self.inspect_confirmation_page(d).get("displayNo") == "plwca1000"
                    or bool(d.find_elements(By.ID, "apply"))
                )
            )
        except Exception:
            self.sleep_func(0.5)
        info = self.inspect_confirmation_page(driver)
        return bool(
            "申込内容確認画面" in getattr(driver, "title", "")
            or info.get("displayNo") == "plwca1000"
            or driver.find_elements(By.ID, "apply")
        )

    def _restore_apply_selection(self, driver, sel_val):
        try:
            select = Select(driver.find_element(By.ID, "apply"))
            current_value = select.first_selected_option.get_attribute("value")
            if current_value != sel_val:
                select.select_by_value(sel_val)
            selected = select.first_selected_option
        except Exception:
            return False
        self.logger.info(
            "Restored apply selection: value=%s text=%s title=%s",
            selected.get_attribute("value"),
            selected.text.strip(),
            driver.title,
        )
        return True

    def _trigger_final_apply(self, driver):
        def _snapshot():
            try:
                return self.navigation_service.execute_script(
                    driver,
                    """
                    const button = document.getElementById("btn-go");
                    const displayNo = document.querySelector('input[name="displayNo"]');
                    const apply = document.getElementById("apply");
                    return {
                      display_no: displayNo ? displayNo.value || "" : "",
                      title: document.title || "",
                      url: window.location.href,
                      has_btn_go: !!button,
                      btn_displayed: button ? !!(button.offsetWidth || button.offsetHeight || button.getClientRects().length) : false,
                      btn_enabled: button ? !button.disabled : false,
                      btn_onclick: button ? button.getAttribute("onclick") || "" : "",
                      apply_value: apply ? apply.value || "" : "",
                    };
                    """,
                )
            except Exception:
                return {}

        result = {
            "method_attempted": "",
            "before_state": {},
        }
        try:
            result["before_state"] = _snapshot()
            self.logger.info(
                "trigger_final_apply pre_click state=%s",
                json.dumps(result["before_state"], ensure_ascii=False),
            )
        except Exception:
            self.logger.info("trigger_final_apply pre_click state unavailable")

        try:
            button = driver.find_element(By.ID, "btn-go")
            self.navigation_service.execute_script(
                driver,
                "arguments[0].scrollIntoView({block: 'center'});",
                button,
            )
            result["method_attempted"] = "native_click"
            self.logger.info("final_apply_click_started method=native_click")
            button.click()
            return result
        except Exception as exc:
            self.logger.warning("trigger_final_apply native_click failed: %s", exc)

        return result

    def _capture_apply_debug_context(self, driver):
        context = {
            "title": "",
            "url": "",
            "current_selection": {},
            "apply_options": [],
        }
        try:
            context["title"] = driver.title
            context["url"] = driver.current_url
        except Exception:
            pass
        try:
            context["current_selection"] = self.navigation_service.execute_script(
                driver,
                """
                return {
                  selectFieldCnt: document.getElementById("selectFieldCnt")
                    ? document.getElementById("selectFieldCnt").value || ""
                    : "",
                  display_date: document.getElementById("selectDisplayUseYMD")
                    ? document.getElementById("selectDisplayUseYMD").textContent.replace(/\\s+/g, " ").trim()
                    : "",
                  display_time: document.getElementById("selectDisplayTime")
                    ? document.getElementById("selectDisplayTime").textContent.replace(/\\s+/g, " ").trim()
                    : "",
                };
                """,
            )
        except Exception:
            pass
        try:
            select = Select(driver.find_element(By.ID, "apply"))
            context["apply_options"] = [
                {
                    "value": option.get_attribute("value"),
                    "text": option.text.strip(),
                }
                for option in select.options
            ]
        except Exception:
            pass
        return context

    def _save_recaptcha_debug(self, driver, prefix):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_name = f"{prefix}_{timestamp}.html"
        json_name = f"{prefix}_{timestamp}_dom_summary.json"
        saved = []
        try:
            path = self.navigation_service.save_debug_html(driver, html_name)
            if path is not None:
                saved.append(str(path))
        except Exception:
            self.logger.exception("Failed to save reCAPTCHA debug HTML")
        try:
            path = self.navigation_service.save_dom_summary(driver, json_name)
            if path is not None:
                saved.append(str(path))
        except Exception:
            self.logger.exception("Failed to save reCAPTCHA debug DOM summary")
        return saved

    def inspect_confirmation_page(self, driver):
        info = {
            "displayNo": "",
            "title": "",
            "url": "",
            "use_date": "",
            "use_time": "",
            "confirm_park_name": "",
            "confirm_facility_name": "",
            "apply_options": [],
            "selected_apply_value": "",
            "selected_apply_text": "",
        }
        try:
            info.update(
                self.navigation_service.execute_script(
                    driver,
                    """
                    const text = (value) => (value || "").replace(/\\s+/g, " ").trim();
                    const displayNo = document.querySelector('input[name="displayNo"]');
                    const thRows = Array.from(document.querySelectorAll("#content-table tr"));
                    const summaryRows = Array.from(document.querySelectorAll("table.sp-block-table tbody tr"));
                    const findValue = (label) => {
                      for (const row of thRows) {
                        const th = row.querySelector("th");
                        const td = row.querySelector("td");
                        if (th && td && text(th.textContent) === label) {
                          return text(td.textContent);
                        }
                      }
                      return "";
                    };
                    const findSummaryValue = (label) => {
                      for (const row of summaryRows) {
                        const pairs = Array.from(row.querySelectorAll("td")).map((td) => {
                          const title = td.querySelector(".title-sp");
                          return {
                            label: title ? text(title.textContent).replace(/：$/, "") : "",
                            value: title
                              ? text(td.textContent).replace(text(title.textContent), "").trim()
                              : text(td.textContent),
                          };
                        });
                        for (const pair of pairs) {
                          if (pair.label === label) {
                            return pair.value;
                          }
                        }
                      }
                      return "";
                    };
                    const findAnyValue = (labels) => {
                      const rows = [...thRows, ...summaryRows];
                      for (const label of labels) {
                        for (const row of rows) {
                          const text = row.textContent || "";
                          if (text.indexOf(label) >= 0) {
                            const tds = Array.from(row.querySelectorAll("td"));
                            if (tds.length >= 2) {
                              const candidates = tds
                                .map((td) => (td.textContent || "").replace(/\s+/g, " ").trim())
                                .filter(Boolean);
                              if (candidates.length >= 2) {
                                return candidates[candidates.length - 1];
                              }
                              return candidates[0] || text.replace(/\s+/g, " ").trim();
                            }
                          }
                        }
                      }
                      return "";
                    };
                    const apply = document.getElementById("apply");
                    return {
                      displayNo: displayNo ? (displayNo.value || "") : "",
                      title: document.title || "",
                      url: window.location.href,
                      use_date: findValue("利用日") || findSummaryValue("利用日"),
                      use_time: findValue("利用時間") || findSummaryValue("利用時間"),
                      confirm_park_name:
                        findValue("公園") ||
                        findSummaryValue("公園") ||
                        findAnyValue(["公園名", "利用公園", "予約施設", "施設名", "利用場所", "公園"]),
                      confirm_facility_name:
                        findValue("施設") ||
                        findSummaryValue("施設") ||
                        findAnyValue(["施設名", "利用施設", "利用場所", "会場", "施設"]),
                      apply_options: apply
                        ? Array.from(apply.options).map((opt) => ({
                            value: opt.value || "",
                            text: text(opt.textContent),
                          }))
                        : [],
                      selected_apply_value: apply ? (apply.value || "") : "",
                      selected_apply_text:
                        apply && apply.selectedIndex >= 0 && apply.options[apply.selectedIndex]
                          ? text(apply.options[apply.selectedIndex].textContent)
                          : "",
                    };
                    """,
                )
            )
        except Exception:
            pass
        return info

    def _get_selected_apply_state(self, driver):
        try:
            select = Select(driver.find_element(By.ID, "apply"))
            selected = select.first_selected_option
            return {
                "value": selected.get_attribute("value"),
                "text": selected.text.strip(),
            }
        except Exception:
            return {"value": "", "text": ""}

    def _validate_confirmation_page(
        self,
        summary,
        confirm_page,
        expected_slot,
        expected_apply_no,
        driver,
        prefix_base,
    ):
        self.logger.info(
            "confirm_page displayNo=%s title=%s use_date=%s use_time=%s apply_options=%s selected_apply_value=%s selected_apply_text=%s",
            confirm_page.get("displayNo"),
            confirm_page.get("title"),
            confirm_page.get("use_date"),
            confirm_page.get("use_time"),
            json.dumps(confirm_page.get("apply_options", []), ensure_ascii=False),
            confirm_page.get("selected_apply_value"),
            confirm_page.get("selected_apply_text"),
        )
        if confirm_page.get("displayNo") != "plwca1000":
            summary["stopped"] = True
            if confirm_page.get("displayNo") == "plwba4000":
                summary["status"] = "confirm_page_not_reached"
                summary["error_message"] = "confirmation page was not reached"
            else:
                summary["status"] = "confirm_page_not_reached"
                summary["error_message"] = "confirmation page displayNo mismatch"
        elif not self._matches_confirmation_slot(confirm_page, expected_slot):
            summary["stopped"] = True
            summary["status"] = "confirm_mismatch"
            summary["error_message"] = "confirmation page slot mismatch"
        elif not any(
            option.get("value") == expected_apply_no
            for option in confirm_page.get("apply_options", [])
        ):
            summary["stopped"] = True
            summary["status"] = "apply_option_missing"
            summary["error_message"] = "expected apply_no not found"
        if summary["stopped"]:
            summary["debug_files"].extend(
                self._save_submission_debug(
                    driver,
                    prefix=f"lottery_failure_account{prefix_base.split('_account')[-1]}",
                )
            )
            return False
        return True

    def _matches_confirmation_slot(self, confirm_page, expected_slot):
        if not expected_slot:
            return True
        confirm_date = self._normalize_confirm_date(confirm_page.get("use_date", ""))
        confirm_time = self._normalize_confirm_time(confirm_page.get("use_time", ""))
        expected_date = str(self._slot_value(expected_slot, "date", "")).strip()
        expected_time = str(self._slot_value(expected_slot, "time_range", "")).strip()
        return confirm_date == expected_date and confirm_time == expected_time

    def _normalize_confirm_date(self, text):
        raw = str(text)
        match = re.search(r"(\d{4})年.*?(\d{1,2})月(\d{1,2})日", raw)
        if match:
            year, month, day = match.groups()
            return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
        match = re.search(r"(\d{1,2})月(\d{1,2})日.*?(\d{4})年", raw)
        if not match:
            return ""
        month, day, year = match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    def _normalize_confirm_time(self, text):
        normalized = str(text).replace("：", ":")
        match = re.search(r"(\d{1,2})時(\d{2})分.*?(\d{1,2})時(\d{2})分", normalized)
        if not match:
            match = re.search(r"(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})", normalized)
        if not match:
            return ""
        sh, sm, eh, em = match.groups()
        return f"{int(sh):02d}:{sm}-{int(eh):02d}:{em}"

    def _selection_ready_for_submit(self, select_result, expected_slot):
        if not isinstance(select_result, dict):
            return False
        checked_input_count = int(select_result.get("checked_input_count", 0) or 0)
        if checked_input_count <= 0:
            return False
        if not expected_slot:
            return bool(select_result.get("selection_applied_to_form"))
        expected_date = str(self._slot_value(expected_slot, "date", "")).replace("-", "")
        expected_start = self._normalize_hhmm_for_compare(
            str(self._slot_value(expected_slot, "time_range", "")).split("-")[0]
        )
        expected_end = self._normalize_hhmm_for_compare(
            str(self._slot_value(expected_slot, "time_range", "")).split("-")[-1]
        )
        expected_field = str(
            self._slot_value(expected_slot, "field_number", "")
            or self._slot_value(expected_slot, "field", "")
        ).strip()
        checked_dates = [str(value) for value in select_result.get("checked_selectUseYMD", [])]
        checked_starts = [
            self._normalize_hhmm_for_compare(value)
            for value in select_result.get("checked_selectStime", [])
        ]
        checked_ends = [
            self._normalize_hhmm_for_compare(value)
            for value in select_result.get("checked_selectEtime", [])
        ]
        checked_fields = [
            str(value).strip()
            for value in select_result.get("checked_selectField", [])
        ]
        return (
            bool(select_result.get("selection_applied_to_form"))
            and expected_date in checked_dates
            and expected_start in checked_starts
            and expected_end in checked_ends
            and (not expected_field or expected_field in checked_fields)
        )

    def _normalize_hhmm_for_compare(self, value):
        text = str(value or "").replace(":", "").strip()
        if not text.isdigit():
            return text
        return str(int(text))

    def _wait_for_target_headers_stable(self, driver, ymd, checks=2, timeout=5):
        deadline = time.time() + max(timeout, 1)
        last_headers = []
        consecutive = 0
        while time.time() < deadline:
            try:
                headers = self.navigation_service.execute_script(
                    driver,
                    "return Array.from(document.querySelectorAll('#usedate-table thead input[name=\"selectUseYMD\"]')).map(h => String(h.value || '').trim());",
                ) or []
            except Exception:
                headers = []
            if ymd in headers and headers == last_headers:
                consecutive += 1
                if consecutive >= checks:
                    return headers
            else:
                consecutive = 1 if ymd in headers else 0
            last_headers = headers
            self.sleep_func(0.2)
        return last_headers

    def _classify_submission_state(self, driver):
        title = ""
        current_url = ""
        try:
            title = driver.title
            current_url = driver.current_url
        except Exception:
            pass
        try:
            alert = driver.switch_to.alert
            return {
                "state": "alert",
                "title": title,
                "url": current_url,
                "alert_text": alert.text,
            }
        except NoAlertPresentException:
            pass
        except Exception:
            pass

        try:
            if self.login_service.detect_captcha(driver):
                return {"state": "recaptcha", "title": title, "url": current_url}
        except Exception:
            pass

        if self._is_lottery_completion_page(driver):
            return {"state": "completed", "title": title, "url": current_url}
        if "申込内容確認画面" in title:
            return {"state": "confirm", "title": title, "url": current_url}
        if "エラー" in title or "Error" in title:
            return {"state": "error", "title": title, "url": current_url}
        return {"state": "unknown", "title": title, "url": current_url}

    def _save_submission_debug(self, driver, prefix):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_name = f"{prefix}_{timestamp}.html"
        json_name = f"{prefix}_{timestamp}_dom_summary.json"
        saved = []
        try:
            path = self.navigation_service.save_debug_html(driver, html_name)
            if path is not None:
                saved.append(str(path))
        except Exception:
            self.logger.exception("Failed to save submission debug HTML")
        try:
            path = self.navigation_service.save_dom_summary(driver, json_name)
            if path is not None:
                saved.append(str(path))
        except Exception:
            self.logger.exception("Failed to save submission debug DOM summary")
        return saved

    def check_lottery(self, id_dict, output_csv_path=""):
        """Check current lottery entries and optionally save CSV."""
        reserv_dict = {}
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
                        self.navigation_service.go_to_lottery_cancel_list(driver)
                        self.sleep_func(0.5)
                        soup = bs(driver.page_source, "html.parser")
                        found_day_list = [
                            elem.text
                            for elem in soup.find_all(string=re.compile("月.*日(.*)"))
                        ]
                        found_time_list = [
                            elem.text
                            for elem in soup.find_all(string=re.compile("時.*分"))
                        ]
                        if len(found_day_list) == 2:
                            print(
                                "ID:"
                                + user_id
                                + " 申込み日1→ "
                                + found_day_list[0]
                                + " "
                                + found_time_list[0]
                                + found_time_list[1]
                            )
                            print(
                                "ID:"
                                + user_id
                                + " 申込み日2→ "
                                + found_day_list[1]
                                + " "
                                + found_time_list[2]
                                + found_time_list[3]
                            )
                            reserv_dict[user_id] = [
                                user_values[0],
                                user_values[1],
                                user_values[2],
                                found_day_list[0]
                                + " "
                                + found_time_list[0]
                                + found_time_list[1],
                                found_day_list[1]
                                + " "
                                + found_time_list[2]
                                + found_time_list[3],
                            ]
                        elif len(found_day_list) == 1:
                            print(
                                "ID:"
                                + user_id
                                + " 申込み日1→ "
                                + found_day_list[0]
                                + " "
                                + found_time_list[0]
                                + found_time_list[1]
                            )
                            reserv_dict[user_id] = [
                                user_values[0],
                                user_values[1],
                                user_values[2],
                                found_day_list[0]
                                + " "
                                + found_time_list[0]
                                + found_time_list[1],
                            ]
                        else:
                            print("ID:" + user_id + " 申込みなし")
                            reserv_dict[user_id] = [
                                user_values[0],
                                user_values[1],
                                user_values[2],
                                "",
                                "",
                            ]
                    except UnexpectedAlertPresentException:
                        print("ID:" + user_id + " 申込みなし")
                        reserv_dict[user_id] = [
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
            self.output_id_dict(reserv_dict, output_csv_path)
        return reserv_dict

    def check_result(self, id_dict, output_csv_path=""):
        """Check current lottery results and optionally save CSV."""
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
                            print(
                                "ID:"
                                + user_id
                                + " 当選日1→ "
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
                        elif len(found_day_list) == 2:
                            print(
                                "ID:"
                                + user_id
                                + " 当選日1→ "
                                + found_day_list[0]
                                + " "
                                + found_time_list[0]
                            )
                            print(
                                "ID:"
                                + user_id
                                + " 当選日2→ "
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
                    except UnexpectedAlertPresentException:
                        print("ID:" + user_id + " 申込みなし")
                        continue
                self.sleep_func(1)
                self.navigation_service.logout(driver)
                self.sleep_func(1)
        finally:
            self.browser_session.safe_close(driver)

        if output_csv_path != "":
            self.output_id_dict(result_dict, output_csv_path)

        return result_dict
