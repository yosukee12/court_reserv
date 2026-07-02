# -*- coding: utf-8 -*-
"""Lottery business logic extracted from the legacy UI class."""

import re
import time

from bs4 import BeautifulSoup as bs
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException


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
            parts = slot.split()
            ymd = parts[0] if parts else ""
            stime = ""
            for part in parts:
                if "-" in part and part.split("-")[0].isdigit():
                    stime = part.split("-")[0]
                    break
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
                            break
                        if headers:
                            try:
                                min_h = min(headers)
                                max_h = max(headers)
                                print(f"[debug] min_h={min_h}, max_h={max_h}")
                                if ymd > max_h:
                                    print("[debug] clicking next-week")
                                    try:
                                        self.navigation_service.execute_script(
                                            driver,
                                            "document.getElementById('next-week').click();",
                                        )
                                    except Exception:
                                        try:
                                            self.navigation_service.execute_script(
                                                driver,
                                                "doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
                                            )
                                        except Exception:
                                            print("[debug] next-week click failed")
                                elif ymd < min_h:
                                    print("[debug] clicking last-week")
                                    try:
                                        self.navigation_service.execute_script(
                                            driver,
                                            "document.getElementById('last-week').click();",
                                        )
                                    except Exception:
                                        try:
                                            self.navigation_service.execute_script(
                                                driver,
                                                "doPrevWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
                                            )
                                        except Exception:
                                            print("[debug] last-week click failed")
                                else:
                                    print("[debug] not in range, clicking next-week")
                                    try:
                                        self.navigation_service.execute_script(
                                            driver,
                                            "document.getElementById('next-week').click();",
                                        )
                                    except Exception:
                                        try:
                                            self.navigation_service.execute_script(
                                                driver,
                                                "doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);",
                                            )
                                        except Exception:
                                            print(
                                                "[debug] fallback next-week click failed"
                                            )
                            except Exception as exc:
                                print(f"[debug] header compare failed: {exc}")
                                try:
                                    self.navigation_service.execute_script(
                                        driver,
                                        "document.getElementById('next-week').click();",
                                    )
                                except Exception:
                                    pass
                        self.sleep_func(0.6)
                    print(f"[debug] finished pagination attempts for {ymd}")
                except Exception:
                    pass
            js = """
            (function(ymd, stime) {
                var info = {found:false, idx:-1, clicked:false, reason:''};
                try {
                    var headers = Array.from(document.querySelectorAll('#usedate-table thead input[name="selectUseYMD"]')).map(h=>h.value);
                    for (var i=0;i<headers.length;i++){
                        if (headers[i] === ymd){ info.idx = i; break; }
                    }
                    if (info.idx === -1){ info.reason='ymd not in headers'; return info; }
                    var rows = document.querySelectorAll('#usedate-table tbody tr');
                    for (var r=0;r<rows.length;r++){
                        var tds = rows[r].querySelectorAll('td');
                        var td = tds[info.idx];
                        if (!td) continue;
                        var st = td.querySelector('input[name="selectStime"]');
                        if (st){
                            try{
                                if (parseInt(st.value,10) === parseInt(stime,10)){
                                    info.found = true;
                                    try{ td.click(); info.clicked = true; }catch(e){}
                                    try{
                                        var num = td.querySelector('span.font-weight-bold');
                                        if(num){ num.click(); info.clicked = true; }
                                    }catch(e){}
                                    try{ Array.from(td.querySelectorAll('input[type=checkbox]')).forEach(c=>c.checked=true); }catch(e){}
                                    try{ document.querySelectorAll('#usedate-table thead input[name="selectUseYMD"]')[info.idx].checked = true; }catch(e){}
                                    return info;
                                }
                            }catch(e){ }
                        }
                    }
                    info.reason='no matching stime in cells';
                    return info;
                } catch(e) { info.reason='exception:'+e; return info; }
            })(arguments[0], arguments[1]);
            """
            try:
                clean_js = js.rstrip()
                if clean_js.endswith(";"):
                    clean_js = clean_js[:-1]
                json_js = "return JSON.stringify(" + clean_js + ");"
                info_str = self.navigation_service.execute_script(
                    driver, json_js, ymd, stime
                )
                import json as _json

                try:
                    info = _json.loads(info_str) if info_str else None
                except Exception:
                    info = None
                print(f"[debug] select attempt for {slot}: {info}")
                ok = bool(info and info.get("found", False))
            except Exception as exc:
                print(f"[debug] execute_script failed: {exc}")
                ok = False
            result[slot] = ok

        if submit and any(result.values()):
            self.submit_selected_slots(
                driver,
                success_count=sum(1 for value in result.values() if value),
                wait_alert_seconds=wait_alert_seconds,
            )

        return result

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
                        Select(driver.find_element(By.ID, "apply")).select_by_value(
                            sel_val
                        )
                    except Exception:
                        self.logger.warning("apply select not found to set %s", sel_val)
                    self.sleep_func(0.2)
                    try:
                        try:
                            if self.login_service.detect_captcha(driver):
                                ok = self.login_service.wait_for_manual_captcha(driver)
                                if not ok:
                                    self.logger.warning(
                                        "User cancelled captcha handling; aborting apply loop"
                                    )
                                    break
                        except Exception:
                            pass
                        self.navigation_service.execute_script(
                            driver,
                            "javascript:sendLotApply(document.form1, gLotWInstLotApplyAction, event);",
                        )
                    except Exception:
                        try:
                            driver.find_element(By.ID, "btn-apply").click()
                        except Exception:
                            pass

                    try:
                        self._get_wait(driver, wait_alert_seconds).until(
                            EC.alert_is_present()
                        )
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        alert.accept()
                        self.logger.info(
                            "Accepted confirmation alert: %s", alert_text
                        )
                    except Exception:
                        self.logger.info(
                            "No confirmation alert appeared for apply %s", sel_val
                        )

                    try:
                        self._get_wait(driver, 10).until(
                            lambda d: "抽選メール送信完了画面" in d.title
                        )
                    except Exception:
                        pass
                    summary["submitted_count"] = applied
                except Exception:
                    self.logger.exception("Error during confirmation/apply loop")
                    break
        except Exception:
            self.logger.exception("Error during submit")

        summary["completed"] = summary["submitted_count"] == summary["requested_count"]
        return summary

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
