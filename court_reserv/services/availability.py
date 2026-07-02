# -*- coding: utf-8 -*-
"""Availability business logic extracted from the legacy UI class."""

import calendar
import csv
import datetime
import re
import time

from selenium.webdriver.common.by import By


class AvailabilityService:
    """Encapsulate availability checking and slot collection flows."""

    def __init__(
        self,
        config,
        browser_session,
        navigation_service,
        logger,
        get_debug_output_dir,
        sleep_func=time.sleep,
    ):
        self.config = config
        self.browser_session = browser_session
        self.navigation_service = navigation_service
        self.logger = logger
        self.get_debug_output_dir = get_debug_output_dir
        self.sleep_func = sleep_func

    def _get_wait(self, driver, timeout=10):
        return self.browser_session.get_wait(driver, timeout)

    def collect_all_available_slots(self, driver, weeks_limit=8, only_weekday=None):
        """Collect available slots from the current lottery availability page."""
        slots = []
        try:
            def extract_from_html(html_text):
                found = set()
                for match in re.findall(
                    r"\d{1,2}月\d{1,2}日[^\n\r]{0,80}(?:\d{1,2}時[^\n\r]{0,40}分)?",
                    html_text,
                ):
                    found.add(match.strip())
                return found

            debug_dir = self.get_debug_output_dir()
            debug_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            try:
                self._get_wait(driver, 8).until(
                    lambda d: d.find_element(
                        By.CSS_SELECTOR, "#usedate-table tbody"
                    ).get_attribute("innerHTML").strip()
                    != ""
                )
            except Exception:
                try:
                    self._get_wait(driver, 8).until(
                        lambda d: d.find_element(By.ID, "usedate-loading")
                        .get_attribute("style")
                        .find("display: none")
                        != -1
                    )
                except Exception:
                    pass
            try:
                self._get_wait(driver, 6).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, "#usedate-table tbody")
                )
            except Exception:
                pass

            html = driver.page_source
            try:
                main_path = debug_dir / f"page_main_{ts}.html"
                with open(main_path, "w", encoding="utf-8") as fh:
                    fh.write(html)
                print("Saved debug HTML:", main_path)
            except Exception:
                self.logger.exception("Failed to save main page HTML")

            slots_set = set()
            try:
                def parse_current_table_and_add():
                    header_inputs = driver.find_elements(
                        By.CSS_SELECTOR,
                        '#usedate-table thead input[name="selectUseYMD"]',
                    )
                    dates = [header.get_attribute("value") for header in header_inputs]
                    rows = driver.find_elements(By.CSS_SELECTOR, "#usedate-table tbody tr")
                    for row in rows:
                        try:
                            time_label = row.find_element(By.TAG_NAME, "th").text.strip()
                        except Exception:
                            time_label = ""
                        tds = row.find_elements(By.TAG_NAME, "td")
                        for idx, td in enumerate(tds):
                            try:
                                ymd = dates[idx] if idx < len(dates) else ""
                                stime = td.find_element(
                                    By.CSS_SELECTOR, 'input[name="selectStime"]'
                                ).get_attribute("value")
                                etime = td.find_element(
                                    By.CSS_SELECTOR, 'input[name="selectEtime"]'
                                ).get_attribute("value")
                                field = td.find_element(
                                    By.CSS_SELECTOR, 'input[name="selectField"]'
                                ).get_attribute("value")
                                applied = ""
                                try:
                                    applied = td.find_element(
                                        By.CSS_SELECTOR, "span.font-weight-bold"
                                    ).text.strip()
                                except Exception:
                                    txt = td.text.strip()
                                    found = re.findall(r"\d+", txt)
                                    applied = found[-1] if found else ""

                                slot_str = (
                                    f"{ymd} {time_label} {stime}-{etime} "
                                    f"fields:{field} applied:{applied}"
                                )
                                if only_weekday is not None and ymd:
                                    try:
                                        parsed = datetime.datetime.strptime(ymd, "%Y%m%d")
                                        if parsed.weekday() != int(only_weekday):
                                            continue
                                    except Exception:
                                        pass
                                slots_set.add(slot_str)
                            except Exception:
                                continue
                    return dates

                try:
                    srch_start = driver.find_element(By.NAME, "srchStartYMD").get_attribute(
                        "value"
                    )
                    year = int(srch_start[0:4])
                    month = int(srch_start[4:6])
                    month_end_day = calendar.monthrange(year, month)[1]
                    target_month_end = f"{year}{month:02d}{month_end_day:02d}"
                except Exception:
                    target_month_end = None

                curr_dates = parse_current_table_and_add()

                if target_month_end:
                    iterations = 0
                    while True:
                        iterations += 1
                        try:
                            max_display = max(curr_dates) if curr_dates else None
                        except Exception:
                            max_display = None
                        if max_display and max_display >= target_month_end:
                            break
                        if iterations > weeks_limit:
                            break
                        try:
                            driver.find_element(By.ID, "next-week").click()
                        except Exception:
                            try:
                                anchors = driver.find_elements(By.XPATH, "//a[@onclick]")
                                clicked = False
                                for anchor in anchors:
                                    onclick = anchor.get_attribute("onclick") or ""
                                    if (
                                        "Next" in onclick
                                        or "next" in onclick
                                        or "doNextWeek" in onclick
                                        or "week" in onclick
                                    ):
                                        try:
                                            anchor.click()
                                            clicked = True
                                            break
                                        except Exception:
                                            continue
                                if not clicked:
                                    break
                            except Exception:
                                break
                        try:
                            self._get_wait(driver, 6).until(
                                lambda d: d.find_element(
                                    By.CSS_SELECTOR,
                                    '#usedate-table thead input[name="selectUseYMD"]',
                                ).get_attribute("value")
                                != (curr_dates[0] if curr_dates else "")
                            )
                        except Exception:
                            self.sleep_func(0.5)
                        try:
                            page_idx = iterations
                            page_path = debug_dir / f"page_{page_idx}_{ts}.html"
                            with open(page_path, "w", encoding="utf-8") as pf:
                                pf.write(driver.page_source)
                            print("Saved debug HTML:", page_path)
                        except Exception:
                            self.logger.exception("Failed to save paged HTML")
                        try:
                            curr_dates = parse_current_table_and_add()
                        except Exception:
                            break
            except Exception:
                self.logger.exception(
                    "DOM parsing of calendar failed, falling back to regex"
                )
                slots_set.update(extract_from_html(html))
                frames = driver.find_elements(By.TAG_NAME, "iframe")
                for frame in frames:
                    try:
                        driver.switch_to.frame(frame)
                        frame_html = driver.page_source
                        try:
                            frame_idx = frames.index(frame)
                            frame_path = debug_dir / f"iframe_{frame_idx}_{ts}.html"
                            with open(frame_path, "w", encoding="utf-8") as ff:
                                ff.write(frame_html)
                            print("Saved debug HTML:", frame_path)
                        except Exception:
                            self.logger.exception("Failed to save iframe HTML")
                        slots_set.update(extract_from_html(frame_html))
                        driver.switch_to.default_content()
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass

            if not slots_set:
                for i in range(weeks_limit):
                    self.sleep_func(0.5)
                    html = driver.page_source
                    try:
                        page_path = debug_dir / f"page_{i}_{ts}.html"
                        with open(page_path, "w", encoding="utf-8") as pf:
                            pf.write(html)
                        print("Saved debug HTML:", page_path)
                    except Exception:
                        self.logger.exception("Failed to save paged HTML")
                    slots_set.update(extract_from_html(html))
                    clicked = False
                    try:
                        anchors = driver.find_elements(By.XPATH, "//a[@onclick]")
                        for anchor in anchors:
                            onclick = anchor.get_attribute("onclick") or ""
                            if "next" in onclick or "Next" in onclick or "week" in onclick:
                                try:
                                    anchor.click()
                                    clicked = True
                                    break
                                except Exception:
                                    continue
                        if not clicked:
                            for txt in ("次へ", "次", ">", ">>"):
                                try:
                                    driver.find_element(By.LINK_TEXT, txt).click()
                                    clicked = True
                                    break
                                except Exception:
                                    continue
                    except Exception:
                        clicked = False
                    if not clicked:
                        break

            entries = []
            for slot in slots_set:
                match = re.match(
                    r"^(?P<ymd>\d{8}).*?(?P<stime>\d{1,4})-(?P<etime>\d{1,4})",
                    slot,
                )
                if match:
                    ymd = match.group("ymd")
                    try:
                        stime = int(match.group("stime"))
                    except Exception:
                        stime = 9999
                else:
                    ymd = ""
                    stime = 9999
                entries.append((ymd, stime, slot))
            entries.sort(key=lambda t: (t[0], t[1]))
            slots = [entry[2] for entry in entries]
        except Exception:
            self.logger.exception("空き日時収集中に例外が発生しました")

        out_path = (
            self.config["PATH"]["OUTPUT_CSV_PATH"]
            + "/available_slots_{0}.csv".format(datetime.date.today())
        )
        try:
            with open(out_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["slot"])
                for slot in slots:
                    writer.writerow([slot])
        except Exception:
            self.logger.exception("空き日時CSVの保存に失敗しました")

        print("Saved available slots to: " + out_path)
        if not slots:
            print("No slots found")
        return slots

    def check_court(self, month):
        """Run the legacy court availability flow."""
        driver = self.browser_session.create_driver()
        try:
            driver.get(self.config["URL"]["TOP_URL"])
            driver.switch_to.frame("pawae1002")
            self.navigation_service.go_to_vacant_search(driver)
            try:
                driver.find_element_by_name("monthGif" + month).click()
            except Exception:
                print("対象月が存在しません")
                self.browser_session.safe_quit(driver)
                return None
            driver.find_element_by_name("weektype5").click()
            self.navigation_service.select_weekly_vacant_conditions(driver)
            driver.find_element_by_name("gifName23").click()
            print(driver.page_source)
            return driver.page_source
        finally:
            pass
