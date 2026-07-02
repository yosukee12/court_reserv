# -*- coding: utf-8 -*-
from pathlib import Path
import time
import logging
import datetime
import calendar
import re
try:
    from .manage_id import Manage_Id as mi
    from .config import get_debug_output_dir, load_config
    from .browser import BrowserSession, LoginService, NavigationService
    from .services import LotteryService
except Exception:
    # allow running the module as a script (no package context)
    from manage_id import Manage_Id as mi
    from config import get_debug_output_dir, load_config
    from browser import BrowserSession, LoginService, NavigationService
    from services import LotteryService
import tkinter as tk
from tkinter import ttk, messagebox
from functools import partial

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException
from bs4 import BeautifulSoup as bs

# Config
config = load_config()

# Log
logfile = config['PATH']['LOG_PATH'] + '/court_reserv.log'
log_fmt = '%(asctime)s - %(levelname)s - %(message)s'
log_level_name = config.get('LOG', 'LEVEL', fallback='INFO').upper()
log_level = getattr(logging, log_level_name, logging.INFO)
logging.basicConfig(filename=logfile, format=log_fmt, level=log_level)

# Output csv path
check_lottery_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_lottery_{0}.csv'.format(datetime.date.today())
check_result_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_result_{0}.csv'.format(datetime.date.today())
determined_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/determined_result_{0}.csv'.format(datetime.date.today())
check_reserv_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_reserv_{0}.csv'.format(datetime.date.today())
alive_id_list_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/ID_list_alive_{0}.csv'.format(datetime.date.today())
dead_id_list_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/ID_list_dead_{0}.csv'.format(datetime.date.today())

top_url = config['URL']['TOP_URL']

class Court_Reserv(tk.Frame):
    def __init__(self, master=None):
        """
        コンストラクタ
          引数でIDリストのCSVファイルを指定してID dictに変換
          Tkinterのウィジェット作成
        """
        # tkinter
        super().__init__(master)
        self.pack()
        self.master.geometry("500x500")
        self.master.title("Court Reservation")
        self.browser_session = BrowserSession(config)
        self.login_service = LoginService(
            top_url=config["URL"]["TOP_URL"],
            wait_factory=self.browser_session.get_wait,
            logger=logging,
            show_info=messagebox.showinfo,
            ask_yes_no=messagebox.askyesno,
            sleep_func=time.sleep,
        )
        self.navigation_service = NavigationService(
            wait_factory=self.browser_session.get_wait,
            sleep_func=time.sleep,
        )
        self.lottery_service = LotteryService(
            config=config,
            browser_session=self.browser_session,
            login_service=self.login_service,
            navigation_service=self.navigation_service,
            logger=logging,
            show_info=messagebox.showinfo,
            output_id_dict=mi.output_csv_from_id_dict,
            sleep_func=time.sleep,
        )
        self.driver = None

        self.create_widgets()

    def create_widgets(self):
        # Mode selection (semi / full auto)
        self.mode_var = tk.StringVar(value='semi')
        self.label_mode = ttk.Label(self, text="動作モード:")
        self.radio_semi = ttk.Radiobutton(self, text='半自動', variable=self.mode_var, value='semi')
        self.radio_full = ttk.Radiobutton(self, text='全自動', variable=self.mode_var, value='full')

        self.label_mode.grid(row=0, column=2, padx=5)
        self.radio_semi.grid(row=0, column=3, padx=2)
        self.radio_full.grid(row=0, column=4, padx=2)
        # Label CSV PATH
        self.label_csvpath = ttk.Label(self, text="CSVファイル出力先: " + config['PATH']['OUTPUT_CSV_PATH'], background="white")

        #Entry
        self.entry_input_csv = ttk.Entry(self)
        self.entry_input_csv.insert(tk.END, config['PATH']['OUTPUT_CSV_PATH'] + "/")

        # Label 1
        self.label1 = ttk.Label(self, text="1. 毎月1日〜10日", background="white")

        #Reserv Button (mode-aware)
        self.button_semiauto_reserv = ttk.Button(self)
        self.button_semiauto_reserv.configure(text="抽選申込み")
        self.button_semiauto_reserv.configure(command = self.start_reservation_button)

        #Check Lottery Button
        self.button_check_lottery = ttk.Button(self)
        self.button_check_lottery.configure(text="抽選申込み状況確認")
        self.button_check_lottery.configure(command = self.check_lottery_button)

        # Label 2
        self.label2 = ttk.Label(self, text="2. 毎月14日〜", background="white")

        # Check Result Button
        self.button_check_result = ttk.Button(self)
        self.button_check_result.configure(text="抽選当選結果確認")
        self.button_check_result.configure(command=self.check_result_button)

        #Entry
        self.entry_result_csv = ttk.Entry(self)
        self.entry_result_csv.insert(tk.END, config['PATH']['OUTPUT_CSV_PATH'] + "/")

        # Determine Reserv Button
        self.button_determine_reserv = ttk.Button(self)
        self.button_determine_reserv.configure(text="予約確定")
        self.button_determine_reserv.configure(command=self.determine_button)

        # Check Reserv Button
        self.button_check_reserv = ttk.Button(self)
        self.button_check_reserv.configure(text="予約確定確認")
        self.button_check_reserv.configure(command=self.check_reserv_button)

        # Check Court Button
        self.button_check_court = ttk.Button(self)
        self.button_check_court.configure(text="空きコート確認")
        self.button_check_court.configure(command=partial(self.check_court, "9"))

        #Entry
        self.entry_check_id_csv = ttk.Entry(self)
        self.entry_check_id_csv.insert(tk.END, config['PATH']['OUTPUT_CSV_PATH'] + "/")

        # Check ID Button
        self.button_check_id = ttk.Button(self)
        self.button_check_id.configure(text="ID有効確認")
        self.button_check_id.configure(command=self.check_id_button)

        # 配置
        self.label_csvpath.grid(row=0, column=0, columnspan=2)
        self.entry_input_csv.grid(row=1, column=0, columnspan=10, padx=5, pady=5, sticky=tk.W+tk.E)
        self.label1.grid(row=2, column=0, columnspan=1)
        self.button_semiauto_reserv.grid(row=2, column=1, columnspan=3, sticky=tk.W + tk.E)
        self.button_check_lottery.grid(row=3, column=1, columnspan=1, padx=5, pady=5, sticky=tk.W + tk.E)
        self.label2.grid(row=4, column=0, columnspan=1)
        self.button_check_result.grid(row=4, column=1, columnspan=3, sticky=tk.W+tk.E)
        self.entry_result_csv.grid(row=5, column=0, columnspan=10, padx=5, pady=5, sticky=tk.W + tk.E)
        self.button_determine_reserv.grid(row=6, column=1, columnspan=1, padx=5, pady=5, sticky=tk.W+tk.E)
        self.button_check_reserv.grid(row=7, column=1, columnspan=1, padx=5, pady=5, sticky=tk.W + tk.E)
        self.button_check_court.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        self.entry_check_id_csv.grid(row=11, column=0, columnspan=10, padx=5, pady=5, sticky=tk.W + tk.E)
        self.button_check_id.grid(row=12, column=1, columnspan=1, padx=5, pady=5, sticky=tk.W + tk.E)

    # Helper methods for driver/login/logout/navigation
    def _start_driver(self):
        if getattr(self, 'driver', None) is None:
            self.driver = self.browser_session.create_driver()

    def _get_wait(self, timeout=10):
        return self.browser_session.get_wait(self.driver, timeout)

    def _login(self, user_id, password):
        self._start_driver()
        return self.login_service.login(self.driver, user_id, password)

    def _logout(self):
        try:
            self.navigation_service.logout(self.driver)
        except Exception:
            pass

    def _detect_captcha(self):
        return self.login_service.detect_captcha(self.driver)

    def _wait_for_captcha_solve(self, timeout_minutes=5):
        return self.login_service.wait_for_manual_captcha(
            self.driver, timeout_minutes=timeout_minutes
        )

    def _navigate_to_lottery_entry(self):
        # 抽選申し込み画面まで移動し、種目と公園を選択する共通処理
        self.navigation_service.go_to_lottery_entry(self.driver)
        self.navigation_service.select_lottery_tennis_park(self.driver)

    def collect_all_available_slots(self, weeks_limit=8, only_weekday=None):
        """
        抽選申込み画面の週毎ページを巡回して、見つかる全ての空き日時を収集してCSVに保存する。
        戻り値: list of slot strings
        """
        slots = []
        try:
            # helper to extract date+time combos from html text
            def extract_from_html(html_text):
                found = set()
                # look for patterns like "6月10日 ... 10時30分" or "6月10日"
                for m in re.findall(r"\d{1,2}月\d{1,2}日[^\n\r]{0,80}(?:\d{1,2}時[^\n\r]{0,40}分)?", html_text):
                    found.add(m.strip())
                return found

            # Search current document and any iframe documents
            searched = 0
            # prepare debug dir
            debug_dir = get_debug_output_dir()
            debug_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            # wait for vacancy table to be populated (AJAX)
            try:
                self._get_wait(8).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, '#usedate-table tbody').get_attribute('innerHTML').strip() != ''
                )
            except Exception:
                # fallback: wait until loading indicator is gone
                try:
                    self._get_wait(8).until(
                        lambda d: d.find_element(By.ID, 'usedate-loading').get_attribute('style').find('display: none') != -1
                    )
                except Exception:
                    pass
            # try current page — prefer parsing the vacancy calendar table
            try:
                self._get_wait(6).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, '#usedate-table tbody')
                )
            except Exception:
                pass

            html = self.driver.page_source
            # save main page html for inspection
            try:
                main_path = debug_dir / f'page_main_{ts}.html'
                with open(main_path, 'w', encoding='utf-8') as fh:
                    fh.write(html)
                print('Saved debug HTML:', main_path)
            except Exception:
                logging.exception('Failed to save main page HTML')

            # Parse the calendar table directly using Selenium DOM (more reliable than regex)
            slots_set = set()
            try:
                # gather header dates
                def parse_current_table_and_add():
                    header_inputs = self.driver.find_elements(By.CSS_SELECTOR, '#usedate-table thead input[name="selectUseYMD"]')
                    dates = [h.get_attribute('value') for h in header_inputs]
                    rows = self.driver.find_elements(By.CSS_SELECTOR, '#usedate-table tbody tr')
                    for row in rows:
                        try:
                            time_label = row.find_element(By.TAG_NAME, 'th').text.strip()
                        except Exception:
                            time_label = ''
                        tds = row.find_elements(By.TAG_NAME, 'td')
                        for idx, td in enumerate(tds):
                            try:
                                ymd = dates[idx] if idx < len(dates) else ''
                                koma = td.find_element(By.CSS_SELECTOR, 'input[name="selectKomaNo"]').get_attribute('value')
                                stime = td.find_element(By.CSS_SELECTOR, 'input[name="selectStime"]').get_attribute('value')
                                etime = td.find_element(By.CSS_SELECTOR, 'input[name="selectEtime"]').get_attribute('value')
                                field = td.find_element(By.CSS_SELECTOR, 'input[name="selectField"]').get_attribute('value')
                                # applied count shown in bold
                                applied = ''
                                try:
                                    applied = td.find_element(By.CSS_SELECTOR, 'span.font-weight-bold').text.strip()
                                except Exception:
                                    txt = td.text.strip()
                                    m = re.findall(r"\d+", txt)
                                    applied = m[-1] if m else ''

                                slot_str = f"{ymd} {time_label} {stime}-{etime} fields:{field} applied:{applied}"
                                # filter by weekday if requested (Monday=0 ... Sunday=6). Saturday==5
                                if only_weekday is not None and ymd:
                                    try:
                                        d = datetime.datetime.strptime(ymd, '%Y%m%d')
                                        if d.weekday() != int(only_weekday):
                                            continue
                                    except Exception:
                                        # if parse failed, skip filtering for this item
                                        pass
                                slots_set.add(slot_str)
                            except Exception:
                                continue
                    return dates

                # determine target month end from srchStartYMD hidden input
                try:
                    srch_start = self.driver.find_element(By.NAME, 'srchStartYMD').get_attribute('value')
                    year = int(srch_start[0:4])
                    month = int(srch_start[4:6])
                    month_end_day = calendar.monthrange(year, month)[1]
                    target_month_end = f"{year}{month:02d}{month_end_day:02d}"
                except Exception:
                    target_month_end = None

                # parse first page
                curr_dates = parse_current_table_and_add()

                # if we know the target month end, paginate until we cover it
                if target_month_end:
                    iterations = 0
                    while True:
                        iterations += 1
                        # get current max date displayed
                        try:
                            max_display = max(curr_dates) if curr_dates else None
                        except Exception:
                            max_display = None
                        if max_display and max_display >= target_month_end:
                            break
                        if iterations > weeks_limit:
                            break
                        # click next-week button
                        try:
                            btn = self.driver.find_element(By.ID, 'next-week')
                            btn.click()
                        except Exception:
                            # fallback: try clicking by onclick anchors
                            try:
                                anchors = self.driver.find_elements(By.XPATH, "//a[@onclick]")
                                clicked = False
                                for a in anchors:
                                    onclick = a.get_attribute('onclick') or ''
                                    if 'Next' in onclick or 'next' in onclick or 'doNextWeek' in onclick or 'week' in onclick:
                                        try:
                                            a.click()
                                            clicked = True
                                            break
                                        except Exception:
                                            continue
                                if not clicked:
                                    break
                            except Exception:
                                break
                        # wait for table update (header dates change)
                        try:
                            self._get_wait(6).until(lambda d: d.find_element(By.CSS_SELECTOR, '#usedate-table thead input[name="selectUseYMD"]').get_attribute('value') != (curr_dates[0] if curr_dates else ''))
                        except Exception:
                            time.sleep(0.5)
                        # save paged html
                        try:
                            page_idx = iterations
                            page_path = debug_dir / f'page_{page_idx}_{ts}.html'
                            with open(page_path, 'w', encoding='utf-8') as pf:
                                pf.write(self.driver.page_source)
                            print('Saved debug HTML:', page_path)
                        except Exception:
                            logging.exception('Failed to save paged HTML')
                        # parse new page and continue
                        try:
                            curr_dates = parse_current_table_and_add()
                        except Exception:
                            break
                else:
                    # unknown month end -> parse current page only
                    pass
            except Exception:
                # if DOM parsing failed, fallback to regex search across page and frames
                logging.exception('DOM parsing of calendar failed, falling back to regex')
                slots_set.update(extract_from_html(html))
                frames = self.driver.find_elements(By.TAG_NAME, 'iframe')
                for f in frames:
                    try:
                        self.driver.switch_to.frame(f)
                        fh = self.driver.page_source
                        # save iframe html
                        try:
                            frame_idx = frames.index(f)
                            frame_path = debug_dir / f'iframe_{frame_idx}_{ts}.html'
                            with open(frame_path, 'w', encoding='utf-8') as ff:
                                ff.write(fh)
                            print('Saved debug HTML:', frame_path)
                        except Exception:
                            logging.exception('Failed to save iframe HTML')
                        slots_set.update(extract_from_html(fh))
                        self.driver.switch_to.default_content()
                    except Exception:
                        try:
                            self.driver.switch_to.default_content()
                        except Exception:
                            pass
            # if nothing found yet, attempt to paginate weekly pages (best-effort)
            if not slots_set:
                for i in range(weeks_limit):
                    time.sleep(0.5)
                    html = self.driver.page_source
                    # save paginated page for debugging
                    try:
                        page_path = debug_dir / f'page_{i}_{ts}.html'
                        with open(page_path, 'w', encoding='utf-8') as pf:
                            pf.write(html)
                        print('Saved debug HTML:', page_path)
                    except Exception:
                        logging.exception('Failed to save paged HTML')
                    slots_set.update(extract_from_html(html))
                    # attempt several types of next controls
                    clicked = False
                    try:
                        # anchors with onclick
                        anchors = self.driver.find_elements(By.XPATH, "//a[@onclick]")
                        for a in anchors:
                            onclick = a.get_attribute('onclick') or ''
                            if 'next' in onclick or 'Next' in onclick or 'week' in onclick:
                                try:
                                    a.click()
                                    clicked = True
                                    break
                                except Exception:
                                    continue
                        if not clicked:
                            # fallback: links with caret or next text
                            for txt in ("次へ", "次", ">", ">>"):
                                try:
                                    el = self.driver.find_element(By.LINK_TEXT, txt)
                                    el.click()
                                    clicked = True
                                    break
                                except Exception:
                                    continue
                    except Exception:
                        clicked = False
                    if not clicked:
                        break

            # sort slots by ymd then numeric start time (avoid unicode string order issues)
            slots = []
            entries = []
            for s in slots_set:
                m = re.match(r'^(?P<ymd>\d{8}).*?(?P<stime>\d{1,4})-(?P<etime>\d{1,4})', s)
                if m:
                    ymd = m.group('ymd')
                    try:
                        stime = int(m.group('stime'))
                    except Exception:
                        stime = 9999
                else:
                    # fallback: put unknowns at end
                    ymd = ''
                    stime = 9999
                entries.append((ymd, stime, s))
            entries.sort(key=lambda t: (t[0], t[1]))
            slots = [t[2] for t in entries]

        except Exception:
            logging.exception("空き日時収集中に例外が発生しました")

        # save to CSV
        out_path = config['PATH']['OUTPUT_CSV_PATH'] + '/available_slots_{0}.csv'.format(datetime.date.today())
        try:
            import csv
            with open(out_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['slot'])
                for s in slots:
                    writer.writerow([s])
        except Exception:
            logging.exception("空き日時CSVの保存に失敗しました")

        print('Saved available slots to: ' + out_path)
        if not slots:
            print('No slots found')
        return slots

    def prompt_user_to_select_slots(self, slots, max_select=2):
        """
        ターミナル上でスロット一覧を表示してユーザーに選択させる。選択値のリストを返す。
        """
        if not slots:
            print('No slots found')
            return []
        print('Available slots:')
        for i, s in enumerate(slots, start=1):
            print(f"{i}: {s}")
        sel = input(f'Select up to {max_select} slots by number (comma separated): ').strip()
        if not sel:
            return []
        nums = []
        for part in sel.split(','):
            try:
                n = int(part.strip())
                if 1 <= n <= len(slots):
                    nums.append(n)
            except ValueError:
                continue
        nums = nums[:max_select]
        selected = [slots[n-1] for n in nums]
        print('Selected:')
        for s in selected:
            print(s)
        return selected

    # ここからボタン実行用メソッド
    def semiauto_reserv_button(self):
        """
        抽選申込みボタンが押された時の処理
        """
        self.semiauto_reserv(mi.get_id_dict_from_csv(self.entry_input_csv.get()))

    def start_reservation_button(self):
        """Mode-aware handler for reservation button: calls semi or full auto based on selection."""
        id_dict = mi.get_id_dict_from_csv(self.entry_input_csv.get())
        mode = self.mode_var.get() if getattr(self, 'mode_var', None) is not None else 'semi'
        if mode == 'full':
            # run full auto
            self.full_auto_reserv(id_dict)
        else:
            self.semiauto_reserv(id_dict)

    def check_lottery_button(self):
        """
        抽選申込み状況確認ボタンが押された時の処理
        """
        self.check_lottery(mi.get_id_dict_from_csv(self.entry_input_csv.get()), check_lottery_csv)

    def check_result_button(self):
        """
        抽選申込み結果確認ボタンが押された時の処理
        """
        self.check_result(mi.get_id_dict_from_csv(self.entry_input_csv.get()), check_result_csv)

    def determine_button(self):
        """
        予約確定ボタンが押された時の処理
        """
        self.determine_reserv(self.entry_result_csv.get(), determined_csv)

    def check_reserv_button(self):
        """
        抽選申込み結果確認ボタンが押された時の処理
        """
        self.check_reserv(mi.get_id_dict_from_csv(self.entry_result_csv.get()), check_reserv_csv)

    def check_id_button(self):
        """
        ID有効確認ボタンが押された時の処理
        """
        alive_id_list, dead_id_list = mi.get_alive_dead_id_dict(mi.get_id_dict_from_csv(self.entry_check_id_csv.get()))
        mi.output_csv_from_id_dict(alive_id_list, alive_id_list_csv)
        mi.output_csv_from_id_dict(dead_id_list, dead_id_list_csv)

    # ここからCourt Reservメソッド
    def semiauto_reserv(self, id_dict={}):
        """
        IDリストを引数にして
        半自動抽選申込み. 抽選申込み日の選択と申込みは手動
        """
        if not id_dict:
            id_dict = self.id_dict
        self.lottery_service.semiauto_reserv(id_dict)

    def full_auto_reserv(self, id_dict={}, max_attempts=2):
        """
        完全自動で抽選申込みを試みる。reCAPTCHA 等がある場合は送信に失敗する可能性あり。
        第2引数 `max_attempts` で各IDに対する最大申込み回数を指定。
        """
        if not id_dict:
            id_dict = self.id_dict
        self.lottery_service.full_auto_reserv(id_dict, max_attempts=max_attempts)

    def auto_select_and_submit_slots(self, selected_slots, submit=True, wait_alert_seconds=10):
        """
        selected_slots: list of strings in format 'YYYYMMDD <time_label> <stime>-<etime> ...'
        Attempts to select matching slots on the current calendar page and optionally submit the application.
        Returns dict {slot: True/False} indicating whether selection was applied for each slot.
        """
        return self.lottery_service.auto_select_and_submit_slots(
            self.driver,
            selected_slots,
            submit=submit,
            wait_alert_seconds=wait_alert_seconds,
        )

    def check_lottery(self, id_dict={}, output_csv_path=""):
        """
        IDリストを引数にして抽選申込み日を取得
        IDに申込み日を追加したdictを返す
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),申込日1,申込み日2]}
        第2引数に出力先CSRファイルパスを指定した場合はCSVを出力
        """
        if not id_dict:
            id_dict = self.id_dict
        return self.lottery_service.check_lottery(id_dict, output_csv_path)

    def check_result(self, id_dict={}, output_csv_path=""):
        """
        IDリストを引数にして抽選当選日を取得
        ※当選確定は手動
        IDに当選日を追加したdictを返す
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),当選日1,当選日2]}
        第2引数に出力先CSRファイルパスを指定した場合はCSVを出力
        """
        if not id_dict:
            id_dict = self.id_dict
        return self.lottery_service.check_result(id_dict, output_csv_path)

    def determine_reserv(self, input_csv_path="", output_csv_path=""):
        """
        抽選確定日が記入されたcsvを引数にして, 半手動抽選確定をする
        IDに確定日を追加したdictを返す
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),確定日1,確定日2]}
        第2引数に出力先CSRファイルパスを指定した場合はCSVを出力
        """
        print(input_csv_path)
        id_dict = mi.get_id_dict_from_csv(input_csv_path)

        result_dict = {}
        # Chromeドライバーの起動
        self.driver = self.browser_session.create_driver()
        for k, v in id_dict.items():
            self.driver.get(config['URL']['TOP_URL'])
            try:
                # ログインページへ移動
                self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserLoginAction);")
                self.driver.find_element(By.NAME,"userId").send_keys(k)
                self.driver.find_element(By.NAME,"password").send_keys(v[2])
                # ログイン
                time.sleep(0.5)
                self.driver.execute_script("javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);")
                # # 有効期限が近づいている画面が出た場合
                # if "お知らせ画面" in self.driver.title:
                #     if "利用者カードの有効期限が切れている" in self.driver.page_source:
                #         print("ID:" + k + " 期限切れ")
                #         continue
                #     else:
                #         self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
                # if "伝言表示画面" in self.driver.title:
                #     self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageNextAction);")
                #     logging.warn("ID:" + k + " 伝言アリ")

            except UnexpectedAlertPresentException:
                print("ID:" + k + " 期限切れ")
                logging.warning("ID:" + k + " 期限切れ")
                continue

            if "ホーム画面" in self.driver.title:
                try:
                    # 抽選結果確認画面へ
                    self.navigation_service.go_to_lottery_result_list(self.driver)
                    # Beautiful soupで申込み日と時間の取得
                    time.sleep(0.5)
                    soup = bs(self.driver.page_source, 'html.parser')
                    found_day_list = [elem.text for elem in soup.find_all('span', string=re.compile("月.*日(.*)"))]
                    found_time_list = [elem.text for elem in soup.find_all(string=re.compile("時.*分～.*時.*分"))]
                    # 当選日1日パターン
                    if len(found_day_list) == 1:
                        self._get_wait(240).until(EC.alert_is_present(),
                                                              'Timed out waiting for PA creation ' +
                                                              'confirmation popup to appear.')
                        alert = self.driver.switch_to.alert
                        alert.accept()
                        print("ID:" + k + " 確定日→ " + found_day_list[0] + " " + found_time_list[0])
                        result_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0]]
                        logging.info("ID:" + k + " 予約確定完了→ " + found_day_list[0] + " " + found_time_list[0])
                    # 当選日2日パターン
                    elif len(found_day_list) == 2:
                        for i in range(2):
                            # 2日当選日があった場合、labelが空になるまで
                            self._get_wait(240).until(EC.alert_is_present(),
                                                                  'Timed out waiting for PA creation ' +
                                                                  'confirmation popup to appear.')
                            alert = self.driver.switch_to.alert
                            alert.accept()
                            if i == 0:
                                print("ID:" + k + " 確定日→ " + found_day_list[0] + " " + found_time_list[0])
                                logging.info("ID:" + k + " 予約確定完了→ " + found_day_list[0] + " " + found_time_list[0])
                            elif i == 1:
                                print("ID:" + k + " 確定日→ " + found_day_list[1] + " " + found_time_list[1])
                                result_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0],found_day_list[1] + " " + found_time_list[1]]
                                logging.info("ID:" + k + " 予約確定完了→ " + found_day_list[1] + " " + found_time_list[1])

                            ## 確定後の画面のhtmlを保存
                            #html = self.driver.page_source
                            #with open(config['PATH']['OUTPUT_CSV_PATH'] + '/' + k + '_' + found_list[2] + found_list[
                            #    3] + '.html', 'w', encoding='utf-8') as f:
                            #    f.write(html)

                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    result_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue

            time.sleep(1)
            # ログアウト
            self.navigation_service.logout(self.driver)
            time.sleep(1)
        self.browser_session.safe_close(self.driver)
        self.driver = None

        if output_csv_path != "":
            mi.output_csv_from_id_dict(result_dict, output_csv_path)

        return result_dict

    def check_reserv(self, id_dict={}, output_csv_path=""):
        """
        IDリストを引数にして予約確定日を取得
        IDに確定日を追加したdictを返す
            ようにしたいが今はsleepで止めて手動確認する方式
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),確定日1,確定日2]}
        第2引数に出力先CSRファイルパスを指定した場合はCSVを出力
        """
        if not id_dict:
            id_dict = self.id_dict

        result_dict = {}
        # Chromeドライバーの起動
        self.driver = self.browser_session.create_driver()
        for k, v in id_dict.items():
            self.driver.get(config['URL']['TOP_URL'])
            try:
                # ログインページへ移動
                self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserLoginAction);")
                self.driver.find_element(By.NAME,"userId").send_keys(k)
                self.driver.find_element(By.NAME,"password").send_keys(v[2])
                # ログイン
                time.sleep(0.5)
                self.driver.execute_script("javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);")
                # # 有効期限が近づいている画面が出た場合
                # if "お知らせ画面" in self.driver.title:
                #     if "利用者カードの有効期限が切れている" in self.driver.page_source:
                #         print("ID:" + k + " 期限切れ")
                #         continue
                #     else:
                #         self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
                # if "伝言表示画面" in self.driver.title:
                #     self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageNextAction);")
                #     logging.warn("ID:" + k + " 伝言アリ")

            except UnexpectedAlertPresentException:
                print("ID:" + k + " 期限切れ")
                logging.warning("ID:" + k + " 期限切れ")
                continue

            if "ホーム画面" in self.driver.title:
                try:
                    # 予約確認画面へ
                    self.navigation_service.go_to_reservation_list(self.driver)
                    # TODO: 当選確定済の当選結果 のみ出力させたい
                    time.sleep(3)
                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    result_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue
            # ログアウト
            self.navigation_service.logout(self.driver)
            time.sleep(1)

        self.browser_session.safe_close(self.driver)
        self.driver = None
        # if output_csv_path != "":
        #     mi.output_csv_from_id_dict(result_dict, output_csv_path)

        return result_dict

    def check_court(self, month):
        """
        コートの空き状況をチェック
        """
        # Chromeドライバーの起動
        self.driver = self.browser_session.create_driver()
        self.driver.get(top_url)
        # フレーム移動
        self.driver.switch_to.frame("pawae1002")
        # 空き状況ページへ移動
        self.navigation_service.go_to_vacant_search(self.driver)
        try:
            self.driver.find_element_by_name("monthGif" + month).click() # 月選択
        except:
            print("対象月が存在しません")
            self.browser_session.safe_quit(self.driver)
            self.driver = None
            exit()
        # 曜日選択 土曜固定
        self.driver.find_element_by_name("weektype5").click()
        self.navigation_service.select_weekly_vacant_conditions(self.driver)
        # 場所選択 府中の森固定
        self.driver.find_element_by_name("gifName23").click()
        print(self.driver.page_source)
        # TODO ページの保存
        
    
def main():
    """
    Compatibility GUI entrypoint for `python court_reserv/court_reserv.py`.
    """
    root = tk.Tk()
    cr = Court_Reserv(master=root)
    cr.mainloop()
    
if __name__ == '__main__':
    main()
