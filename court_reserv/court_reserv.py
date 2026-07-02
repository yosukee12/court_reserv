# -*- coding: utf-8 -*-
import configparser
from pathlib import Path
import time
import logging
import datetime
import calendar
import re
try:
    from .manage_id import Manage_Id as mi
except Exception:
    # allow running the module as a script (no package context)
    from manage_id import Manage_Id as mi
import tkinter as tk
from tkinter import ttk, messagebox
from functools import partial

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, JavascriptException
from bs4 import BeautifulSoup as bs

# Config
config = configparser.ConfigParser()
# Load config.ini from package directory (so scripts can run from workspace root)
_base_dir = Path(__file__).resolve().parent
_config_path = _base_dir / 'config.ini'
if not _config_path.exists():
    raise FileNotFoundError(f'config.ini not found at {_config_path}. Run from package or create config.ini there.')
config.read(_config_path, encoding='utf-8')

# Log
logfile = config['PATH']['LOG_PATH'] + '/court_reserv.log'
log_fmt = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(filename=logfile, format=log_fmt, level=logging.INFO)

# Output csv path
check_lottery_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_lottery_{0}.csv'.format(datetime.date.today())
check_result_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_result_{0}.csv'.format(datetime.date.today())
determined_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/determined_result_{0}.csv'.format(datetime.date.today())
check_reserv_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/check_reserv_{0}.csv'.format(datetime.date.today())
alive_id_list_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/ID_list_alive_{0}.csv'.format(datetime.date.today())
dead_id_list_csv = config['PATH']['OUTPUT_CSV_PATH'] + '/ID_list_dead_{0}.csv'.format(datetime.date.today())

# Selenium Options
options = Options()
options.add_argument('--disable-gpu');
options.add_argument('--disable-extensions');
options.add_argument('--proxy-server="direct://"');
options.add_argument('--proxy-bypass-list=*');
options.add_argument('--start-maximized');

driver_path = config['PATH']['DRIVER_PATH']
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
            self.driver = webdriver.Chrome(service=Service(driver_path), options=options)

    def _login(self, user_id, password):
        """ログイン処理。成功するとホーム画面になるまで待つ。"""
        self._start_driver()
        self.driver.get(config['URL']['TOP_URL'])
        try:
            try:
                self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserLoginAction);")
            except JavascriptException:
                # page may not expose doAction; continue and try locating form directly
                pass

            # try to find login fields; if not on main document, try common iframe
            try:
                user_el = self.driver.find_element(By.NAME, "userId")
            except Exception:
                try:
                    self.driver.switch_to.frame("pawae1002")
                    user_el = self.driver.find_element(By.NAME, "userId")
                except Exception:
                    logging.warning("Login form not found for user %s", user_id)
                    try:
                        self.driver.switch_to.default_content()
                    except Exception:
                        pass
                    return False

            user_el.send_keys(user_id)
            self.driver.find_element(By.NAME, "password").send_keys(password)
            time.sleep(0.5)
            try:
                self.driver.execute_script("javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);")
            except JavascriptException:
                # fallback: try to submit by sending Enter
                try:
                    self.driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
                except Exception:
                    pass

            # login may produce an alert on failure; detect and accept it
            try:
                WebDriverWait(self.driver, 2).until(EC.alert_is_present())
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                logging.warning("ID:%s login alert: %s", user_id, alert_text)
                return False
            except TimeoutException:
                # no alert -> proceed
                # If a captcha is present, prompt the user to solve it instead of failing
                try:
                    if self._detect_captcha():
                        ok = self._wait_for_captcha_solve()
                        return bool(ok)
                except Exception:
                    pass
                return True
        except UnexpectedAlertPresentException:
            # unexpected alert caught during commands
            try:
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                logging.warning("ID:%s unexpected alert during login: %s", user_id, alert_text)
            except Exception:
                pass
            return False

    def _logout(self):
        try:
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
        except Exception:
            pass

    def _detect_captcha(self):
        """Detect whether a reCAPTCHA-like widget appears on the current page."""
        if not getattr(self, 'driver', None):
            return False
        try:
            # look for iframe or div indicators
            els = self.driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha'], div.g-recaptcha")
            if els and len(els) > 0:
                return True
            ps = self.driver.page_source
            if 'g-recaptcha' in ps or 'recaptcha' in ps:
                return True
        except Exception:
            return False
        return False

    def _wait_for_captcha_solve(self, timeout_minutes=5):
        """Prompt user to solve captcha manually and wait until it's gone (or cancelled).
        Returns True if captcha cleared, False if cancelled/timed-out.
        """
        try:
            messagebox.showinfo('reCAPTCHA 検出', 'ページ上にreCAPTCHAが検出されました。ブラウザで手動で認証してください。完了したら「OK」を押してください。')
        except Exception:
            print('reCAPTCHA detected: please solve it manually in the browser and then continue.')
        start = time.time()
        timeout = timeout_minutes * 60
        while True:
            try:
                if not self._detect_captcha():
                    return True
            except Exception:
                return False
            if time.time() - start > timeout:
                try:
                    cont = messagebox.askyesno('reCAPTCHA まだ有効', 'reCAPTCHAがまだ残っています。続行して再試行しますか？(OK=再確認 / キャンセル=中断)')
                except Exception:
                    cont = False
                if not cont:
                    return False
                start = time.time()
            time.sleep(1)

    def _navigate_to_lottery_entry(self):
        # 抽選申し込み画面まで移動し、種目と公園を選択する共通処理
        self.driver.execute_script("javascript:doAction(document.form1, gLotWOpeLotSearchAction);")
        # 種目選択（テニス（人工芝））
        self.driver.execute_script("javascript:doLotEntry('130');")
        time.sleep(1)
        Select(self.driver.find_element(By.ID, "bname")).select_by_value("1301270")
        self.driver.execute_script("changeBname(document.form1);")
        wait = WebDriverWait(self.driver, 10)
        wait.until(lambda d: any(
            opt.get_attribute("value") == "12700020"
            for opt in Select(d.find_element(By.ID, "iname")).options
        ))
        Select(self.driver.find_element(By.ID, "iname")).select_by_value("12700020")

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
            debug_dir = Path(config['PATH']['OUTPUT_CSV_PATH']) / 'debug_pages'
            debug_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            # wait for vacancy table to be populated (AJAX)
            try:
                WebDriverWait(self.driver, 8).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, '#usedate-table tbody').get_attribute('innerHTML').strip() != ''
                )
            except Exception:
                # fallback: wait until loading indicator is gone
                try:
                    WebDriverWait(self.driver, 8).until(
                        lambda d: d.find_element(By.ID, 'usedate-loading').get_attribute('style').find('display: none') != -1
                    )
                except Exception:
                    pass
            # try current page — prefer parsing the vacancy calendar table
            try:
                WebDriverWait(self.driver, 6).until(
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
                            WebDriverWait(self.driver, 6).until(lambda d: d.find_element(By.CSS_SELECTOR, '#usedate-table thead input[name="selectUseYMD"]').get_attribute('value') != (curr_dates[0] if curr_dates else ''))
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
        # 引数でID dictを指定しない場合
        if not id_dict:
            id_dict = self.id_dict
        # 申し込み人数カウント用
        list_count = 1
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
        for k, v in id_dict.items():
            reserv_count = 0
            self.driver.get(config['URL']['TOP_URL'])
            print("申し込み " + str(list_count) + "人目/" + str(len(id_dict)) + "人" + v[0])
            try:
                # ログインページへ移動
                self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserLoginAction);")
                self.driver.find_element(By.NAME,"userId").send_keys(k)
                self.driver.find_element(By.NAME,"password").send_keys(v[2])
                # ログイン
                time.sleep(0.5)
                self.driver.execute_script("javascript:submitLogin(document.form1,gRsvWUserAttestationLoginAction, event);")
            except UnexpectedAlertPresentException:
                print("ID:" + k + " 期限切れ")
                logging.warning("ID:" + k + " 期限切れ")
                continue

            # # 有効期限が近づいている画面が出た場合
            # if "お知らせ画面" in self.driver.title:
            #     self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
            #     logging.warn("ID:" + k + " 期限が近くなっています")

            # if "伝言表示画面" in self.driver.title:
            #     self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageNextAction);")
            #     logging.warn("ID:" + k + " 伝言アリ")
            logging.info("ID:" + k + " ログイン")

            if "ホーム画面" in self.driver.title:
                # 抽選申し込み画面へ
                self.driver.execute_script("javascript:doAction(document.form1, gLotWOpeLotSearchAction);")
                # 種目選択（テニス（人工芝））
                self.driver.execute_script("javascript:doLotEntry('130');")
                time.sleep(1)

                # 公園選択（府中の森公園）
                Select(self.driver.find_element(By.ID,"bname")).select_by_value("1301270")
                # time.sleep(1)
                self.driver.execute_script("changeBname(document.form1);")

                wait = WebDriverWait(self.driver, 10)
                wait.until(lambda d: any(
                    opt.get_attribute("value") == "12700020"
                    for opt in Select(d.find_element(By.ID, "iname")).options
                    ))
                # 種目選択2回目（テニス（人工芝））
                Select(self.driver.find_element(By.ID,"iname")).select_by_value("12700020")

                # # 公園選択（大井ふ頭Bオムニ）
                # Select(self.driver.find_element(By.ID,"bname")).select_by_value("1301315")
                # time.sleep(1)
                # # 種目選択2回目（テニス（人工芝））
                # Select(self.driver.find_element(By.ID,"iname")).select_by_value("13150110")


                # # 種目選択（テニス（ハード））
                # self.driver.execute_script("javascript:doLotEntry('120');")

                # # 公園選択（大井ふ頭Bハード）
                # Select(self.driver.find_element(By.ID,"bname")).select_by_value("1201315")
                # time.sleep(1)
                # # 種目選択2回目（テニス（ハード））
                # Select(self.driver.find_element(By.ID,"iname")).select_by_value("13150050")

                while reserv_count < 2:
                    # 申し込み中処理（手動申し込み）
                    time.sleep(0.5)
                    try:
                        if "東京都スポーツ施設サービス" in self.driver.title:
                            logging.info("ID:" + k + " ログアウト")
                            break
                        elif "申込内容確認画面" in self.driver.title:
                            reserv_count += 1
                            soup = bs(self.driver.page_source, 'html.parser')
                            # Beautiful soupで申込み日と時間の取得
                            foundlist = [elem.string for elem in soup.find_all('td', string=['年', '月', '日', '時', '分'])]
                            if reserv_count == 1:
                                # 申し込み番号入力（1件目）
                                time.sleep(0.3)
                                Select(self.driver.find_element(By.ID,"apply")).select_by_value("1-1")
                                time.sleep(0.2)
                            elif reserv_count == 2:
                                time.sleep(0.3)
                                Select(self.driver.find_element(By.ID,"apply")).select_by_value("2-1")
                                time.sleep(0.2)
                            # 申込み実行 → reCAPTCHA が出る可能性があるため手動クリックを促し、
                            # 出た場合はブラウザでの認証を待機する
                            # Show a GUI prompt only if captcha is present; otherwise print a terminal hint
                            try:
                                if self._detect_captcha():
                                    messagebox.showinfo('手動操作が必要です', '申込み実行画面が表示されました。画面上で申込み実行を手動で実施してください。reCAPTCHAが表示された場合はブラウザ上で認証してください。')
                                else:
                                    print('Manual submit required: please click submit in the browser. If reCAPTCHA appears, solve it manually.')
                            except Exception:
                                print('Manual submit required: please click submit in the browser. If reCAPTCHA appears, solve it manually.')

                            # wait for completion; if captcha appears at any point, prompt user to solve
                            while not "抽選メール送信完了画面" in self.driver.title:
                                try:
                                    # if captcha detected, ask user to solve it before continuing
                                    try:
                                        if self._detect_captcha():
                                            solved = self._wait_for_captcha_solve()
                                            if not solved:
                                                logging.warning('User cancelled captcha solving; aborting this reservation')
                                                break
                                    except Exception:
                                        pass

                                    # ポップアップ処理
                                    WebDriverWait(self.driver, 60).until(EC.alert_is_present(),
                                                            'Timed out waiting for PA creation ' +
                                                            'confirmation popup to appear.')
                                    alert = self.driver.switch_to.alert
                                    alert.accept()
                                    WebDriverWait(self.driver, 1).until(EC.alert_is_present(),
                                                            'Timed out waiting for PA creation ' +
                                                            'confirmation popup to appear.')
                                    time.sleep(0.3)
                                except (TimeoutException, UnexpectedAlertPresentException):
                                    continue
                            print("reserved: ID = " + k + ", reserv_count = " + str(reserv_count))
                    except TimeoutException or UnexpectedAlertPresentException:
                        continue
            list_count += 1
            time.sleep(0.5)
            # ログアウト
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
            time.sleep(0.5)
        self.driver.close()

    def full_auto_reserv(self, id_dict={}, max_attempts=2):
        """
        完全自動で抽選申込みを試みる。reCAPTCHA 等がある場合は送信に失敗する可能性あり。
        第2引数 `max_attempts` で各IDに対する最大申込み回数を指定。
        """
        if not id_dict:
            id_dict = self.id_dict
        list_count = 1
        self._start_driver()
        for k, v in id_dict.items():
            reserv_count = 0
            print("自動申し込み " + str(list_count) + "人目/" + str(len(id_dict)) + "人" + v[0])
            if not self._login(k, v[2]):
                continue
            logging.info("ID:%s ログイン", k)

            if "ホーム画面" in self.driver.title:
                try:
                    self._navigate_to_lottery_entry()
                    # 申込内容確認画面が表示されるのを待ち、見つかれば自動で送信する
                    while reserv_count < max_attempts:
                        time.sleep(0.5)
                        if "申込内容確認画面" in self.driver.title:
                            reserv_count += 1
                            # 申し込み番号入力
                            sel_val = f"{reserv_count}-1"
                            Select(self.driver.find_element(By.ID, "apply")).select_by_value(sel_val)
                            time.sleep(0.2)
                            # 申込み実行（可能なら自動でクリック）
                            try:
                                try:
                                    if self._detect_captcha():
                                        ok = self._wait_for_captcha_solve()
                                        if not ok:
                                            logging.warning("ID:%s captcha not solved, aborting auto apply", k)
                                            break
                                except Exception:
                                    pass
                                self.driver.execute_script("javascript:sendLotApply(document.form1, gLotWInstLotApplyAction, event);")
                            except Exception:
                                logging.warning("ID:%s 自動申込みスクリプト実行に失敗", k)
                            # ポップアップ処理
                            try:
                                WebDriverWait(self.driver, 60).until(EC.alert_is_present())
                                alert = self.driver.switch_to.alert
                                alert.accept()
                            except TimeoutException:
                                logging.warning("ID:%s 申込み確認ポップアップが表示されませんでした", k)
                            # 完了画面になるのを待つ（短時間）
                            try:
                                WebDriverWait(self.driver, 10).until(lambda d: "抽選メール送信完了画面" in d.title)
                            except Exception:
                                logging.info("ID:%s 抽選送信完了画面への遷移を確認できませんでした", k)
                            print("reserved: ID = " + k + ", reserv_count = " + str(reserv_count))
                        elif "東京都スポーツ施設サービス" in self.driver.title:
                            logging.info("ID:%s ログアウト検出", k)
                            break
                        else:
                            time.sleep(0.5)
                except Exception as e:
                    logging.exception("ID:%s 自動申込み中に例外", k)

            list_count += 1
            time.sleep(0.5)
            self._logout()
            time.sleep(0.5)
        try:
            self.driver.close()
        except Exception:
            pass

    def auto_select_and_submit_slots(self, selected_slots, submit=True, wait_alert_seconds=10):
        """
        selected_slots: list of strings in format 'YYYYMMDD <time_label> <stime>-<etime> ...'
        Attempts to select matching slots on the current calendar page and optionally submit the application.
        Returns dict {slot: True/False} indicating whether selection was applied for each slot.
        """
        result = {}
        if not getattr(self, 'driver', None):
            raise RuntimeError('Driver not started')

        for s in selected_slots:
            # parse ymd and stime
            parts = s.split()
            ymd = parts[0] if parts else ''
            stime = ''
            # find pattern like 900-1100 after time label
            for p in parts:
                if '-' in p and p.split('-')[0].isdigit():
                    stime = p.split('-')[0]
                    break
            # ensure the week containing ymd is displayed (try paginating)
            if ymd:
                try:
                    attempts = 0
                    print(f"[debug] Ensure week for {ymd} is displayed")
                    while attempts < 10:
                        attempts += 1
                        headers = self.driver.execute_script("return Array.from(document.querySelectorAll('#usedate-table thead input[name=\"selectUseYMD\"]')).map(h=>h.value);")
                        print(f"[debug] attempt {attempts}, headers={headers}")
                        if ymd in headers:
                            print(f"[debug] target {ymd} found in headers")
                            break
                        # decide direction: if ymd > max -> click next, if ymd < min -> click prev
                        if headers:
                            try:
                                min_h = min(headers)
                                max_h = max(headers)
                                print(f"[debug] min_h={min_h}, max_h={max_h}")
                                if ymd > max_h:
                                    # next-week
                                    print('[debug] clicking next-week')
                                    try:
                                        self.driver.execute_script("document.getElementById('next-week').click();")
                                    except Exception:
                                        try:
                                            self.driver.execute_script("doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);")
                                        except Exception:
                                            print('[debug] next-week click failed')
                                elif ymd < min_h:
                                    print('[debug] clicking last-week')
                                    try:
                                        self.driver.execute_script("document.getElementById('last-week').click();")
                                    except Exception:
                                        try:
                                            self.driver.execute_script("doPrevWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);")
                                        except Exception:
                                            print('[debug] last-week click failed')
                                else:
                                    # not in current range, attempt next
                                    print('[debug] not in range, clicking next-week')
                                    try:
                                        self.driver.execute_script("document.getElementById('next-week').click();")
                                    except Exception:
                                        try:
                                            self.driver.execute_script("doNextWeek(document.form1, gLotWTransLotInstSrchVacantAjaxAction);")
                                        except Exception:
                                            print('[debug] fallback next-week click failed')
                            except Exception as e:
                                print(f'[debug] header compare failed: {e}')
                                try:
                                    self.driver.execute_script("document.getElementById('next-week').click();")
                                except Exception:
                                    pass
                        time.sleep(0.6)
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
                                    // try to click the cell or inner number element
                                    try{ td.click(); info.clicked = true; }catch(e){}
                                    try{
                                        var num = td.querySelector('span.font-weight-bold');
                                        if(num){ num.click(); info.clicked = true; }
                                    }catch(e){}
                                    // also set checkboxes so state is consistent
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
                # return JSON string for reliable serialization
                # ensure the JS IIFE does not leave a trailing semicolon inside the
                # JSON.stringify wrapper which causes `Unexpected token ';'`.
                clean_js = js.rstrip()
                if clean_js.endswith(';'):
                    clean_js = clean_js[:-1]
                json_js = 'return JSON.stringify(' + clean_js + ');'
                info_str = self.driver.execute_script(json_js, ymd, stime)
                import json as _json
                try:
                    info = _json.loads(info_str) if info_str else None
                except Exception:
                    info = None
                print(f"[debug] select attempt for {s}: {info}")
                ok = bool(info and info.get('found', False))
            except Exception as e:
                print(f"[debug] execute_script failed: {e}")
                ok = False
            result[s] = ok

        # submit if requested and at least one selection succeeded
        if submit and any(result.values()):
            try:
                # open the confirmation/apply flow
                try:
                    self.driver.execute_script("javascript:doApplay(document.form1, gLotWInstTempLotApplyAction);")
                except Exception:
                    try:
                        btn = self.driver.find_element(By.ID, 'btn-go')
                        btn.click()
                    except Exception:
                        pass

                # Count how many slots were successfully selected
                success_count = sum(1 for v in result.values() if v)

                # For each successful slot, wait for the confirmation page, set the apply select,
                # then execute the final apply action (matching semiauto_reserv behavior)
                applied = 0
                for i in range(success_count):
                    try:
                        WebDriverWait(self.driver, wait_alert_seconds).until(lambda d: '申込内容確認画面' in d.title)
                    except Exception:
                        # if confirmation page doesn't appear, try short sleep and continue
                        time.sleep(0.5)
                    try:
                        applied += 1
                        sel_val = f"{applied}-1"
                        try:
                            Select(self.driver.find_element(By.ID, 'apply')).select_by_value(sel_val)
                        except Exception:
                            logging.warning('apply select not found to set %s', sel_val)
                        time.sleep(0.2)
                        try:
                            # If captcha present, prompt user to solve before submitting
                            try:
                                if self._detect_captcha():
                                    ok = self._wait_for_captcha_solve()
                                    if not ok:
                                        logging.warning('User cancelled captcha handling; aborting apply loop')
                                        break
                            except Exception:
                                pass
                            self.driver.execute_script("javascript:sendLotApply(document.form1, gLotWInstLotApplyAction, event);")
                        except Exception:
                            # fallback: try clicking a known apply button if present
                            try:
                                btn_apply = self.driver.find_element(By.ID, 'btn-apply')
                                btn_apply.click()
                            except Exception:
                                pass

                        # handle confirmation alert
                        try:
                            WebDriverWait(self.driver, wait_alert_seconds).until(EC.alert_is_present())
                            alert = self.driver.switch_to.alert
                            alert_text = alert.text
                            alert.accept()
                            logging.info('Accepted confirmation alert: %s', alert_text)
                        except Exception:
                            logging.info('No confirmation alert appeared for apply %s', sel_val)

                        # wait for completion page briefly
                        try:
                            WebDriverWait(self.driver, 10).until(lambda d: '抽選メール送信完了画面' in d.title)
                        except Exception:
                            # not fatal; continue to next apply
                            pass
                    except Exception:
                        logging.exception('Error during confirmation/apply loop')
                        break
            except Exception:
                logging.exception('Error during submit')

        return result

    def check_lottery(self, id_dict={}, output_csv_path=""):
        """
        IDリストを引数にして抽選申込み日を取得
        IDに申込み日を追加したdictを返す
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),申込日1,申込み日2]}
        第2引数に出力先CSRファイルパスを指定した場合はCSVを出力
        """
        # 引数でID dictを指定しない場合
        if not id_dict:
            id_dict = self.id_dict
        reserv_dict = {}
        
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
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
                    # 抽選申し込み確認画面へ
                    self.driver.execute_script("javascript:doAction(document.form1, gLotWTransLotCancelListAction);")
                    # Beautiful soupで申込み日と時間の取得
                    time.sleep(0.5)
                    soup = bs(self.driver.page_source, 'html.parser')
                    found_day_list = [elem.text for elem in soup.find_all(string=re.compile("月.*日(.*)"))]
                    found_time_list = [elem.text for elem in soup.find_all(string=re.compile("時.*分"))]
                    if len(found_day_list) == 2:
                        print("ID:" + k + " 申込み日1→ " + found_day_list[0] + " " + found_time_list[0] + found_time_list[1])
                        print("ID:" + k + " 申込み日2→ " + found_day_list[1] + " " + found_time_list[2]+ found_time_list[3])
                        reserv_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0] + found_time_list[1], found_day_list[1] + " " + found_time_list[2] + found_time_list[3]]
                    elif len(found_day_list) == 1:
                        print("ID:" + k + " 申込み日1→ " + found_day_list[0] + " " + found_time_list[0] + found_time_list[1])
                        reserv_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0] + found_time_list[1]]
                    else:
                        print("ID:" + k + " 申込みなし")
                        reserv_dict[k] = [v[0], v[1], v[2], "", ""]
                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    reserv_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue

            time.sleep(1)
            # ログアウト
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
            time.sleep(1)
        self.driver.close()

        if output_csv_path != "":
            mi.output_csv_from_id_dict(reserv_dict, output_csv_path)
        return reserv_dict

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

        result_dict = {}
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
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
                # 有効期限が近づいている画面が出た場合
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
                    self.driver.execute_script("javascript:doAction(document.form1, gLotWTransLotElectListAction);")
                    # Beautiful soupで申込み日と時間の取得
                    time.sleep(0.5)
                    soup = bs(self.driver.page_source, 'html.parser')
                    found_day_list = [elem.text for elem in soup.find_all('span', string=re.compile("月.*日(.*)"))]
                    found_time_list = [elem.text for elem in soup.find_all(string=re.compile("時.*分～.*時.*分"))]
                    # 当選日1日パターン
                    if len(found_day_list) == 1:
                        print("ID:" + k + " 当選日1→ " + found_day_list[0] + " " + found_time_list[0])
                        result_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0]]
                    # 当選日2日パターン
                    elif len(found_day_list) == 2:
                        print("ID:" + k + " 当選日1→ " + found_day_list[0] + " " + found_time_list[0])
                        print("ID:" + k + " 当選日2→ " + found_day_list[1] + " " + found_time_list[1])
                        result_dict[k] = [v[0], v[1], v[2], found_day_list[0] + " " + found_time_list[0], found_day_list[1] + " " + found_time_list[1]]

                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    #result_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue
            time.sleep(1)
            # ログアウト
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
            time.sleep(1)
        self.driver.close()

        if output_csv_path != "":
            mi.output_csv_from_id_dict(result_dict, output_csv_path)

        return result_dict

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
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
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
                    self.driver.execute_script("javascript:doAction(document.form1, gLotWTransLotElectListAction);")
                    # Beautiful soupで申込み日と時間の取得
                    time.sleep(0.5)
                    soup = bs(self.driver.page_source, 'html.parser')
                    found_day_list = [elem.text for elem in soup.find_all('span', string=re.compile("月.*日(.*)"))]
                    found_time_list = [elem.text for elem in soup.find_all(string=re.compile("時.*分～.*時.*分"))]
                    # 当選日1日パターン
                    if len(found_day_list) == 1:
                        WebDriverWait(self.driver, 240).until(EC.alert_is_present(),
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
                            WebDriverWait(self.driver, 240).until(EC.alert_is_present(),
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
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
            time.sleep(1)
        self.driver.close()

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
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
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
                    self.driver.execute_script("javascript:doAction(document.form1, gRsvWGetCancelRsvDataAction);")
                    # TODO: 当選確定済の当選結果 のみ出力させたい
                    time.sleep(3)
                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    result_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue
            # ログアウト
            self.driver.execute_script("javascript:doAction(document.form1, gRsvWTransUserAttestationEndAction);")
            time.sleep(1)

        self.driver.close()
        # if output_csv_path != "":
        #     mi.output_csv_from_id_dict(result_dict, output_csv_path)

        return result_dict

    def check_court(self, month):
        """
        コートの空き状況をチェック
        """
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(service=Service(driver_path), options=options)
        self.driver.get(top_url)
        # フレーム移動
        self.driver.switch_to.frame("pawae1002")
        # 空き状況ページへ移動
        self.driver.execute_script("javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formWTransInstSrchVacantAction : document.formWTransInstSrchVacantAction ), gRsvWTransInstSrchVacantAction);")
        self.driver.execute_script("javascript:doComplexSearchAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction);")
        try:
            self.driver.find_element_by_name("monthGif" + month).click() # 月選択
        except:
            print("対象月が存在しません")
            self.driver.quit()
            exit()
        # 曜日選択 土曜固定
        self.driver.find_element_by_name("weektype5").click()
        self.driver.execute_script("javaScript: sendSelectWeekNum2((_dom == 3) ? document.layers['disp'].document.form1: document.form1, gRsvWTransInstSrchPpsAction);")
        self.driver.execute_script("javascript:doTransInstSrchMultipleAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction, '1000', '1030');")
        # 場所選択 府中の森固定
        self.driver.find_element_by_name("gifName23").click()
        self.driver.execute_script("javascript:sendSelectWeekNum((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWGetInstSrchInfAction);")
        print(self.driver.page_source)
        # TODO ページの保存
        
    
def main():
    """
    main
    """
    root = tk.Tk()
    cr = Court_Reserv(master=root)
    cr.mainloop()
    
if __name__ == '__main__':
    main()