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
    from .services import (
        LotteryService,
        ReservationService,
        AvailabilityService,
        IdManagerService,
    )
except Exception:
    # allow running the module as a script (no package context)
    from manage_id import Manage_Id as mi
    from config import get_debug_output_dir, load_config
    from browser import BrowserSession, LoginService, NavigationService
    from services import (
        LotteryService,
        ReservationService,
        AvailabilityService,
        IdManagerService,
    )
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
        self.reservation_service = ReservationService(
            config=config,
            browser_session=self.browser_session,
            navigation_service=self.navigation_service,
            logger=logging,
            get_id_dict_from_csv=mi.get_id_dict_from_csv,
            output_id_dict=mi.output_csv_from_id_dict,
            sleep_func=time.sleep,
        )
        self.availability_service = AvailabilityService(
            config=config,
            browser_session=self.browser_session,
            navigation_service=self.navigation_service,
            logger=logging,
            get_debug_output_dir=get_debug_output_dir,
            sleep_func=time.sleep,
        )
        self.id_manager_service = IdManagerService(
            config=config,
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
        return self.availability_service.collect_all_available_slots(
            self.driver,
            weeks_limit=weeks_limit,
            only_weekday=only_weekday,
        )

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
        self.semiauto_reserv(self.id_manager_service.load_accounts(self.entry_input_csv.get()))

    def start_reservation_button(self):
        """Mode-aware handler for reservation button: calls semi or full auto based on selection."""
        id_dict = self.id_manager_service.load_accounts(self.entry_input_csv.get())
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
        self.check_lottery(self.id_manager_service.load_accounts(self.entry_input_csv.get()), check_lottery_csv)

    def check_result_button(self):
        """
        抽選申込み結果確認ボタンが押された時の処理
        """
        self.check_result(self.id_manager_service.load_accounts(self.entry_input_csv.get()), check_result_csv)

    def determine_button(self):
        """
        予約確定ボタンが押された時の処理
        """
        self.determine_reserv(self.entry_result_csv.get(), determined_csv)

    def check_reserv_button(self):
        """
        抽選申込み結果確認ボタンが押された時の処理
        """
        self.check_reserv(self.id_manager_service.load_accounts(self.entry_result_csv.get()), check_reserv_csv)

    def check_id_button(self):
        """
        ID有効確認ボタンが押された時の処理
        """
        id_dict = self.id_manager_service.load_accounts(self.entry_check_id_csv.get())
        alive_id_list, dead_id_list = self.id_manager_service.check_account_validity(id_dict)
        self.id_manager_service.save_accounts(alive_id_list, alive_id_list_csv)
        self.id_manager_service.save_accounts(dead_id_list, dead_id_list_csv)

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
        return self.reservation_service.determine_reserv(
            input_csv_path,
            output_csv_path,
        )

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
        return self.reservation_service.check_reserv(id_dict, output_csv_path)

    def check_court(self, month):
        """
        コートの空き状況をチェック
        """
        return self.availability_service.check_court(month)
        
    
def main():
    """
    Compatibility GUI entrypoint for `python court_reserv/court_reserv.py`.
    """
    root = tk.Tk()
    cr = Court_Reserv(master=root)
    cr.mainloop()
    
if __name__ == '__main__':
    main()
