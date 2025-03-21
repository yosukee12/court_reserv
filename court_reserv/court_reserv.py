# -*- coding: utf-8 -*-
import configparser
import time
import logging
import datetime
import sys
import re
from manage_id import Manage_Id as mi
import tkinter as tk
from tkinter import ttk
from functools import partial

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException
from bs4 import BeautifulSoup as bs

# Config
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

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
        # Label CSV PATH
        self.label_csvpath = ttk.Label(self, text="CSVファイル出力先: " + config['PATH']['OUTPUT_CSV_PATH'], background="white")

        #Entry
        self.entry_input_csv = ttk.Entry(self)
        self.entry_input_csv.insert(tk.END, config['PATH']['OUTPUT_CSV_PATH'] + "/")

        # Label 1
        self.label1 = ttk.Label(self, text="1. 毎月1日〜10日", background="white")

        #SemiAuto Reserv Button
        self.button_semiauto_reserv = ttk.Button(self)
        self.button_semiauto_reserv.configure(text="抽選申込み")
        self.button_semiauto_reserv.configure(command = self.semiauto_reserv_button)

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

    # ここからボタン実行用メソッド
    def semiauto_reserv_button(self):
        """
        抽選申込みボタンが押された時の処理
        """
        self.semiauto_reserv(mi.get_id_dict_from_csv(self.entry_input_csv.get()))

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
                # 種目選択
                self.driver.execute_script("javascript:doLotEntry('130');")
                # 公園選択（府中の森公園）
                Select(self.driver.find_element(By.ID,"bname")).select_by_value("1301270")
                time.sleep(1)
                # 種目選択2回目（テニス（人工芝））
                Select(self.driver.find_element(By.ID,"iname")).select_by_value("12700020")
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
                            # 申込み実行 → rechapcha対策で手動クリックする
                            #self.driver.execute_script("javascript:sendLotApply(document.form1, gLotWInstLotApplyAction, event);")
                            #time.sleep(0.5)
                            while not "抽選メール送信完了画面" in self.driver.title:
                                try:
                                    #print("loop2: count = " + str(reserv_count) + ": title = " + self.driver.title)
                                    # ポップアップ処理
                                    WebDriverWait(self.driver, 60).until(EC.alert_is_present(),
                                                            'Timed out waiting for PA creation ' +
                                                            'confirmation popup to appear.')
                                    alert = self.driver.switch_to.alert
                                    alert.accept()
                                    #time.sleep(0.5)
                                    WebDriverWait(self.driver, 0.3).until(EC.alert_is_present(),
                                                            'Timed out waiting for PA creation ' +
                                                            'confirmation popup to appear.')
                                    time.sleep(0.3)
                                except TimeoutException or UnexpectedAlertPresentException:
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