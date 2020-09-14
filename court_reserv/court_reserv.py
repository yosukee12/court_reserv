# -*- coding: utf-8 -*-
import configparser
import time
import logging
import datetime
import sys
import manage_id as mi
import tkinter as tk
from tkinter import ttk
from functools import partial

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
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
logfile = config['PATH']['LOG_PATH'] + '/court_reserv_{0}.log'.format(datetime.date.today())
log_fmt = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(filename=logfile, format=log_fmt, level=logging.INFO)

# Output csv path
output_csv_file = config['PATH']['OUTPUT_CSV_PATH'] + '/check_reserv_{0}.csv'.format(datetime.date.today())

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
    def __init__(self, input_csv_path, master=None):
        """
        コンストラクタ
          引数でIDリストのCSVファイルを指定してID dictに変換
          Tkinterのウィジェット作成
        """
        self.id_dict = mi.Manage_Id.get_id_dict_from_csv(input_csv_path)
        # tkinter
        super().__init__(master)
        self.pack()
        self.master.geometry("300x300")
        self.master.title("Court Reservation")

        self.create_widgets()

    def create_widgets(self):
        #Check Court Button
        self.button_check_court = ttk.Button(self)
        self.button_check_court.configure(text="空きコート確認")
        self.button_check_court.configure(command = partial(self.check_court, "9"))
        self.button_check_court.pack()

        #Check Reserv Button
        self.button_check_reserv = ttk.Button(self)
        self.button_check_reserv.configure(text="抽選申込み状況確認")
        self.button_check_reserv.configure(command = partial(self.check_reserv, self.id_dict, output_csv_file))
        self.button_check_reserv.pack()

        #SemiAuto Reserv Button
        self.button_semiauto_reserv = ttk.Button(self)
        self.button_semiauto_reserv.configure(text="抽選申込み")
        self.button_semiauto_reserv.configure(command = partial(self.semiauto_reserv, self.id_dict))
        self.button_semiauto_reserv.pack()

    #     #Label
    #     self.label_hello = ttk.Label(self)
    #     self.label_hello.configure(text='A Label')
    #     self.label_hello.pack()
    #
    #     #Entry
    #     self.name = tk.StringVar()
    #     self.entry_name = ttk.Entry(self)
    #     self.entry_name.configure(textvariable = self.name)
    #     self.entry_name.pack()
    #
    #     #Label2
    #     self.label_name=ttk.Label(self)
    #     self.label_name.configure(text = 'Please input something in Entry')
    #     self.label_name.pack()
    
    def check_court(self, month):
        """
        コートの空き状況をチェック
        """
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
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

    def check_reserv(self, id_dict={}, output_csv_path=""):
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
        self.driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
        for k, v in id_dict.items():
            self.driver.get(config['URL']['TOP_URL'])
            # フレーム移動
            self.driver.switch_to.frame("pawae1002")
            # ログインページへ移動
            try:
                self.driver.execute_script("javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formdisp : document.formdisp ), gRsvLoginUserAction);")
                self.driver.find_element_by_name("userId").send_keys(k)
                self.driver.find_element_by_name("password").send_keys(v[2])
                time.sleep(5)
                # ログイン
                self.driver.find_element_by_xpath("//*[contains(@href, 'submitLogin')]").click()
                # 有効期限が近づいている画面が出た場合
                if "お知らせ画面" in self.driver.title:
                    if "利用者カードの有効期限が切れている" in self.driver.page_source:
                        print("ID:" + k + " 期限切れ")
                        continue
                    else:
                        self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
            except UnexpectedAlertPresentException:
                print("ID:" + k + " 期限切れ")
                logging.warning("ID:" + k + " 期限切れ")
                continue

            if "登録メニュー画面" in self.driver.title:
                try:
                    # 抽選申し込み確認画面へ
                    self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gLotWTransCompleteLotListAction);")
                    soup = bs(self.driver.page_source)
                    # Beautiful soupで申込み日と時間の取得
                    found_list = [elem.text for elem in soup.find_all('td', class_='tablelist')]
                    if len(found_list) == 8:
                        print("ID:" + k + " 申込み日1→ " + found_list[2] + " " + found_list[3])
                        print("ID:" + k + " 申込み日2→ " + found_list[6] + " " + found_list[7])
                        reserv_dict[k] = [v[0], v[1], v[2], found_list[2] + " " + found_list[3], found_list[6] + " " + found_list[7]]
                    elif len(found_list) == 4:
                        print("ID:" + k + " 申込み日1→ " + found_list[2] + " " + found_list[3])
                        reserv_dict[k] = [v[0], v[1], v[2], found_list[2] + " " + found_list[3], ""]
                    else:
                        print("ID:" + k + " 申込みなし")
                        reserv_dict[k] = [v[0], v[1], v[2], "", ""]
                except UnexpectedAlertPresentException:
                    print("ID:" + k + " 申込みなし")
                    reserv_dict[k] = [v[0], v[1], v[2], "", ""]
                    continue

        self.driver.close()
        if output_csv_path != "":
            mi.Manage_Id.output_csv_from_id_dict(reserv_dict, output_csv_path)
        return reserv_dict

    def semiauto_reserv(self, id_dict={}):
        """
        IDリストを引数にして
        半自動抽選申込み. 抽選申込み日の選択と申込みは手動
        """
        # 引数でID dictを指定しない場合
        if not id_dict:
            id_dict = self.id_dict
            
        # Chromeドライバーの起動
        self.driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
        for k, v in id_dict.items():
            reserv_count = 0
            self.driver.get(config['URL']['TOP_URL'])
            # フレーム移動
            self.driver.switch_to.frame("pawae1002")
            # ログインページへ移動
            try:
                self.driver.execute_script("javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formdisp : document.formdisp ), gRsvLoginUserAction);")
                self.driver.find_element_by_name("userId").send_keys(k)
                self.driver.find_element_by_name("password").send_keys(v[2])
                time.sleep(5)
                # ログイン
                self.driver.find_element_by_xpath("//*[contains(@href, 'submitLogin')]").click()
            except UnexpectedAlertPresentException:
                print("ID:" + k + " 期限切れ")
                logging.warning("ID:" + k + " 期限切れ")
                continue

            # 有効期限が近づいている画面が出た場合
            if "お知らせ画面" in self.driver.title:
                self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
                logging.warn("ID:" + k + " 期限が近くなっています")
            logging.info("ID:" + k + " ログイン")
            # time.sleep(0.5)
            if "登録メニュー画面" in self.driver.title:
                # 抽選申し込み画面へ
                self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gLotWSetupLotAcceptAction);")
                # 利用日
                self.driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), lotWTransLotAcceptListAction);")
                # 種目選択
                self.driver.execute_script("javascript:sendWTransLotBldGrpAction(((_dom == 3) ? document.layers[disp].document.form1 : document.form1 ), lotWTransLotBldGrpAction, 130);")
                # 公園選択
                self.driver.execute_script("javascript:sendBldGrpCd(((_dom == 3) ? document.layers[disp].document.form1 : document.form1 ), lotWTransLotInstGrpAction, 1301270)")
                while True:
                    # 申し込み中処理（手動申し込み）
                    try:
                        if "東京都スポーツ施設サービス" in self.driver.title:
                            logging.info("ID:" + k + " ログアウト")
                            break
                        elif "抽選申込完了確認画面" in self.driver.title:
                            reserv_count += 1
                            soup = bs(self.driver.page_source)
                            # Beautiful soupで申込み日と時間の取得
                            foundlist = [elem.text for elem in soup.find_all('td', class_='tablelist')]
                            print("ID:" + k + " 申込み" + str(reserv_count) + "完了→ " + " ".join(foundlist))
                            logging.info("ID:" + k + " 申込み" + str(reserv_count) + "完了→ " + " ".join(foundlist))
                            ## 2日分申込み完了したら次のIDへ
                            if reserv_count == 2:
                                break
                        # ポップアップアラートの表示待ち
                        WebDriverWait(self.driver, 240).until(EC.alert_is_present(),
                                                    'Timed out waiting for PA creation ' +
                                                    'confirmation popup to appear.')
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except TimeoutException or UnexpectedAlertPresentException:
                        continue

def main():
    # TODO ユーザの入力によって処理を選択
    # manage_id.output_csv_from_id_dict(
    #     check_reserv(manage_id.get_id_dict_from_csv("/Users/Yousuke/Documents/FFTC/test.csv")),
    #     "/Users/Yousuke/Documents/FFTC/test_output.csv")
    # alive_dict, dead_dict = manage_id.get_alive_dead_id_dict(manage_id.get_id_dict_from_csv("/Users/Yousuke/Documents/FFTC/test.csv"))
    # manage_id.output_csv_from_id_dict(alive_dict, "/Users/Yousuke/Documents/FFTC/test_output_alive.csv")
    # manage_id.output_csv_from_id_dict(dead_dict, "/Users/Yousuke/Documents/FFTC/test_output_dead.csv")
    root = tk.Tk()
    cr = Court_Reserv("/Users/Yousuke/Documents/FFTC/test_output_alive.csv", master=root)
    cr.mainloop()
if __name__ == '__main__':
    main()