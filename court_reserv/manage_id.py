# -*- coding: utf-8 -*-
import os
import csv
import time
import configparser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException

# Config
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

# Selenium Options
options = Options()
options.add_argument('--disable-gpu');
options.add_argument('--disable-extensions');
options.add_argument('--proxy-server="direct://"');
options.add_argument('--proxy-bypass-list=*');
options.add_argument('--start-maximized');

class Manage_Id():
    @staticmethod
    def get_id_dict_from_csv(csv_file_path):
        """
        IDリストのCSVファイルを読み込み、dictを返す
        CSV形式:
            ID,名前(漢字),名前(カタカナ),パスワード(生年月日)
        dict形式:
            {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日)]}
        """
        id_dict = {}
        if os.path.exists(csv_file_path):
            with open(csv_file_path) as f:
                for row in csv.reader(f):
                    # print(f"{row}")
                    if len(row) > 3 and row[0] != "" and row[3] != "":
                        # IDとパスワードの長さが8文字以外だと飛ばす
                        if len(row[0]) == 8 and len(row[3]) == 8:
                            # TODO ここ下手くそだから最適化する
                            if len(row) == 4:
                                id_dict.update({row[0]:[row[1], row[2], row[3]]})
                            elif len(row) == 5:
                                id_dict.update({row[0]:[row[1], row[2], row[3], row[4
                                ]]})
                            elif len(row) == 6:
                                id_dict.update({row[0]:[row[1], row[2], row[3], row[4], row[5]]})

                        else:
                            print["ID: " + row[0] + ",pass: " + row[1] + "不正なID"]
                            continue
        else:
            print("csvファイルが存在しません")
            exit()
        return id_dict

    @staticmethod
    def output_csv_from_id_dict(id_dict, output_file_path):
        """
        任意のID dictを読み込み、csvを出力する
        CSV形式:
            ID,名前(漢字),名前(カタカナ),パスワード(生年月日)
        """
        id_list = []
        with open(output_file_path, 'w') as f:
            writer = csv.writer(f)
            for k, v in id_dict.items():
                if k != "" and v[2] != "":
                    # TODO ここ下手くそなので最適化する
                    if len(v) == 3:
                        writer.writerow([k, v[0], v[1], v[2]])
                    elif len(v) == 4:
                        writer.writerow([k, v[0], v[1], v[2], v[3]])
                    elif len(v) == 5:
                        writer.writerow([k, v[0], v[1], v[2], v[3], v[4]])
                    else:
                        continue

    @staticmethod
    def get_alive_dead_id_dict(id_dict):
        """
        任意のID dictを読み込み、有効なIDと有効期限切れIDのdictを返す
        戻り値:
            [alive_id_dict, dead_id_dict]
        """
        dead_id_dict = {} # 期限切れIDのdict
        dead_soon_id_dict = {} # 期限切れが近いIDのdict
        alive_id_dict = {} # 有効なIDのdict (返さないが一応）
        # Chromeドライバーの起動
        driver = webdriver.Chrome(executable_path=config['PATH']['DRIVER_PATH'], chrome_options=options)
        for k, v in id_dict.items():
            driver.get(config['URL']['TOP_URL'])
            # フレーム移動
            driver.switch_to.frame("pawae1002")
            # ログインページへ移動
            try:
                driver.execute_script("javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formdisp : document.formdisp ), gRsvLoginUserAction);")
                driver.find_element_by_name("userId").send_keys(k)
                driver.find_element_by_name("password").send_keys(v[2])
                time.sleep(5)
                # ログイン
                driver.find_element_by_xpath("//*[contains(@href, 'submitLogin')]").click()
                if "お知らせ画面" in driver.title:
                    if "利用者カードの有効期限が切れている" in driver.page_source:
                        # 有効期限切れ画面の場合
                        print("ID:" + k + " 期限切れ")
                        dead_id_dict[k] = v
                        continue
                    else:
                        # 有効期限が近づいている画面の場合
                        dead_soon_id_dict[k] = v
                        print("ID:" + k + " 期限近い")
                        driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
            except UnexpectedAlertPresentException:
                # パスワードが間違っているポップアップが出た場合
                print("ID:" + k + " 期限切れ")
                dead_id_dict[k] = v
            else:
                print("ID:" + k + " 有効")
                alive_id_dict[k] = v
        driver.close()
        return alive_id_dict, dead_id_dict