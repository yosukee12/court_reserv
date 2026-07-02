# -*- coding: utf-8 -*-
"""ID management and CSV helpers extracted from legacy manage_id.py."""

import csv
import os
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import UnexpectedAlertPresentException


class IdManagerService:
    """Encapsulate legacy ID CSV handling and validity checks."""

    def __init__(self, config, sleep_func=time.sleep):
        self.config = config
        self.sleep_func = sleep_func

    def _build_options(self):
        options = Options()
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument('--proxy-server="direct://"')
        options.add_argument("--proxy-bypass-list=*")
        options.add_argument("--start-maximized")
        return options

    def load_accounts(self, csv_file_path):
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
                    if len(row) > 3 and row[0] != "" and row[3] != "":
                        if len(row[0]) == 8:
                            if len(row) == 4:
                                id_dict.update({row[0]: [row[1], row[2], row[3]]})
                            elif len(row) == 5:
                                id_dict.update(
                                    {row[0]: [row[1], row[2], row[3], row[4]]}
                                )
                            elif len(row) == 6:
                                id_dict.update(
                                    {row[0]: [row[1], row[2], row[3], row[4], row[5]]}
                                )
                        else:
                            print["ID: " + row[0] + ",pass: " + row[1] + "不正なID"]
                            continue
        else:
            print("csvファイルが存在しません")
            exit()
        return id_dict

    def save_accounts(self, id_dict, output_file_path):
        """
        任意のID dictを読み込み、csvを出力する
        CSV形式:
            ID,名前(漢字),名前(カタカナ),パスワード(生年月日)
        """
        with open(output_file_path, "w") as f:
            writer = csv.writer(f)
            for key, values in id_dict.items():
                if key != "" and values[2] != "":
                    if len(values) == 3:
                        writer.writerow([key, values[0], values[1], values[2]])
                    elif len(values) == 4:
                        writer.writerow([key, values[0], values[1], values[2], values[3]])
                    elif len(values) == 5:
                        writer.writerow(
                            [key, values[0], values[1], values[2], values[3], values[4]]
                        )
                    else:
                        continue

    def check_account_validity(self, id_dict):
        """
        任意のID dictを読み込み、有効なIDと有効期限切れIDのdictを返す
        戻り値:
            [alive_id_dict, dead_id_dict]
        """
        dead_id_dict = {}
        dead_soon_id_dict = {}
        alive_id_dict = {}
        driver = webdriver.Chrome(options=self._build_options())
        try:
            for key, values in id_dict.items():
                driver.get(self.config["URL"]["TOP_URL"])
                driver.switch_to.frame("pawae1002")
                try:
                    driver.execute_script(
                        "javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formdisp : document.formdisp ), gRsvLoginUserAction);"
                    )
                    driver.page_source
                    driver.find_element(By.NAME, "userId").send_keys(key)
                    driver.find_element(By.NAME, "password").send_keys(values[2])
                    self.sleep_func(5)
                    driver.find_element(
                        By.XPATH, "//*[contains(@href, 'submitLogin')]"
                    ).click()
                    if "お知らせ画面" in driver.title:
                        if "利用者カードの有効期限が切れている" in driver.page_source:
                            print("ID:" + key + " 期限切れ")
                            dead_id_dict[key] = values
                            continue
                        dead_soon_id_dict[key] = values
                        print("ID:" + key + "," + values[1] + " 期限近い")
                        driver.execute_script(
                            "javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);"
                        )
                    if "伝言表示画面" in driver.title:
                        driver.execute_script(
                            "javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageNextAction);"
                        )
                except UnexpectedAlertPresentException:
                    print("ID:" + key + "," + values[1] + " 期限切れ")
                    dead_id_dict[key] = values
                else:
                    print("ID:" + key + "," + values[1] + " 有効")
                    alive_id_dict[key] = values
        finally:
            driver.close()
        return alive_id_dict, dead_id_dict

    def get_active_accounts(self, csv_file_path):
        """Load accounts and return active / expired splits."""
        return self.check_account_validity(self.load_accounts(csv_file_path))
