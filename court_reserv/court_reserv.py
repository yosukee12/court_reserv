import configparser
import time
import logging
import datetime
import manage_id

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException

from bs4 import BeautifulSoup

# Config
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

# Log
logfile = config['PATH']['LOG_PATH'] + '/court_reserv_{0}.log'.format(datetime.date.today())
log_fmt = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(filename=logfile, format=log_fmt, level=logging.INFO)

# Selenium Options
options = Options()
options.add_argument('--disable-gpu');
options.add_argument('--disable-extensions');
options.add_argument('--proxy-server="direct://"');
options.add_argument('--proxy-bypass-list=*');
options.add_argument('--start-maximized');

def check_court(month):
    """
    コートの空き状況をチェック
    """
    driver_path = config['PATH']['DRIVER_PATH']
    top_url = config['URL']['TOP_URL']

    # Chromeドライバーの起動
    driver = webdriver.Chrome(executable_path=config['PATH']['DRIVER_PATH'], chrome_options=options)
    driver.get(config['URL']['TOP_URL'])
    # フレーム移動
    driver.switch_to.frame("pawae1002")
    # 空き状況ページへ移動
    driver.execute_script("javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formWTransInstSrchVacantAction : document.formWTransInstSrchVacantAction ), gRsvWTransInstSrchVacantAction);")
    driver.execute_script("javascript:doComplexSearchAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction);")
    try:
        driver.find_element_by_name("monthGif" + month).click() # 月選択
    except:
        print("対象月が存在しません")
        driver.quit()
        exit()
    # 曜日選択 土曜固定
    driver.find_element_by_name("weektype5").click()
    driver.execute_script("javaScript: sendSelectWeekNum2((_dom == 3) ? document.layers['disp'].document.form1: document.form1, gRsvWTransInstSrchPpsAction);")
    driver.execute_script("javascript:doTransInstSrchMultipleAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction, '1000', '1030');")
    # 場所選択 府中の森固定
    driver.find_element_by_name("gifName23").click()
    driver.execute_script("javascript:sendSelectWeekNum((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWGetInstSrchInfAction);")
    print(driver.page_source)
    # TODO ページの保存
    time.sleep(300)

def check_reserv(id_dict):
    """
    IDリストを読み込んで抽選申込み日を取得
    IDに申込み日を追加したdictを返す
    dict形式:
        {ID, [名前(漢字),名前(カタカナ),パスワード(生年月日),申込日1,申込み日2]}
    """
    reserv_dict = {}
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
            # 有効期限が近づいている画面が出た場合
            if "お知らせ画面" in driver.title:
                if "利用者カードの有効期限が切れている" in driver.page_source:
                    print("ID:" + k + " 期限切れ")
                    continue
                else:
                    driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
        except UnexpectedAlertPresentException:
            print("ID:" + k + " 期限切れ")
            logging.warning("ID:" + k + " 期限切れ")

        if "登録メニュー画面" in driver.title:
            # 抽選申し込み確認画面へ
            driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gLotWTransCompleteLotListAction);")
            soup = BeautifulSoup(driver.page_source)
            # Beautiful soupで申込み日と時間の取得
            foundlist = [elem.text for elem in soup.find_all('td', class_='tablelist')]
            if len(foundlist) == 8:
                print("ID:" + k + " 申込み日1→ " + foundlist[2] + " " + foundlist[3])
                print("ID:" + k + " 申込み日2→ " + foundlist[6] + " " + foundlist[7])
                reserv_dict[k] = [v[0], v[1], v[2], foundlist[2] + " " + foundlist[3], foundlist[6] + " " + foundlist[7]]
            elif len(foundlist) == 4:
                print("ID:" + k + " 申込み日1→ " + foundlist[2] + " " + foundlist[3])
                reserv_dict[k] = [v[0], v[1], v[2], foundlist[2] + " " + foundlist[3], ""]
            else:
                print("ID:" + k + " 申込みなし")
                reserv_dict[k] = [v[0], v[1], v[2], "", ""]
    return reserv_dict

def semiauto_reserv(id_dict):
    """
    半自動抽選申込み. 抽選申込み日の選択と申込みは手動
    """
    # Chromeドライバーの起動
    driver = webdriver.Chrome(executable_path=config['PATH']['DRIVER_PATH'], chrome_options=options)
    for k, v in id_dict.items():
        reserv_count = 0
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
        except UnexpectedAlertPresentException:
            print("ID:" + k + " 期限切れ")
            logging.warning("ID:" + k + " 期限切れ")
            continue

        # 有効期限が近づいている画面が出た場合
        if "お知らせ画面" in driver.title:
            driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gRsvWUserMessageAction);")
            logging.warn("ID:" + k + " 期限が近くなっています")
        logging.info("ID:" + k + " ログイン")
        # time.sleep(0.5)
        if "登録メニュー画面" in driver.title:
            # 抽選申し込み画面へ
            driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), gLotWSetupLotAcceptAction);")
            # 利用日
            driver.execute_script("javascript:doAction(((_dom == 3) ? document.layers['disp'].document.form1 : document.form1 ), lotWTransLotAcceptListAction);")
            # 種目選択
            driver.execute_script("javascript:sendWTransLotBldGrpAction(((_dom == 3) ? document.layers[disp].document.form1 : document.form1 ), lotWTransLotBldGrpAction, 130);")
            # 公園選択
            driver.execute_script("javascript:sendBldGrpCd(((_dom == 3) ? document.layers[disp].document.form1 : document.form1 ), lotWTransLotInstGrpAction, 1301270)")
            while True:
                # 申し込み中処理（手動申し込み）
                try:
                    if "東京都スポーツ施設サービス" in driver.title:
                        logging.info("ID:" + k + " ログアウト")
                        break
                    elif "抽選申込完了確認画面" in driver.title:
                        reserv_count += 1
                        soup = BeautifulSoup(driver.page_source)
                        # Beautiful soupで申込み日と時間の取得
                        foundlist = [elem.text for elem in soup.find_all('td', class_='tablelist')]
                        print("ID:" + k + " 申込み" + str(reserv_count) + "完了→ " + " ".join(foundlist))
                        logging.info("ID:" + k + " 申込み" + str(reserv_count) + "完了→ " + " ".join(foundlist))
                        ## 2日分申込み完了したら次のIDへ
                        if reserv_count == 2:
                            break
                    # ポップアップアラートの表示待ち
                    WebDriverWait(driver, 240).until(EC.alert_is_present(),
                                                'Timed out waiting for PA creation ' +
                                                'confirmation popup to appear.')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException or UnexpectedAlertPresentException:
                    continue

def main():
    # TODO ユーザの入力によって処理を選択
    # manage_id.output_csv_from_id_dict(
    #     check_reserv(manage_id.get_id_dict_from_csv("/Users/Yousuke/Documents/FFTC/test.csv")),
    #     "/Users/Yousuke/Documents/FFTC/test_output.csv")
    alive_dict, dead_dict = manage_id.get_alive_dead_id_dict(manage_id.get_id_dict_from_csv("/Users/Yousuke/Documents/FFTC/test.csv"))
    manage_id.output_csv_from_id_dict(alive_dict, "/Users/Yousuke/Documents/FFTC/test_output_alive.csv")
    manage_id.output_csv_from_id_dict(dead_dict, "/Users/Yousuke/Documents/FFTC/test_output_dead.csv")

if __name__ == '__main__':
    main()