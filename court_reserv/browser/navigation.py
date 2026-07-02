# -*- coding: utf-8 -*-
"""Navigation helpers for legacy Selenium flows."""

import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select


class NavigationService:
    """Encapsulate repeated JavaScript execution and page transitions."""

    def __init__(self, wait_factory, sleep_func=time.sleep):
        self.wait_factory = wait_factory
        self.sleep_func = sleep_func

    def execute_script(self, driver, script, *args):
        return driver.execute_script(script, *args)

    def do_action(self, driver, action_name):
        return self.execute_script(
            driver,
            f"javascript:doAction(document.form1, {action_name});",
        )

    def logout(self, driver):
        return self.do_action(driver, "gRsvWTransUserAttestationEndAction")

    def go_to_lottery_entry(self, driver):
        return self.do_action(driver, "gLotWOpeLotSearchAction")

    def select_lottery_tennis_park(
        self,
        driver,
        park_value="1301270",
        court_value="12700020",
    ):
        self.execute_script(driver, "javascript:doLotEntry('130');")
        wait = self.wait_factory(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "bname")))
        Select(driver.find_element(By.ID, "bname")).select_by_value(park_value)
        self.execute_script(driver, "changeBname(document.form1);")
        wait.until(
            lambda d: any(
                opt.get_attribute("value") == court_value
                for opt in Select(d.find_element(By.ID, "iname")).options
            )
        )
        Select(driver.find_element(By.ID, "iname")).select_by_value(court_value)

    def go_to_lottery_cancel_list(self, driver):
        return self.do_action(driver, "gLotWTransLotCancelListAction")

    def go_to_lottery_result_list(self, driver):
        return self.do_action(driver, "gLotWTransLotElectListAction")

    def go_to_reservation_list(self, driver):
        return self.do_action(driver, "gRsvWGetCancelRsvDataAction")

    def go_to_temp_apply(self, driver):
        return self.execute_script(
            driver,
            "javascript:doApplay(document.form1, gLotWInstTempLotApplyAction);",
        )

    def go_to_vacant_search(self, driver):
        self.execute_script(
            driver,
            "javaScript:doActionFrame(((_dom == 3) ? document.layers['disp'].document.formWTransInstSrchVacantAction : document.formWTransInstSrchVacantAction ), gRsvWTransInstSrchVacantAction);",
        )
        self.execute_script(
            driver,
            "javascript:doComplexSearchAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction);",
        )

    def select_weekly_vacant_conditions(self, driver):
        self.execute_script(
            driver,
            "javaScript: sendSelectWeekNum2((_dom == 3) ? document.layers['disp'].document.form1: document.form1, gRsvWTransInstSrchPpsAction);",
        )
        self.execute_script(
            driver,
            "javascript:doTransInstSrchMultipleAction((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWTransInstSrchMultipleAction, '1000', '1030');",
        )
        self.execute_script(
            driver,
            "javascript:sendSelectWeekNum((_dom == 3) ? document.layers['disp'].document.form1 : document.form1, gRsvWGetInstSrchInfAction);",
        )
