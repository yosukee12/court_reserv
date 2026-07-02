# -*- coding: utf-8 -*-
"""Browser session helpers for legacy Selenium flows."""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait


class BrowserSession:
    """Encapsulate legacy WebDriver creation and teardown behavior."""

    def __init__(self, config):
        self.config = config
        self.driver_path = config["PATH"]["DRIVER_PATH"]

    def build_options(self):
        options = Options()
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument('--proxy-server="direct://"')
        options.add_argument("--proxy-bypass-list=*")
        options.add_argument("--start-maximized")
        return options

    def create_driver(self):
        return webdriver.Chrome(
            service=Service(self.driver_path),
            options=self.build_options(),
        )

    def get_wait(self, driver, timeout=10):
        return WebDriverWait(driver, timeout)

    def safe_close(self, driver):
        if driver is None:
            return
        try:
            driver.close()
        except Exception:
            pass

    def safe_quit(self, driver):
        if driver is None:
            return
        try:
            driver.quit()
        except Exception:
            pass
