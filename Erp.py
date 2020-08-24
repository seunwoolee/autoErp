import pickle
import time
from time import sleep

from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.worksheet.worksheet import Worksheet


class ERP:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument('--window-size=1920,1080')
        self.chrome = webdriver.Chrome('chromedriver', chrome_options=chrome_options)
        self.url = 'http://kcerpjas1:81/jde/E1Menu.maf?jdeowpBackButtonProtect=PROTECTED'
        self.main_iframe = 'e1menuAppIframe'
        self.main_iframe_id = '#e1menuAppIframe'

    def login(self):
        MAX_WAIT = 5
        start_time = time.time()
        self.chrome.get(self.url)
        while True:
            try:
                login_button: WebElement = self.chrome.find_element_by_xpath('//*[@id="mainLoginTable"]/tbody/tr[7]/td/input')
                assert login_button.get_attribute('value'), "Sign In"
                self.chrome.find_element_by_id('User').send_keys('swl21803')
                # self.chrome.find_element_by_id('Password').send_keys('tmddn132!')
                self.chrome.find_element_by_id('Password').send_keys('kcfeed12!')
                self.chrome.find_element_by_xpath('//*[@id="mainLoginTable"]/tbody/tr[7]/td/input').click()
                return
            except (AssertionError, WebDriverException, TimeoutException) as e:
                if time.time() - start_time > MAX_WAIT:
                    raise e
                time.sleep(0.5)

        # self.chrome.get(self.url)
        # self.wait_until_element_clicked('#mainLoginTable > tbody > tr:nth-child(8) > td > input')
        # self.chrome.find_element_by_id('User').send_keys('swl21803')
        # self.chrome.find_element_by_id('Password').send_keys('tmddn132!')
        # self.chrome.find_element_by_xpath('//*[@id="mainLoginTable"]/tbody/tr[7]/td/input').click()

    def wait_until_element_clicked(self, css_selector: str):
        try:
            WebDriverWait(self.chrome, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector)))
        except TimeoutException:
            print('## 기다리다가 지친다')

    def go_to_main_iframe(self):
        self.wait_until_element_clicked(self.main_iframe_id)
        sub_frame = self.chrome.find_element_by_css_selector(f'iframe[id={self.main_iframe}]')
        self.chrome.switch_to_frame(sub_frame)

    def go_to_sub_iframe(self, css_selector: str):
        sub_frame = self.chrome.find_element_by_css_selector(css_selector)
        self.chrome.switch_to_frame(sub_frame)

    def go_to_P0411(self):
        sleep(0.3)
        self.wait_until_element_clicked(self.main_iframe_id)
        self.go_to_main_iframe()
        self.go_to_sub_iframe('#wcFrame10')
        self.chrome.find_element_by_id('tileDescription_1').click()

    def click_add(self):
        self.go_to_main_iframe()
        self.wait_until_element_clicked('#hc_Add')
        self.chrome.find_element_by_id('hc_Add').click()

    def set_P0411_value(self, sheet: Worksheet):
        assert type(sheet) != Worksheet, '데이터없음'

        css_selector_company = 'input[title="회사"]'
        css_selector_supply_number = 'input[title="공급자 번호"]'
        css_selector_GL = 'input[title="G/L 일자"]'
        css_selector_price = '#G0_1_R0 > td:nth-child(4) > div > input'
        css_selector_GL_code = '#G0_1_R0 > td:nth-child(5) > div > input'
        css_selector_VAT_type = '#G0_1_R0 > td:nth-child(6) > div > input'
        css_selector_VAT_type2 = '#G0_1_R0 > td:nth-child(7) > div > input'
        css_selector_bigo = '#G0_1_R0 > td:nth-child(10) > div > input'
        css_selector_ok_button = '#hc_OK'
        self.wait_until_element_clicked(css_selector_company)

        self.set_value(css_selector_company, 1)
        self.set_value(css_selector_supply_number, sheet[1].value)
        self.set_value(css_selector_GL, sheet[2].value)
        self.set_value(css_selector_bigo, sheet[10].value)
        sleep(0.3)
        self.set_value(css_selector_price, sheet[3].value)
        self.set_value(css_selector_GL_code, sheet[4].value)
        self.set_value(css_selector_VAT_type, sheet[5].value)
        self.set_value(css_selector_VAT_type2, sheet[6].value)
        self.ok_button_three_click(css_selector_ok_button)

    def set_next_P0411_value(self, sheet: Worksheet):
        sleep(1)
        css_selector_count_number = '#G0_1_R0 > td:nth-child(3) > div > input'
        css_selector_employee_type = '#G0_1_R0 > td:nth-child(8) > div > input'
        css_selector_employee_number = '#G0_1_R0 > td:nth-child(9) > div > input'
        css_selector_ok_button = '#hc_OK'

        self.set_value(css_selector_count_number, sheet[7].value)
        self.set_value(css_selector_employee_number, sheet[9].value)
        self.set_value(css_selector_employee_type, sheet[8].value)
        self.ok_button_db_click(css_selector_ok_button)

    def set_value(self, css_selector: str, value: any):
        input_element: WebElement = self.chrome.find_element_by_css_selector(css_selector)
        input_element.click()
        self.chrome.execute_script(f"$('{css_selector}').val('{value}')")

    def ok_button_db_click(self, css_selector: str):
        self.chrome.find_element_by_css_selector(css_selector).click()
        sleep(0.2)
        self.chrome.find_element_by_css_selector(css_selector).click()

    def ok_button_three_click(self, css_selector: str):
        """ 채무전표에서 db_click 으로 안될때 한번 더 클릭"""
        self.chrome.find_element_by_css_selector(css_selector).click()
        sleep(0.2)
        self.chrome.find_element_by_css_selector(css_selector).click()
        sleep(0.2)
        self.chrome.find_element_by_css_selector(css_selector).click()


if __name__ == "__main__":
    erp = ERP()
