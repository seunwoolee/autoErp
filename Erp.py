import pickle
from time import sleep

from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class ERP:
    def __init__(self):
        self.chrome = webdriver.Chrome('chromedriver')
        self.url = 'http://kcerpjas1:81/jde/E1Menu.maf?jdeowpBackButtonProtect=PROTECTED'
        self.main_iframe = 'e1menuAppIframe'
        self.main_iframe_id = '#e1menuAppIframe'
        self.login()

    def login(self):
        self.chrome.get(self.url)
        self.wait_until_element_clicked('#mainLoginTable > tbody > tr:nth-child(8) > td > input')
        self.chrome.find_element_by_id('User').send_keys('swl21803')
        self.chrome.find_element_by_id('Password').send_keys('tmddn132!')
        self.chrome.find_element_by_xpath('//*[@id="mainLoginTable"]/tbody/tr[7]/td/input').click()
        sleep(0.2)
        self.go_to_P0411()
        self.click_add()
        self.set_P0411_value()

    def wait_until_element_clicked(self, css_selector: str):
        try:
            print("##기다림 시작")
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
        self.go_to_main_iframe()
        self.go_to_sub_iframe('#wcFrame10')
        self.chrome.find_element_by_id('tileDescription_1').click()

    def click_add(self):
        self.go_to_main_iframe()
        self.wait_until_element_clicked('#hc_Add')
        self.chrome.find_element_by_id('hc_Add').click()

    def set_P0411_value(self):
        css_selector_company = 'input[title="회사"]'
        css_selector_supply_number = 'input[title="공급자 번호"]'
        css_selector_GL = 'input[title="G/L 일자"]'
        css_selector_price = '#G0_1_R0 > td:nth-child(4) > div > input'
        css_selector_GL_code = '#G0_1_R0 > td:nth-child(5) > div > input'
        css_selector_VAT_type = '#G0_1_R0 > td:nth-child(6) > div > input'
        css_selector_VAT_type2 = '#G0_1_R0 > td:nth-child(7) > div > input'
        css_selector_ok_button = '#hc_OK'
        self.wait_until_element_clicked(css_selector_company)

        self.set_value(css_selector_company, 1)
        self.set_value(css_selector_supply_number, 91362)  # TODO Excel 파일 가져오기(INPUT 처리)
        self.set_value(css_selector_GL, "20/01/15")  # TODO Excel 파일 가져오기(INPUT 처리)
        self.set_value(css_selector_price, 1078000)  # TODO Excel 파일 가져오기(INPUT 처리)
        self.set_value(css_selector_GL_code, "B10")  # TODO Excel 파일 가져오기(INPUT 처리)
        self.set_value(css_selector_VAT_type, "V")  # TODO Excel 파일 가져오기(INPUT 처리)
        self.set_value(css_selector_VAT_type2, "VAT")  # TODO Excel 파일 가져오기(INPUT 처리)
        self.chrome.find_element_by_css_selector(css_selector_ok_button).click()
        self.chrome.find_element_by_css_selector(css_selector_ok_button).click()

    def set_value(self, css_selector: str, value: any):
        input_element: WebElement = self.chrome.find_element_by_css_selector(css_selector)
        input_element.click()
        self.chrome.execute_script(f"$('{css_selector}').val('{value}')")


if __name__ == "__main__":
    erp = ERP()
