import time
import os
from RPA.Browser.Selenium import Selenium
from RPA.Tables import Tables
from RPA.Excel.Files import Files


class Monday:
    def __init__(self):
        self.browser = Selenium()
        self.table = Tables()
        self.files = Files()
        self.list = []
        self.date = ["23 JAN,2022", "30 JAN,2022", "06 FEB,2022"]
    def monday(self):
        self.browser.open_available_browser('https://matloobahmad.monday.com/auth/login_monday/email_password',
                                            maximized=True)
        self.browser.input_text_when_element_is_visible('//input[@id="user_email"]', 'matloobahmad235@gmail.com')
        self.browser.input_text_when_element_is_visible('//input[@id="user_password"]', 'Matloob@42')
        self.browser.find_element('//button[@aria-label="Log in"]').click()
        time.sleep(5)
        self.browser.find_element('//div[@id="workspaces-button"]').click()
        time.sleep(3)
        for i in self.date:
            self.browser.click_element_when_visible('//span[@class="icon icon-v2-search"]')
            self.browser.input_text(f'//*[@id="board-header-view-bar"]//input', f'{i}')
            time.sleep(2)
            self.browser.click_element_when_visible('//div[@data-walkthrough-id="item-name-text"]')
            time.sleep(1)
            self.browser.click_element_when_visible('//span[contains(text(),"Files")]')
            time.sleep(1)
            self.browser.choose_file('//*[@id="slide-panel-container"]/div/div[2]/div[1]/div[2]/div/div/div[2]/input',
                                     "/home/saboor/monday/test.xlsx")
            self.browser.click_element_when_visible('//span[contains(text(),"Updates")]')
            time.sleep(2)
            self.browser.click_button_when_visible('//button[@class="new_post_placeholder"]')
            self.browser.click_element_when_visible(
                '//*[@id="pulse-content"]/div[1]/div[1]/div/div/div[1]/div/div[1]/div/a[10]/i')
            self.browser.click_element_when_visible('//span[contains(text(),"Insert table")]')
            time.sleep(2)
            for i in range(1, 9):
                self.browser.click_element_when_visible(
                    '//*[@id="pulse-content"]/div[1]/div[1]/div/div/div[1]/div/div[1]/div/a[10]/i')
                self.browser.click_element_when_visible('//span[contains(text(),"Insert row above")]')
                time.sleep(1)
            self.files.open_workbook(path=f"{os.getcwd()}/test.xlsx")
            data = self.files.read_worksheet_as_table(header=True)
            for num, j in enumerate(data, 1):
                self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[1]', f'{j["First Name"]}')
                self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[2]', f'{j["Last Name"]}')
                self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[3]', f'{j["Sales"]}')
            self.browser.click_element_when_visible('//*[@id="files-uploader-container"]/div/div/div[1]/div[2]/button')
            self.browser.click_element_when_visible('//button[contains(text(),"Update")]')
            time.sleep(1)
            try:
                self.browser.click_element_when_visible('//*[@id="notification_reminder_bar_close"]/i')
            except:
                pass
            self.browser.click_element_when_visible('//div[@id="slide-panel-container"]//button[@aria-label="Close"]')
            self.browser.click_element_when_visible('//span[contains(text(),"Working on it")]')
            self.browser.click_element_when_visible('//span[contains(text(),"Done")]')

            self.browser.click_element_when_visible('//span[@class="icon icon-dapulse-close clear-button"]')
        print('Done')


obj = Monday()
obj.monday()
