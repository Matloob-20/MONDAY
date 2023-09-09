def sheet(self):
    self.files.open_workbook(path=f"{os.getcwd()}/test.xlsx")
    data = self.files.read_worksheet_as_table(header=True)
    for num, j in enumerate(data, 1):
        # for i in range(1, 11):
        self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[1]', f'{j["First Name"]}')
        self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[2]', f'{j["Last Name"]}')
        self.browser.input_text(f'//div[1]/div/div[2]/figure/table/tbody/tr[{num}]/td[3]', f'{j["Sales"]}')

    self.browser.wait_and_click_button('//button[contains(text(),"Update")]')
    time.sleep(3)

