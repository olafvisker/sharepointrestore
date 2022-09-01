from time import sleep
from typing import List
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromiumService
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import argparse

class ScElem():
    def __init__(self, driver: WebDriver, el: WebElement) -> None: 
        self.driver = driver
        self.el = el

        self.visible = self.el.is_displayed()
        self.enabled = self.el.is_enabled()
        self.text = self.el.text

    def click(self): self.driver.execute_script("arguments[0].click();", self.el)

    def __str__(self) -> str: return self.el.__str__()

class Scraper():
    def __init__(self, url, headless=False, maximize=False, implicite_wait=30) -> None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        if maximize: options.add_argument("--start-fullscreen")
        if headless: options.add_argument("--headless")
        self.driver = webdriver.Chrome(options=options, service=ChromiumService(ChromeDriverManager().install()))
        self.driver.implicitly_wait(implicite_wait)
        if maximize: self.driver.fullscreen_window()
        self.driver.get(url)

    def get(self, by, value, visible=False, enabled=False) -> ScElem: 
        sc_els = self.get_all(by, value, visible, enabled)
        if len(sc_els) > 0: return sc_els[0]
        return None

    def get_all(self, by, value, visible=False, enabled=False) -> List[ScElem]: 
        els = WebDriverWait(self.driver, 30).until(EC.presence_of_all_elements_located((by, value)))
        sc_els = []
        for el in els:
            if visible and not el.is_displayed(): continue
            if enabled and not el.is_enabled(): continue
            sc_els.append(ScElem(self.driver, el))
        return sc_els

    def scroll_to_bottom(self): self.driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
    def run(self): raise NotImplementedError()
    def wait(self, time): sleep(time)
    def exit(self): self.driver.quit()

class RestoreSharepoint(Scraper):
    def __init__(self, url, headless=False) -> None:
        super().__init__(url, headless, True, 10)
        self.no_messages_id = "ctl00_PlaceHolderMain_noItemsMessage"
        self.select_all_id = "idSelectAll"
        self.restore_selection_js = 'Renderer.FunctionDispatcher.Execute(this,0,"itemClick",event,Renderer.FunctionDispatcher.GetObject(0))'

    def run(self):
        while (self._items_left()):
            self._select_all()
            self._restore_selection()
            self.driver.switch_to.alert.accept()
        
    def _select_all(self):
        print("  > Selecting all items...")
        select_all_btn = self.get(By.ID, self.select_all_id)
        self.wait(1)
        select_all_btn.click()
        self.wait(1)

    def _restore_selection(self):
        print("  > Restoring selection...")
        self.driver.execute_script(self.restore_selection_js)
        self.wait(1)

    def _items_left(self):
        self.wait(1)
        try: 
            self.driver.find_element(By.ID, self.no_messages_id)
            return False
        except: return True

parser = argparse.ArgumentParser(description='Restore sharepoint items from recycle bin')
parser.add_argument('-u', '--url', required=True, help='url to recycle from e.g. https://356tno.sharepoint.com/teams/***/_layouts/*/RecycleBin.aspx')
args = parser.parse_args()

def main():
    print('\n------------------------\n|  RESTORE SHAREPOINT  |\n------------------------')
    rs = RestoreSharepoint(args.url)

    print('- Restoring...')
    rs.run()
    print('- Done!')
    rs.exit()

if __name__ == "__main__": main()