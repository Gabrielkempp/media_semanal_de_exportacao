from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class WebScraper:
    def __init__(self, url):
        self.url = url
        self.driver = None

    def setup_driver(self):
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service)

    def open_website(self):
        self.driver.get(self.url)

    def get_first_h4_date_text(self):
        first_h4_date_element = self.driver.find_element(By.CSS_SELECTOR, "h4.date")
        return first_h4_date_element.text

    def close_driver(self):
        self.driver.quit()        
