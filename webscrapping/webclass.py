class clssAccessWebsite():
    base_url = "https://xxxxxx"
    
    def setUp(self):
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.implicitly_wait(10)
        
    def load_home_page(self):
        driver = self.driver
        driver.get(self.base_url)
        
        
    def reset_page(self):
        time.sleep(4)
#         id_submit=self.driver.find_element(By.XPATH,  '//*[@id="searchbox"]')

        self.driver.find_element_by_xpath('//*[@id="searchBox"]').click()
    
        self.driver.find_element_by_xpath('//*[@id="searchBox"]').send_keys("xxx")
        self.driver.find_element_by_xpath('//*[@id="searchBox"]').send_keys(Keys.RETURN)
        time.sleep(1)
        
