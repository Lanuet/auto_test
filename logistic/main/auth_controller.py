from .locator import LoginLocator


class AuthController:
    def __init__(self, driver) -> None:
        super().__init__()
        self.driver = driver

    def login(self, url, username, password):
        self.driver.get(url)
        self.driver.find_element_by_xpath(LoginLocator.INPUT_USERNAME).send_keys(username)
        self.driver.find_element_by_xpath(LoginLocator.INPUT_PASSWORD).send_keys(password)
        self.driver.find_element_by_xpath(LoginLocator.BTN_LOGIN).click()
