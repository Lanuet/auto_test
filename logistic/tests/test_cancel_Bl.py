from unittest import TestCase
from selenium import webdriver

from logistic.main.auth_controller import AuthController


class TestCancelBl(TestCase):
    def __init__(self, methodName='runTest'):
        super().__init__(methodName)
        self.driver = webdriver.Chrome("C:/Users/nguye/OneDrive/Documents/chromedriver.exe")

    def test_case_1(self):
        controller = AuthController(self.driver)
        controller.login("https://www.google.com.vn/", "lannt", "123")
        print("Verify login successful")
        self.assertEqual(self.driver.title, "google")

