import pytest
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
#
#
# # @pytest.fixture
# # def setup():
# #     print("start excuted")
# #     yield
# #     print("close the browser")
# #
# # def test_1(setup):
# #     print("test1 excuted")
# #
# # def test_2(setup):
# #     print("test12 excuted")
# #
# # def test_3(setup):
# #     print("test13 excuted")
#
# import pytest
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.support.ui import WebDriverWait
# import time
#
#
# driver = None
#
# @pytest.fixture
# def setup():
#     print("start browser")
#     global driver
#     # Use Service() to set ChromeDriver path properly
#     service_obj = Service(ChromeDriverManager().install())
#     driver = webdriver.Chrome(service=service_obj)
#     driver.maximize_window()
#     yield driver
#     time.sleep(4)
#     driver.quit()
#     print("close the browser")
#
# def test_1(setup):
#     setup.get("http://www.facebook.com")
#     print("test1 executed")
#
# def test_2(setup):
#     setup.get("http://www.google.com")
#     print("test2 executed")
#
# def test_3(setup):
#     setup.get("http://www.gmail.com")
#     print("test3 executed")
#



class Test:

    def test_method1(self):
        print("method1 is called")

    def test_method2(self):
        print("method2 is called")

