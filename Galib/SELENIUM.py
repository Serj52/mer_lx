"""Импорты"""
import os
import logging as lg
from datetime import datetime
import time

from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from Galib.ERROR_HANDLER import ErrorHandler
from config import Config


class Connection:
    """Класс по настройке соединения"""




    def __init__(self):
        self.config = Config()
        self.options = webdriver.ChromeOptions()
        self.browser_path = self.config.browser_path
        self.driver_path = self.config.driver_path


    def set_options(self, folder_load):
        # Без графического интерфейса браузера
        self.options.add_argument("--headless")
        self.options.add_argument("--no-sandbox")
        self.options.add_argument("--disable-dev-shm-usage")
        self.options.add_argument("--window-size=1920x1080")

        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])
        # self.options.binary_location = self.browser_path
        self.options.add_experimental_option('prefs', {'download.default_directory': folder_load,
                                                       "safebrowsing.enabled": False,
                                                       "download.prompt_for_download": False,
                                                       "download.directory_upgrade": True,
                                                       })


    # def set_options(self, folder_load):
    #     """Установка опций вебдрайвера"""
    #     self.options.add_experimental_option('excludeSwitches', ['enable-logging'])
    #     self.options.binary_location = self.browser_path
    #     self.options.add_experimental_option('prefs', {'download.default_directory': folder_load,
    #                                                    "safebrowsing.enabled": False,
    #                                                    "download.prompt_for_download": False,
    #                                                    "download.directory_upgrade": True,
    #                                                    })
    #     # Без графического интерфейса браузера
    #     # self.options.add_argument("--headless")

    def set_driver(self):
        """Создание вебдрайвера"""
        # self.driver = webdriver.Chrome(options=self.options, executable_path=self.driver_path)
        self.driver = webdriver.Chrome(options=self.options)


class BaseSelenium:
    """Класс по работе с Selenium"""
    def __init__(self):
        self.conn = Connection()
        self.config = Config()

    def open_site(self, folder_load, site_url):
        # """Открытие сайта"""
        # eh = ErrorHandler('open_site', self.config, tries_count=2)
        # while True:
        #     with eh:
                self.conn.set_options(folder_load)
                self.conn.set_driver()
                lg.info(f'Open URL:{site_url}')
                print(f'{datetime.now()} Open URL:{site_url}')
                self.conn.driver.get(site_url)
                # break

    def close_site(self):
        """Закрытие сайта"""
        try:
            self.conn.driver.implicitly_wait(2)
            lg.info("Идет завершение сессии...")
            print("Идет завершение сессии...")
            self.conn.driver.quit()
            lg.info("Драйвер успешно завершил работу.")
            print("Драйвер успешно завершил работу.")
        except Exception as error:
            lg.exception(error, exc_info=True)
            os.system("tskill chrome")
            lg.info('Chrome браузер закрыт принудительно.')
            os.system("tskill chromedriver")
            lg.info('Chrome драйвер закрыт принудительно.')

    def find_by_xpath(self, selector, timeout=None):
        """Возвращает элемент по xpath"""
        eh = ErrorHandler('download_file', self.config, tries_count=2)
        while True:
            with eh:
                if timeout:
                    wt = WebDriverWait(self.conn.driver, timeout=timeout)
                    # Ожидание загрузки тела страницы
                    wt.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    return self.conn.driver.find_element(By.XPATH, selector)
                return self.conn.driver.find_element(By.XPATH, selector)

    def download_file(self, selector, timeout=None):
        """в этом блоке скачивается файл"""
        eh = ErrorHandler('download_file', self.config, tries_count=2)
        while True:
            with eh:
                lg.info('Начинаю загрузку данных')
                print(f'{datetime.now()} Начинаю загрузку данных')
                self.find_by_xpath(selector, timeout).click()
                # Ожидание завершения загрузки 5с
                time.sleep(5)
                lg.info('Данные загружены')
                print(f'{datetime.now()} Данные загружены')
                break

    def get_screen_shot(self, screen_path):
        """Сделать скриншот"""
        time.sleep(5)
        date = datetime.now().strftime("%d.%m.%Y %H-%M")
        name = f'{date}.png'
        self.conn.driver.save_screenshot(os.path.join(screen_path, name))

    def switch_to_active_tab(self):
        """переключиться на последнее открытое окно"""
        self.conn.driver.switch_to.window(self.conn.driver.window_handles[-1])

    def switch_to_main(self):
        """переключиться на главное окно"""
        self.conn.driver.switch_to.window(self.conn.driver.window_handles[0])

    def get_params_to_attach(self):
        """получить идентификаторы текущей сессии"""
        if self.conn.driver:
            return self.conn.driver.command_executor._url, self.conn.driver.session_id
        pass
        # raise ChromeDriverNotFoundException()

    def attach_to_session(self, executor_url, session_id):
        """подключиться к существующей сессии"""
        original_execute = WebDriver.execute

        def new_command_execute(self, command, params=None):
            if command == "newSession":
                return {'success': 0, 'value': None, 'sessionId': session_id}
            return original_execute(self, command, params)

        WebDriver.execute = new_command_execute
        driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
        driver.session_id = session_id
        WebDriver.execute = original_execute
        return driver