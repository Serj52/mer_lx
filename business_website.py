import time
import logging as lg
from Galib.SELENIUM import BaseSelenium, Connection
from selenium.common.exceptions import NoSuchElementException



class Selenium(BaseSelenium):
    def __init__(self):
        super().__init__()

    def exists_by_xpath(self, xpath):
        """Проверяет наличие элемента на странице"""
        try:
            self.find_by_xpath(xpath)
            return True
        except NoSuchElementException:
            return False


class Website:
    """Класс для построения бизнес процесса при работе с сайтом"""
    def __init__(self, config):
        self.config = config
        self.web = Selenium()
        self.current_year = config.current_year

    def work_with_site(self, folder_load, url, forecast, screen_path):
        """Метод для загрузке данных с сайта"""
        count = 1
        #В цикле проходим по страницам сайта ищем элемент
        xpath = self.return_xpath(forecast)
        while True:
            full_url = f'{url}?page={count}'

            self.web.open_site(folder_load, site_url=full_url)
            time.sleep(3)
            # Если страница содержит нужный div, то ищем Прогноз
            if self.web.exists_by_xpath(xpath['div']):
                # Если страница содержит нужный прогноз
                if self.web.exists_by_xpath(xpath['link']):
                    element = self.web.find_by_xpath(xpath['link'])
                    #Проверяем год выгрузки прогноза
                    date = self.web.find_by_xpath(f"{xpath['link']}//small[@class='e-date']").text
                    # Если выгрузка текущего года, забираем. !!!! Проверить необходимость данной проверки после пояпвления ОПР
                    if str(self.current_year) in date:
                        print(f'Найден {element.text}')
                        lg.info(f'Найден {element.text}')
                        element.click()
                        self.web.get_screen_shot(screen_path)
                        self.web.download_file(xpath['downfile'])
                        self.web.close_site()
                        return True

                    count += 1
                    self.web.close_site()
                    print('Ищем дальше')
                    lg.info('Ищем дальше')
                    continue
                # Если Прогноз на странице не найден, то переходим на следующую
                else:
                    count += 1
                    self.web.close_site()
                    print('Ищем дальше')
                    lg.info('Ищем дальше')
                    continue
            else:
                # Если на cтранице нет блока div прекращаем поиск
                self.web.close_site()
                return False

    def return_xpath(self, name):
        """По имени Прогноза возвращает xpath"""
        self.config.xpath['link'] = self.config.xpath_link[name]
        return self.config.xpath