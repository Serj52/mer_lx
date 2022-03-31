"""Импорты"""
import logging as lg
import os
import pandas as pd
from business_website import Website
from config import Config
from zip_worker import Zipworker
# from outlook import Outlook
from Galib.LOGGER import Log
from exchangelib import Credentials, Account, Configuration, DELEGATE, IMPERSONATION, Message, Mailbox


class Business:
    """Класс реализации главной логики"""
    def __init__(self, config):
        self.config = config
        self.wb = Website(config)  # объект для работы с сайтом
        self.zp = Zipworker(config)  # объект для работы с архивом
        # self.otl = Outlook() # объект для работы с Outlook

    # def search_mail_out(self):
    #     """Функция проверки почты"""
    #     #Запуск почты
    #     self.otl.run()
    #     # Если письмо пришло возвращаем True
    #     if self.otl.get_mail_data_by_subjects(self.config):
    #         return True
    #     return False

    # def search_mail_out(self):
    #     """Функция проверки почты"""
    #     #Запуск почты
    #     self.otl.run()
    #     # Если письмо пришло возвращаем True
    #     if self.otl.get_mail_data_by_subjects(self.config):
    #         return True
    #     return False

    def search_mail_exc(self):
        server = 'mail.rosatom.ru'
        email = 'GREN-r-000004@Greenatom.ru'
        name = r'gk\GREN-r-000004'
        password = 'Bayern2024'
        credentials = Credentials(name, password)
        config = Configuration(server=server, credentials=credentials)
        account = Account(primary_smtp_address=email, config=config, credentials=credentials, autodiscover=True,
                          access_type=DELEGATE)


        for item in account.inbox.all().order_by('-datetime_received')[:100]:
            # {item.sender}, {item.datetime_received}
            # print(f'{count}. {item.subject}')
            if item.subject == 'Сценарные условия':
                print(f'Нашел письмо от {item.sender}')
                return True

    def data_process(self):
        """в этом блоке реализуется главная логика загрузки архивов с сайта и извлечению файлов"""
        print('Робот начал обработку сайта')
        lg.info('Робот начал обработку сайта')
        #Читаем данные из входного файла
        data_frame = pd.read_excel(io=self.config.file_path,
                                   engine='openpyxl',
                                   usecols=['Наименование прогноза', 'Ссылка на базу'])
        #Убираем пустые сроки в столбцах
        clean_data_frame = data_frame.dropna()
        # Создаем словарь с данными
        data_dict = clean_data_frame.to_dict(orient='records')
        # Перебираем Прогнозы
        for row in data_dict:
            name = row.get('Наименование прогноза')
            url = row.get('Ссылка на базу')
            folder_load = self.config.folder_load_forecast[name]
            screen_path = self.config.folder_screen_forecast[name]
            exctract_dir = self.config.folder_load_forecast[name]
        # Если прогнозы найдены на сайте
            if self.wb.work_with_site(folder_load=folder_load,
                                      url=url,
                                      forecast=name,
                                      screen_path=screen_path):
                # Извлекаем скаченный архив
                self.zp.unpack_zipfile(exctract_dir, self.config.zip_templates)
                # Если найден долгосрочный прогноз, перезаписываем год в файле
                if name == 'Долгосрочный прогноз':
                    with open('year.txt', 'r+') as file:
                        text = file.readline()
                        year = int(text) + 12
                        file.seek(0)
                        file.write(str(year))
            # Если данные не найдены
            else:
                lg.info(f"Прогноз {name} не найден на сайте.")
                print(f"Прогноз {name} не найден на сайте. ")
                # Outlook.send_mail(to=self.config.support_email,
                #                   body=f'{name} не найден на сайте')


if __name__ == '__main__':
    config = Config()
    lg = Log().set()
    robot = Business(config)
    robot.data_process()