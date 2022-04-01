# Default ignored files
"""Импорты"""
# import keyring  # для паролей
import os
from datetime import datetime

MODE = 'test'

def exctract_year(folder_root):
    """Функция извлечения значения года из файла"""
    path_tofile = os.path.join(folder_root, 'year.txt')
    with open(path_tofile) as file:
        text = file.readline()
        year = int(text) + 12
        return year

class Config:
    """Конфигурационный класс"""
    mode = MODE

    #Общие параметры
    robot_name = 'mer'
    theme_mail = 'Выгрузка индексов цен'
    folder_root = os.path.dirname(os.path.abspath(__file__))
    folder_load = os.path.join(folder_root, 'Load')
    folder_attachment = os.path.join(folder_root, 'Attachments')
    folder_logs = os.path.join(folder_root, 'Logs')
    file_path = os.path.join(folder_attachment, 'inputfile.xlsx')
    # Директории для хранения загруженных с сайта файлов
    folder_load_forecast = {
        'Сценарные условия': os.path.join(folder_load, 'Сценарные условия'),
        'Среднесрочный прогноз': os.path.join(folder_load, 'Среднесрочный прогноз'),
        'Долгосрочный прогноз': os.path.join(folder_load, 'Долгосрочный прогноз')
    }
    #Директории для хранения скриншотов
    folder_screen_forecast = {
        'Сценарные условия': os.path.join(folder_load_forecast['Сценарные условия'], 'Screen'),
        'Среднесрочный прогноз': os.path.join(folder_load_forecast['Среднесрочный прогноз'], 'Screen'),
        'Долгосрочный прогноз': os.path.join(folder_load_forecast['Долгосрочный прогноз'], 'Screen')
    }
    # Шаблон для поиска файлов в архиве  ______________________________
    zip_templates = {'Дефлятор базовый.xlsx': r'.{0,}Дефлятор.{0,}баз.{0,}',
                     'Внешняя торговля базовый.xlsx': r'.{0,}Внеш.{0,}торг.{0,}баз.{0,}'}
    # Текущий год нужен для поиска Среднесрочного прогноза
    current_year = datetime.now().year
    #Год из файла нужен для поиска Долгосрочного прогноза
    exctract_year = exctract_year(folder_root)
    #Общий параметры xpath для всех прогнозов.
    xpath = {'downfile': '//span[text()="Приложения"]',
             'div': '//div[@id="submaterails"]'}
    # индивидуальный xpath для каждого прогноза
    xpath_link = {'Сценарные условия': '//a[contains(@title,"Сценарные условия")]',
                       'Среднесрочный прогноз': f'//a[contains(@title, "Прогноз социально-экономического развития Российской Федерации на {current_year + 1} год и на плановый период {current_year + 2} и {current_year + 3} годов")]',
                       'Долгосрочный прогноз': f'//a[contains(@title, "Прогноз социально-экономического развития Российской Федерации на период до {exctract_year} года")]'}

    # Создаем директории
    # [os.makedirs(dir, exist_ok=True) for dir in
    #  [
    #      folder_load_forecast['Сценарные условия'],
    #      folder_load_forecast['Среднесрочный прогноз'],
    #      folder_load_forecast['Долгосрочный прогноз'],
    #      folder_screen_forecast['Сценарные условия'],
    #      folder_screen_forecast['Среднесрочный прогноз'],
    #      folder_screen_forecast['Долгосрочный прогноз'],
    #      folder_attachment,
    #      folder_logs,
    #  ]]

    #Выбор типа параметров
    if mode.lower() == 'prod':
        support_email = ''
        browser_path = r''
        driver_path = r''
        load_path = r''
        current_year = ''
        work_date = []
        work_month = []
        log_limit = 50
    elif mode.lower() == 'test':
        support_email = 'blackday52@mail.ru'
        driver_path = os.path.join(folder_root, 'Soft/chromedriver')
        browser_path = r'/opt/google/chrome/google-chrome'
        current_year = int('2021')
        work_date = [16, 17]
        work_month = [12, '']
        log_limit = 2