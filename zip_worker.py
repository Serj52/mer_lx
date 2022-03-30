"""Импорты"""
import os
import logging as lg
import re
from zipfile import ZipFile
import shutil
# from outlook import Outlook

class Zipworker:
    """Класс по работе со скаченным архивом"""
    def __init__(self, config):
        self.config = config

    def remove_zip(self, name_zip, match):
        """Удаляет скаченный архив, если все файлы найдены"""
        if len(match) == 0:
            os.remove(name_zip)
        else:
            #Если какой-то файл не найден возвращаем его имя. Останавливем робота.
            #Для возобновления работы нужно удалить скаченный архив
            lg.info(f'Не найден файл {list(match.keys())}. Отправляю письмо. Робот остановлен')
            print(f'Не найден файл {list(match.keys())}. Отправляю письмо. Робот остановлен')
            # Outlook.send_mail(to=self.config.support_email,
            #                   body=f'{list(match.keys())}')
            exit(-1)

    @staticmethod
    def return_zipname(exctract_dir):
        """Возвращает имя скаченного архива """
        for name in os.listdir(exctract_dir):
            if '.zip' in name:
                name_zip = os.path.join(exctract_dir, name)
                return name_zip

    def unpack_zipfile(self, exctract_dir, zip_templates):
        """Распаковываем архив"""
        name_zip = self.return_zipname(exctract_dir)
        lg.info("Начинаю работу с архивом")
        print("Начинаю работу с архивом")
        with ZipFile(name_zip) as archive:
            # Перебирем файлы в архиве
            for entry in archive.infolist():
                name = entry.filename.encode('cp437').decode('cp866')
                #Перебираем шаблон имен файлов
                for name_file, template in zip_templates.items():
                    # Если имя совпадает с шаблоном, то извлекаем файл
                    if re.fullmatch(template, name):
                        target = os.path.join(exctract_dir, name_file)
                        with archive.open(entry) as source, open(target, 'wb') as dest:
                            shutil.copyfileobj(source, dest)
                            #Удаляем из шаблона найденный файл
                            zip_templates.pop(name_file)
                            break
        #Удаляем архив после извлечения файлов
        self.remove_zip(name_zip, zip_templates)
        lg.info(f"Файлы извлечены")
        print(f"Файлы извлечены")