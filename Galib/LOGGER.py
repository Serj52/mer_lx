"""Импорты"""
import logging
import os
import zipfile
from datetime import datetime
from config import Config


class Log:
    """Класс по реализации логирования"""

    def init_logs(self):
        """Запись логфайлов в архив"""
        list_logs = [file for file in os.listdir(
            Config.folder_logs) if '.zip' not in file]
        # Если число логфайлов превышает лимит, то архивируем их
        if len(list_logs) > Config.log_limit:
            # создаем переменную - название и местоположение файла
            zname = os.path.join(
                Config.folder_logs,
                f'logs_archive {datetime.strftime(datetime.now(),"%d-%m-%Y %H-%M-%S")}.zip')
            newzip = zipfile.ZipFile(zname, 'w')  # создаем архив
            for file in list_logs:
                if '.zip' not in file:
                    # добавляем файл в архив
                    newzip.write(os.path.join(Config.folder_logs, file), file)
                    os.remove(os.path.join(Config.folder_logs, file))
            newzip.close()  # закрываем архив

    def set(self):
        """Установка настроек"""

        self.init_logs()
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            filename=os.path.join(
                Config.folder_logs,
                f'{Config.robot_name} {datetime.strftime(datetime.now(),"%d-%m-%Y %H-%M-%S")}.log')
        )
        return logging