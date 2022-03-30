import os
import traceback
import pyautogui
import win32com
import win32com.client as win32
from datetime import datetime, timedelta
from threading import Thread
import logging
from datetime import datetime
from outlook import Outlook


class ErrorHandler:
    def __init__(self,
                 name_def,
                 config,
                 tries_count=0,
                 minutes_wait=None,
                 attach_screen=False,
                 stop_robot=True,
                 except_execute=None,
                 finally_execute=None,
                 make_screen = True):

        self.try_num = 1
        self.config = config
        self.name_def = name_def
        self.robot_name = config.robot_name
        self.screen_shot_path = None
        self.tries_count = tries_count
        self.time_checker = None
        self.flag_to_stop_timer = False
        self.minutes_wait = minutes_wait
        self.attach_screen = attach_screen
        self.stop_robot = stop_robot
        self.except_execute = except_execute
        self.finally_execute = finally_execute
        self.make_screen = make_screen
        self.support_email = config.support_email


    def __enter__(self):

        # Запуск таймера
        if self.minutes_wait:
            self.time_checker = Thread(name='time_cheker', target=self.timer, args=(self.minutes_wait, ))
            self.flag_to_stop_timer = False
            self.time_checker.start()
            return self.time_checker

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Останавливаем таймер
        if self.time_checker:
            self.flag_to_stop_timer = True
            self.time_checker.join()
            del self.time_checker
            # Обрабатываем ошибки
        if exc_val:
            if self.tries_count == 0:
                self.simple_error()
            else:
                self.restart_block_on_error()
        return True

    def simple_error(self):
        """
        При появлении ошибки отправляет письмо в поддержку
        Применение:
        with ErrorHndler(name_def, config):
            your_finction()
        Применение с таймером:
        with ErrorHndler(name_def, config, minutes_wait=5):
            your_finction()
        Применение без остановки робота при ошибке:
        with ErrorHndler(name_def, config, stop_robot=False):
            your_finction()
        """
        trace = traceback.format_exc()
        logging.error(f'\n\n{trace}')
        if self.stop_robot or self.tries_count != 0:
            logging.info('Робот остановлен')
            print(f'{datetime.now()} Робот остановлен')
            self.finally_execution()
            exit(-1)
            Outlook.send_mail(to=self.support_email,
                              body=f'Робот {self.robot_name} остановлен. Ошибка в блоке {self.name_def}',
                              subject='ERROR')
        else:
            pass
        self.finally_execution()

    def timer(self, minutes_wait):
        """ПРоверяет текущее время, если ожидаемое время истекло, направляет письмо в поддержку
        После отправки робот продолжает работу
        minutes_wait - ожидаемое время выполнения блока выделенного блока
        eh = ErrorHndler(name_def, config, mitutes_wait=5)
        while True:
            with eh:
                your_funcrtion()
                break
        """
        import pythoncom
        pythoncom.CoInitialize()
        start = datetime.now()
        while True:
            if not self.flag_to_stop_timer:
                # Проверяем текущее время, если больше minutes_wait, направляем птсьмл в поддержку
                if datetime.now() - timedelta(minutes=minutes_wait) >= start:
                    outlook = None
                    try:
                        outlook = win32com.client.Dispatch('Outlook.Application')
                        mail = outlook.CreateItem(0)
                        mail.To = self.config.support_email
                        mail.Subject = f'TIME OUT in {self.robot_name}'
                        mail.Body = f'Ожидаемое время выполнения задачи: {self.name_def}'
                        mail.Send()
                        self.flag_to_stop_timer = True
                    except Exception:
                        logging.critical(f'Письмо не отправлено {traceback.format_exc()}')

                    finally:
                        del outlook
            else:
                break

    def restart_block_on_error(self):
        """
        Выполняет перезапуск блока, создает скриншот на момент перезапуска, отправляет сообщение в поддержку
        При успешном выполнении блока - выход из цикла (break). В случае ошибок - повтор (tries_count)
        tries_count: количество попыток до остановки. Если tries_count=0 - вызов метода simple_error
        Применение:
        eh = ErrorHndler(name_def, config, tries_count=3)
        while True:
            with eh:
                your_function()
                break
        """
        logging.error('\n\n' + traceback.format_exc())
        logging.error(f'Ошибка в блоке: "{self.name_def}". Перезапуск блока')
        print(f'{datetime.now()} Ошибка в блоке: "{self.name_def}". Перезапуск блока')
        self.try_num += 1
        if self.try_num > self.tries_count:
            logging.error('Кол-во попыток изчерпано. Отправляем письмо в поддержку')
            print(f'{datetime.now()} Кол-во попыток изчерпано. Отправляем письмо в поддержку')
            self.simple_error()
        logging.info(f'Попытка {str(self.try_num)} из {str(self.tries_count)}')
        print(f'{datetime.now()} Попытка {str(self.try_num)} из {str(self.tries_count)}')


    def except_execution(self):
        """Вызов переданного метода при возникновении ошибки"""
        result_execution =''
        if self.except_execute:
            if callable(self.except_execute):
                try:
                    result_execution = self.except_execute()
                except Exception as err:
                    logging.error(f'Ошибка при выполнении переданного метода {err}')
                if result_execution not in ['', None]:
                    result_execution = f'\nExecution info {str(result_execution)}\n'
                else:
                    result_execution = ''
            else:
                result_execution = f'\nExecution info {str(result_execution)}\n'
        return result_execution

    def finally_execution(self):
        if self.finally_execute:
            if callable(self.finally_execute):
                try:
                    self.finally_execute()
                except Exception as err:
                    logging.error(f'Ошибка при выполнении переданного метода {err}')