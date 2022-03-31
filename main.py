from business import Business
from config import Config
import logging as lg
from datetime import datetime
import os
# import schedule
import time
from Galib.LOGGER import Log

class Robot:
    def __init__(self, config):
        self.config = config
        self.bus = Business(config)

    def robot_start(self):
        print('Робот запущен')
        lg.info('Робот запущен')
        # while True:
            #Если событие наступило проверяем почту
            # if self.check_shedule():
            #     print('Робот начинает работу')
            #     lg.info('Робот начинает работу')
                # Если пришло нужное сообщение начинаем работу с сайтом
        if self.bus.search_mail_exc():
            self.bus.data_process()
        # print('Перерыв 1 мин')
        # time.sleep(60)

    # def check_shedule(self):
    #     month_now = datetime.now().month
    #     date_now = datetime.now().day
    #     work_month = self.config.work_month
    #     work_date = self.config.work_date
    #     if date_now in work_date and month_now in work_month:
    #         return True
    #     else:
    #         return False


if __name__ == '__main__':
    config = Config()
    lg = Log().set()
    robot = Robot(config)
    robot.robot_start()

