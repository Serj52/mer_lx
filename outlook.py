# import win32com.client as win32
# from win32com.client import Dispatch
# import logging as lg
# import os
# import time
# import win32ui
# import win32gui
# import win32con
#
#
# class Outlook:
#     """Стандартные методы работы с Outlook"""
#
#     @staticmethod
#     def send_mail(to, body, subject='Програмный робот', attachment=None, CC=None, BCC=None, recovery=False):
#         """Отправка писем"""
#         lg.info('Отправляю почту')
#         print('Отправляю почту')
#         time_sleep = 30
#         def worker():
#             nonlocal recovery
#             nonlocal time_sleep
#             try:
#                 outlook = Dispatch("Outlook.application")
#                 mail = outlook.CreateItem(0)
#                 mail.To = to
#                 if CC:
#                     mail.CC = CC
#                 if BCC:
#                     mail.CC = BCC
#                 mail.Subject = subject
#                 mail.HTMLBody = body
#                 if attachment:
#                     for att in attachment:
#                         if att:
#                             mail.Attachments.Add(att)
#                 mail.Send()
#                 recovery = False
#             except Exception as er:
#                 lg.info('Неудалось отправить письмо')
#                 print('Неудалось отправить письмо')
#                 if recovery:
#                     lg.info(f'Отправляю повторно через {time_sleep} сек')
#                     print(f'Отправляю повторно через {time_sleep} сек')
#                     time.sleep(time_sleep)
#         worker()
#         while recovery:
#             worker()
#         lg.info(f'Отправка писем завершена')
#         print(f'Отправка писем завершена')
#
#     @staticmethod
#     def get_mail_data_by_subjects(config):
#         """Поиск нового письма по теме"""
#         lg.info("Проверяю почту")
#         print("Проверяю почту")
#         mails = []
#         themes = config.theme_mail
#         outlook = Dispatch("Outlook.application")
#         namespace = outlook.GetNamespace('MAPI')
#         inbox = namespace.GetDefaultFolder(6).Items.Restrict("[Unread] = True")
#         #Если есть новые сообщений
#         if inbox.COUNT:
#             for item in inbox:
#                 mail_theme = item.Subject.strip()
#                 if themes.lower() in mail_theme.lower():
#                     try:
#                         lg.info(f"Получено письмо {mail_theme}")
#                         print(f"Получено письмо {mail_theme}")
#                         item.UnRead = False
#                         for attachment in item.Attachments:
#                             if 'inputfile' in str(attachment):
#                                 path = os.path.join(config.folder_attachment, str(attachment))
#                                 attachment.SaveAsFile(path)
#                                 lg.info(f"Вложение {str(attachment)} сохранено")
#                                 print(f"Вложение {str(attachment)} сохранено")
#                                 mails.append(str(attachment))
#                                 return mails
#                         #Если нет нужного вложения в письме вызываем исключение
#                         raise Exception
#                     except Exception as er:
#                         message = f'Не удалось считать данные из письма {er}'
#                         body = f'{message} Проверьте работу'
#                         Outlook.send_mail(config.support_email, body)
#                 else:
#                     try:
#                         item.UnRead = False
#                     except Exception as er:
#                         message = f'Робот остановлен. В outlook появился неопознанный объект {er}'
#                         lg.info(f'Робот остановлен. В outlook появился неопознанный объект {er}')
#                         print(f'Робот остановлен. В outlook появился неопознанный объект {er}')
#                         body = f'{message} Проверьте работу'
#                         Outlook.send_mail(config.support_email, body)
#                         lg.info(f'Робот остановлен')
#                         print(f'Робот остановлен')
#                         exit(-1)
#             lg.info(f'Нет нужного письма')
#             print(f'Нет нужного письма')
#             return mails
#         else:
#             lg.info(f'Нет новых писем')
#             print(f'Нет новых писем')
#             return mails
#
#     @staticmethod
#     def run():
#         """Проверяет запущен ли отлук в системе"""
#
#         lg.info('Проверяю запущен ли Outlook')
#         print(f'Проверяю запущен ли Outlook')
#         try:
#             win32ui.FindWindow(None, "Microsoft Outlook")
#             lg.info('Outlook запущен')
#             print(f'Outlook запущен')
#         except win32ui.error:
#             lg.info('Outlook не запущен. Запускаю')
#             print(f'Outlook не запущен. Запускаю')
#             os.startfile('outlook')
#             time.sleep(30)
#             def enumHandler(hwnd, lParam):
#                 res = str(win32gui.GetWindowText(hwnd)).lower()
#                 if ' - outlook' in str(win32gui.GetWindowText(hwnd)).lower():
#                     win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
#                     win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
#             #перечисляет все окна верхнего уровня передавая дескриптор каждого окна, по очереди,
#             # в определяемую приложением функцию обратного вызова.
#             win32gui.EnumWindows(enumHandler, None)
#             lg.info('Outlook запущен')
#             print(f'Outlook запущен')