import time
import datetime
import openpyxl
import os
import threading
from variables import p_time, path, report_path, insert_technics, sql_shovs_month, sql_veh_month, sql_bul_month, sql_pogr_month
from tkinter import ttk, Button, messagebox, W
from tkcalendar import Calendar


class CreateWindowPO6Month(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.date_entry = Calendar(
            self,
            selectmode='day',
            locale='ru_RU',
            year=int(datetime.datetime.now().strftime('%Y')),
            month=int(datetime.datetime.now().strftime('%m')),
            day=int(datetime.datetime.now().strftime('%d'))
        )
        self.start_button = Button(self, text='Выгрузить отчет', command=self.start, fg='green', width=42)
        self.text_progress_wb = ttk.Label(self, text='Тип техники: ', state='disable', width=42)
        self.progress_wb = ttk.Progressbar(self, orient='horizontal', mode='determinate', length=304, maximum=300)
        self.text_progress_ws = ttk.Label(self, text='Техника: ', state='disable', width=42)
        self.text_progress_row = ttk.Label(self, text='День: ', state='disable', width=42)
        self.po6_month_menu()

    def po6_month_menu(self):
        self.date_entry.grid(row=12, column=1, sticky=W)
        self.start_button.grid(row=13, column=0, columnspan=2, pady=5)

    def refresh(self):
        self.master.update()
        self.master.after(1000, self.refresh)

    def start(self):
        self.refresh()
        threading.Thread(target=self.take_doc).start()

    def timer_start(self):
        self.date_entry.grid_forget()
        self.start_button.grid_forget()
        self.progress_wb.grid(row=14, column=0, columnspan=3)
        self.text_progress_wb.grid(row=15, column=0, columnspan=3, sticky=W)
        self.text_progress_ws.grid(row=16, column=0, columnspan=3, sticky=W)
        self.text_progress_row.grid(row=17, column=0, columnspan=3, sticky=W)
        self.update_idletasks()

    def timer_stop(self):
        self.master.destroy()

    def changes(self, technics, tech, day, percent, sec=2):
        self.text_progress_wb['text'] = f'Тип техники:  {technics}'
        self.text_progress_ws['text'] = f'Техника: {tech}'
        if tech == '':
            self.text_progress_row['text'] = f'Идёт выгрузка данных из базы данных...'
        else:
            self.text_progress_row['text'] = f'День: {day} г.'
        self.progress_wb['value'] += percent
        self.update_idletasks()
        time.sleep(sec)

    def take_doc(self):

        try:
            self.timer_start()
            wb = openpyxl.load_workbook(f'{report_path}\PO6_days_new.xlsx')
            date = str(self.date_entry.get_date()).split('.')
            for num_date in range(1, int(date[0]) + 1):
                insert_date = f'{num_date}.{date[1]}.{date[2]}'
                percent = round(100 / (int(date[0]) * 14), 3) * 2.5
                ws = wb[str(num_date)]
                ws['A9'] = f'за {insert_date} г.'
                self.changes('Экскаваторы', '', 0, percent, 1)
                for row in sql_shovs_month(str(insert_date)).getvalue().fetchall():
                    if row[0] == '' or row[0] is None:
                        continue
                    if int(row[0]) == 17:
                        self.changes('Экскаваторы', 'Шагающий 6/45 №17', insert_date, percent)
                        insert_technics('shov', ws, 23, row)
                self.changes('Экскаваторы', '', 0, percent, 1)
                for row in sql_veh_month(str(insert_date)).getvalue().fetchall():
                    if int(row[0]) == 32:
                        self.changes('Самосвалы', 'Самосвал №32', insert_date, percent)
                        insert_technics('veh', ws,26, row)
                    if int(row[0]) == 35:
                        self.changes('Самосвалы', 'Самосвал №35', insert_date, percent)
                        insert_technics('veh', ws, 27, row)
                self.changes('Бульдозеры', '', 0, percent, 1)
                for row in sql_bul_month(str(insert_date)).getvalue().fetchall():
                    if row[0] == '' or row[0] is None:
                        continue
                    if int(row[0]) == 25:
                        self.changes('Бульдозеры', 'Трактор МТЗ-82 №25', insert_date, percent)
                        insert_technics('bul', ws, 43, row)
                    if int(row[0]) == 38:
                        self.changes('Бульдозеры', 'Будьдозер ДТ-75А №38', insert_date, percent)
                        insert_technics('bul', ws, 45, row)
                    if int(row[0]) == 28:
                        self.changes('Бульдозеры', 'Будьдозер CAT-D6R №28', insert_date, percent)
                        insert_technics('bul', ws, 47, row)
                    if int(row[0]) == 33:
                        self.changes('Бульдозеры', 'Будьдозер CAT-D9R №33', insert_date, percent)
                        insert_technics('bul', ws, 48, row)
                    if int(row[0]) == 4:
                        self.changes('Бульдозеры', 'Автогрейдер CAT-140М №4', insert_date, percent)
                        insert_technics('bul', ws, 53, row)
                    if int(row[0]) == 5:
                        self.changes('Бульдозеры', 'Автогрейдер Komatsu GD825А-2 №5', insert_date, percent)
                        insert_technics('bul', ws, 54, row)
                self.changes('Погрузчики', '', 0, 1)
                for row in sql_pogr_month(str(insert_date)).getvalue().fetchall():
                    if row[0] == '' or row[0] is None:
                        continue
                    if int(row[0]) == 15:
                        self.changes('Погрузчики', 'Погрузчик CAT-962 №15', insert_date, percent)
                        insert_technics('pogr', ws, 56, row)
                    if int(row[0]) == 16:
                        self.changes('Погрузчики', 'Погрузчик CAT-988H №16', insert_date, percent)
                        insert_technics('pogr', ws, 57, row)
                    if int(row[0]) == 26:
                        self.changes('Погрузчики', 'Погрузчик Shantui SL60W №26', insert_date, percent)
                        insert_technics('pogr', ws, 58, row)
                    if int(row[0]) == 30:
                        self.changes('Погрузчики', 'Погрузчик Shantui SL60W №30', insert_date, percent)
                        insert_technics('pogr', ws, 59, row)
                    if int(row[0]) == 25:
                        self.changes('Погрузчики', 'Погрузчик Volvo L120G №25', insert_date, percent)
                        insert_technics('pogr', ws, 60, row)

            wb.save(f'{path}\PO6_month_on_{date[0]}.{date[1]}.{date[2]}-{p_time}.xlsx')
            self.timer_stop()
            os.system(f'start excel.exe "{path}\PO6_month_on_{date[0]}.{date[1]}.{date[2]}-{p_time}.xlsx"')
        except Exception as e:
            messagebox.showerror('Ошибка', f'{e}')
            return
