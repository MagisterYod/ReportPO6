import os
import tkinter as tk
from tkinter import Tk, ttk, Button, messagebox, EW
from dotenv import load_dotenv, find_dotenv
from CreatePO6YearWindow import CreateWindowPO6Year
from CreatePO6MonthWindow import CreateWindowPO6Month
from CreatePO6DayWindow import CreateWindowPO6Day
from InsertParametersWindow import InsertParameters

load_dotenv(find_dotenv())
path = os.getenv('PATH_DIR')


class MainWindow(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        for c in range(3):
            self.columnconfigure(index=c, weight=1)
        for r in range(17):
            self.rowconfigure(index=r, weight=1)

        self.PO6_day_button = Button(root, text="ПО6 за смену", command=self.create_po6_day)
        self.PO6_month_button = Button(root, text="ПО6 накопительный", command=self.create_po6_month, width=42)
        self.PO6_year_button = Button(root, text="ПО6 годовой", command=self.create_po6_year)
        self.back_button = Button(root, text="Назад", command=self.back, width=20, fg='orange', background='white')
        self.insert_PO6_values_button = Button(root, text="Вставка значений", command=self.insert_parameters, width=42)
        self.create_PO6_button = Button(root, text="Выгрузка отчетов ПО6", command=self.create_po6)
        self.exit_button = Button(root, text="Завершить", command=self.exit, width=20, fg='red', background='white')
        self.main_menu()

    def main_menu(self):
        self.insert_PO6_values_button.grid(row=0, column=0, columnspan=2, sticky=EW)
        self.create_PO6_button.grid(row=1, column=0, columnspan=2, sticky=EW)
        self.exit_button.grid(row=2, column=0, columnspan=2, sticky=EW, pady=10)

    def insert_parameters(self):
        root_create = tk.Toplevel(self.master)
        frame_insert = InsertParameters(root_create)
        frame_insert.grid()  # Не забываем фрейм разместить в окне

    def po6_menu(self):
        self.PO6_day_button.grid(row=0, column=0, columnspan=4, sticky=EW)
        self.PO6_month_button.grid(row=1, column=0, columnspan=4, sticky=EW)
        self.PO6_year_button.grid(row=2, column=0, columnspan=4, sticky=EW)
        self.back_button.grid(row=4, column=0, sticky=EW, pady=10)
        self.exit_button.grid(row=4, column=3, sticky=EW, pady=10)

    def create_po6(self):
        self.insert_PO6_values_button.grid_forget()
        self.create_PO6_button.grid_forget()
        self.po6_menu()

    def create_po6_day(self):
        root_create = tk.Toplevel(self.master)
        frame_day = CreateWindowPO6Day(root_create)
        frame_day.grid()  # Не забываем фрейм разместить в окне

    def create_po6_month(self):
        root_create = tk.Toplevel(self.master)
        frame_month = CreateWindowPO6Month(root_create)
        frame_month.grid()  # Не забываем фрейм разместить в окне

    def create_po6_year(self):
        root_create = tk.Toplevel(self.master)
        frame_year = CreateWindowPO6Year(root_create)
        frame_year.grid()  # Не забываем фрейм разместить в окне

    def back(self):
        self.PO6_year_button.grid_forget()
        self.PO6_month_button.grid_forget()
        self.PO6_day_button.grid_forget()
        self.back_button.grid_forget()
        self.exit_button.grid_forget()
        self.main_menu()

    @staticmethod
    def exit():
        for filename in os.listdir(f'{path}'):
            if os.path.isfile(f'{path}\\{filename}'):
                try:
                    os.remove(f'{path}\\{filename}')
                except Exception as e:
                    messagebox.showerror('Ошибка', 'Закройте документ!')
                    print(e)
                    return
        root.destroy()


root = Tk()
root.title("Выгрузка из АСД")
frame = MainWindow(root)
frame.grid()  # не забываем фрейм размещать в окне
root.mainloop()
