import threading
import tkinter
from tkinter import ttk, Label, Entry, Button, Spinbox, W, EW, E, NW, messagebox
from typing import Callable, Any

from variables import p_date, techs, months, colors, sql_get_po6, sql_insert_po6


class InsertParameters(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        def init_separator_horizontal():
            return ttk.Separator(self, orient='horizontal', takefocus=1, cursor='plus')

        def init_label_bgd(tech, bgr):
            return Label(self, text=tech, background=bgr)

        def init_label(tech):
            return Label(self, text=tech)

        def init_entry():
            return Entry(self, width=10, justify='center')

        def init_spin(fr, to, bgr, width=3):
            return Spinbox(self, from_=fr, to=to, justify='center', wrap=True, width=width, background=bgr)

        self.months_box = ttk.Combobox(self, values=list(months.keys()))
        self.months_box.set(list(months.keys())[0])
        self.months_box_label = Label(self)
        self.year_label = init_label(p_date[2])
        self.get_button = Button(self, text='Выгрузить', fg='green', command=self.start, width=30)
        self.volume = init_label('Объем')
        self.top = init_label_bgd('ТО', colors.get('top'))
        self.trp = init_label_bgd('ТР', colors.get('trp'))
        self.krp = init_label_bgd('КР', colors.get('krp'))
        self.rgp = init_label_bgd('РП', colors.get('rgp'))
        self.oip = init_label_bgd('Обед', colors.get('oip'))
        self.ir = init_label_bgd('ИP', colors.get('ir'))
        self.bvr = init_label_bgd('БВР', colors.get('bvr'))
        self.sep_0 = init_separator_horizontal()
        self.esh_label = init_label(techs.get(1))
        self.esh_entry_volume = init_entry()
        self.esh_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.esh_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.esh_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.esh_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.esh_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.esh_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.esh_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.esh_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.esh_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.esh_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.esh_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.esh_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.esh_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.esh_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_1 = init_separator_horizontal()
        self.bel32_label = init_label(techs.get(2))
        self.bel32_entry_volume = init_entry()
        self.bel32_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.bel32_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.bel32_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.bel32_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.bel32_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.bel32_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.bel32_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.bel32_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.bel32_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.bel32_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.bel32_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.bel32_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.bel32_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.bel32_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.bel35_label = init_label(techs.get(3))
        self.bel35_entry_volume = init_entry()
        self.bel35_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.bel35_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.bel35_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.bel35_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.bel35_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.bel35_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.bel35_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.bel35_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.bel35_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.bel35_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.bel35_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.bel35_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.bel35_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.bel35_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_2 = init_separator_horizontal()
        self.mtz_label = init_label(techs.get(4))
        self.mtz_entry_volume = init_entry()
        self.mtz_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.mtz_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.mtz_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.mtz_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.mtz_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.mtz_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.mtz_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.mtz_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.mtz_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.mtz_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.mtz_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.mtz_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.mtz_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.mtz_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.dt_label = init_label(techs.get(5))
        self.dt_entry_volume = init_entry()
        self.dt_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.dt_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.dt_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.dt_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.dt_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.dt_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.dt_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.dt_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.dt_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.dt_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.dt_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.dt_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.dt_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.dt_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_3 = init_separator_horizontal()
        self.cat28_label = init_label(techs.get(6))
        self.cat28_entry_volume = init_entry()
        self.cat28_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.cat28_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.cat28_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.cat28_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.cat28_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.cat28_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.cat28_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.cat28_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.cat28_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.cat28_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.cat28_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.cat28_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.cat28_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.cat28_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.cat33_label = init_label(techs.get(7))
        self.cat33_entry_volume = init_entry()
        self.cat33_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.cat33_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.cat33_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.cat33_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.cat33_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.cat33_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.cat33_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.cat33_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.cat33_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.cat33_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.cat33_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.cat33_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.cat33_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.cat33_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_4 = init_separator_horizontal()
        self.cat4_label = init_label(techs.get(8))
        self.cat4_entry_volume = init_entry()
        self.cat4_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.cat4_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.cat4_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.cat4_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.cat4_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.cat4_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.cat4_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.cat4_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.cat4_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.cat4_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.cat4_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.cat4_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.cat4_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.cat4_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.kom5_label = init_label(techs.get(9))
        self.kom5_entry_volume = init_entry()
        self.kom5_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.kom5_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.kom5_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.kom5_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.kom5_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.kom5_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.kom5_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.kom5_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.kom5_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.kom5_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.kom5_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.kom5_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.kom5_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.kom5_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_5 = init_separator_horizontal()
        self.cat15_label = init_label(techs.get(10))
        self.cat15_entry_volume = init_entry()
        self.cat15_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.cat15_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.cat15_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.cat15_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.cat15_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.cat15_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.cat15_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.cat15_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.cat15_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.cat15_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.cat15_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.cat15_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.cat15_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.cat15_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.cat16_label = init_label(techs.get(11))
        self.cat16_entry_volume = init_entry()
        self.cat16_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.cat16_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.cat16_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.cat16_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.cat16_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.cat16_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.cat16_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.cat16_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.cat16_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.cat16_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.cat16_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.cat16_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.cat16_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.cat16_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.vol25_label = init_label(techs.get(12))
        self.vol25_entry_volume = init_entry()
        self.vol25_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.vol25_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.vol25_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.vol25_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.vol25_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.vol25_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.vol25_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.vol25_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.vol25_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.vol25_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.vol25_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.vol25_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.vol25_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.vol25_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sha26_label = init_label(techs.get(13))
        self.sha26_entry_volume = init_entry()
        self.sha26_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.sha26_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.sha26_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.sha26_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.sha26_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.sha26_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.sha26_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.sha26_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.sha26_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.sha26_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.sha26_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.sha26_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.sha26_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.sha26_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sha30_label = init_label(techs.get(14))
        self.sha30_entry_volume = init_entry()
        self.sha30_entry_top_hours = init_spin(0, 744, colors.get('top'), 5)
        self.sha30_entry_top_minutes = init_spin(0, 59, colors.get('top'))
        self.sha30_entry_trp_hours = init_spin(0, 744, colors.get('trp'), 5)
        self.sha30_entry_trp_minutes = init_spin(0, 59, colors.get('trp'))
        self.sha30_entry_krp_hours = init_spin(0, 744, colors.get('krp'), 5)
        self.sha30_entry_krp_minutes = init_spin(0, 59, colors.get('krp'))
        self.sha30_entry_rgp_hours = init_spin(0, 744, colors.get('rgp'), 5)
        self.sha30_entry_rgp_minutes = init_spin(0, 59, colors.get('rgp'))
        self.sha30_entry_oip_hours = init_spin(0, 744, colors.get('oip'), 5)
        self.sha30_entry_oip_minutes = init_spin(0, 59, colors.get('oip'))
        self.sha30_entry_ir_hours = init_spin(0, 744, colors.get('ir'), 5)
        self.sha30_entry_ir_minutes = init_spin(0, 59, colors.get('ir'))
        self.sha30_entry_bvr_hours = init_spin(0, 744, colors.get('bvr'), 5)
        self.sha30_entry_bvr_minutes = init_spin(0, 59, colors.get('bvr'))
        self.sep_6 = init_separator_horizontal()
        self.accept_button = Button(self, text='Внести данные', fg='green', command=self.insert_data)
        self.start_menu()

    def refresh(self):
        self.master.update()
        self.master.after(1000, self.refresh)

    def start(self):
        self.refresh()
        threading.Thread(target=self.take_doc).start()

    def start_menu(self):
        self.months_box.grid(row=0, column=0, sticky=W)
        self.year_label.grid(row=0, column=1, sticky=W)
        self.get_button.grid(row=1, columnspan=2, sticky=W)

    def take_doc(self):

        def insert_in_entry(item, element1, element2):
            print(item)
            if type(item) is int:
                var = tkinter.StringVar(self)
                var.set(str(item))
                element1['textvariable'] = var
            if type(item) is float:
                item_h = str(item).split('.')[0]
                item_m = lambda x: f'{str(item).split(".")[1]}0'\
                    if len(str(item).split('.')[1]) == 1 else f'{str(item).split(".")[1]}'
                # if len(str(item).split('.')[1]) == 1:
                #     item_m = f'{str(item).split(".")[1]}0'
                # else:
                #     item_m = f'{str(item).split(".")[1]}'
                var_h = tkinter.StringVar(self)
                var_m = tkinter.StringVar(self)
                var_h.set(item_h)
                var_m.set(str(int(int(item_m(item)) / 1.7)))
                element1['textvariable'] = var_h
                element2['textvariable'] = var_m

        for row in sql_get_po6(f'01.{months.get(self.months_box.get())}.{p_date[2]}').getvalue().fetchall():
            if row[2] == 1:
                insert_in_entry(row[3], self.mtz_entry_volume, self.mtz_entry_volume)
                insert_in_entry(row[4], self.mtz_entry_top_hours, self.mtz_entry_top_minutes)
                insert_in_entry(row[5], self.mtz_entry_trp_hours, self.mtz_entry_trp_minutes)
                insert_in_entry(row[6], self.mtz_entry_krp_hours, self.mtz_entry_krp_minutes)
                insert_in_entry(row[7], self.mtz_entry_rgp_hours, self.mtz_entry_rgp_minutes)
                insert_in_entry(row[8], self.mtz_entry_oip_hours, self.mtz_entry_oip_minutes)
                insert_in_entry(row[9], self.mtz_entry_ir_hours, self.mtz_entry_ir_minutes)
                insert_in_entry(row[9], self.mtz_entry_bvr_hours, self.mtz_entry_bvr_minutes)
            if row[2] == 4:
                insert_in_entry(row[3], self.cat4_entry_volume, self.cat4_entry_volume)
                insert_in_entry(row[4], self.cat4_entry_top_hours, self.cat4_entry_top_minutes)
                insert_in_entry(row[5], self.cat4_entry_trp_hours, self.cat4_entry_trp_minutes)
                insert_in_entry(row[6], self.cat4_entry_krp_hours, self.cat4_entry_krp_minutes)
                insert_in_entry(row[7], self.cat4_entry_rgp_hours, self.cat4_entry_rgp_minutes)
                insert_in_entry(row[8], self.cat4_entry_oip_hours, self.cat4_entry_oip_minutes)
                insert_in_entry(row[9], self.cat4_entry_ir_hours, self.cat4_entry_ir_minutes)
                insert_in_entry(row[9], self.cat4_entry_bvr_hours, self.cat4_entry_bvr_minutes)
            if row[2] == 9:
                insert_in_entry(row[3], self.cat28_entry_volume, self.cat28_entry_volume)
                insert_in_entry(row[4], self.cat28_entry_top_hours, self.cat28_entry_top_minutes)
                insert_in_entry(row[5], self.cat28_entry_trp_hours, self.cat28_entry_trp_minutes)
                insert_in_entry(row[6], self.cat28_entry_krp_hours, self.cat28_entry_krp_minutes)
                insert_in_entry(row[7], self.cat28_entry_rgp_hours, self.cat28_entry_rgp_minutes)
                insert_in_entry(row[8], self.cat28_entry_oip_hours, self.cat28_entry_oip_minutes)
                insert_in_entry(row[9], self.cat28_entry_ir_hours, self.cat28_entry_ir_minutes)
                insert_in_entry(row[9], self.cat28_entry_bvr_hours, self.cat28_entry_bvr_minutes)
            if row[2] == 10:
                insert_in_entry(row[3], self.cat33_entry_volume, self.cat33_entry_volume)
                insert_in_entry(row[4], self.cat33_entry_top_hours, self.cat33_entry_top_minutes)
                insert_in_entry(row[5], self.cat33_entry_trp_hours, self.cat33_entry_trp_minutes)
                insert_in_entry(row[6], self.cat33_entry_krp_hours, self.cat33_entry_krp_minutes)
                insert_in_entry(row[7], self.cat33_entry_rgp_hours, self.cat33_entry_rgp_minutes)
                insert_in_entry(row[8], self.cat33_entry_oip_hours, self.cat33_entry_oip_minutes)
                insert_in_entry(row[9], self.cat33_entry_ir_hours, self.cat33_entry_ir_minutes)
                insert_in_entry(row[9], self.cat33_entry_bvr_hours, self.cat33_entry_bvr_minutes)
            if row[2] == 15:
                insert_in_entry(row[3], self.kom5_entry_volume, self.kom5_entry_volume)
                insert_in_entry(row[4], self.kom5_entry_top_hours, self.kom5_entry_top_minutes)
                insert_in_entry(row[5], self.kom5_entry_trp_hours, self.kom5_entry_trp_minutes)
                insert_in_entry(row[6], self.kom5_entry_krp_hours, self.kom5_entry_krp_minutes)
                insert_in_entry(row[7], self.kom5_entry_rgp_hours, self.kom5_entry_rgp_minutes)
                insert_in_entry(row[8], self.kom5_entry_oip_hours, self.kom5_entry_oip_minutes)
                insert_in_entry(row[9], self.kom5_entry_ir_hours, self.kom5_entry_ir_minutes)
                insert_in_entry(row[9], self.kom5_entry_bvr_hours, self.kom5_entry_bvr_minutes)
            if row[2] == 28:
                insert_in_entry(row[3], self.bel32_entry_volume, self.bel32_entry_volume)
                insert_in_entry(row[4], self.bel32_entry_top_hours, self.bel32_entry_top_minutes)
                insert_in_entry(row[5], self.bel32_entry_trp_hours, self.bel32_entry_trp_minutes)
                insert_in_entry(row[6], self.bel32_entry_krp_hours, self.bel32_entry_krp_minutes)
                insert_in_entry(row[7], self.bel32_entry_rgp_hours, self.bel32_entry_rgp_minutes)
                insert_in_entry(row[8], self.bel32_entry_oip_hours, self.bel32_entry_oip_minutes)
                insert_in_entry(row[9], self.bel32_entry_ir_hours, self.bel32_entry_ir_minutes)
                insert_in_entry(row[9], self.bel32_entry_bvr_hours, self.bel32_entry_bvr_minutes)
            if row[2] == 31:
                insert_in_entry(row[3], self.bel35_entry_volume, self.bel35_entry_volume)
                insert_in_entry(row[4], self.bel35_entry_top_hours, self.bel35_entry_top_minutes)
                insert_in_entry(row[5], self.bel35_entry_trp_hours, self.bel35_entry_trp_minutes)
                insert_in_entry(row[6], self.bel35_entry_krp_hours, self.bel35_entry_krp_minutes)
                insert_in_entry(row[7], self.bel35_entry_rgp_hours, self.bel35_entry_rgp_minutes)
                insert_in_entry(row[8], self.bel35_entry_oip_hours, self.bel35_entry_oip_minutes)
                insert_in_entry(row[9], self.bel35_entry_ir_hours, self.bel35_entry_ir_minutes)
                insert_in_entry(row[9], self.bel35_entry_bvr_hours, self.bel35_entry_bvr_minutes)
            if row[2] == 43:
                insert_in_entry(row[3], self.cat16_entry_volume, self.cat16_entry_volume)
                insert_in_entry(row[4], self.cat16_entry_top_hours, self.cat16_entry_top_minutes)
                insert_in_entry(row[5], self.cat16_entry_trp_hours, self.cat16_entry_trp_minutes)
                insert_in_entry(row[6], self.cat16_entry_krp_hours, self.cat16_entry_krp_minutes)
                insert_in_entry(row[7], self.cat16_entry_rgp_hours, self.cat16_entry_rgp_minutes)
                insert_in_entry(row[8], self.cat16_entry_oip_hours, self.cat16_entry_oip_minutes)
                insert_in_entry(row[9], self.cat16_entry_ir_hours, self.cat16_entry_ir_minutes)
                insert_in_entry(row[9], self.cat16_entry_bvr_hours, self.cat16_entry_bvr_minutes)
            if row[2] == 45:
                insert_in_entry(row[3], self.esh_entry_volume, self.esh_entry_volume)
                insert_in_entry(row[4], self.esh_entry_top_hours, self.esh_entry_top_minutes)
                insert_in_entry(row[5], self.esh_entry_trp_hours, self.esh_entry_trp_minutes)
                insert_in_entry(row[6], self.esh_entry_krp_hours, self.esh_entry_krp_minutes)
                insert_in_entry(row[7], self.esh_entry_rgp_hours, self.esh_entry_rgp_minutes)
                insert_in_entry(row[8], self.esh_entry_oip_hours, self.esh_entry_oip_minutes)
                insert_in_entry(row[9], self.esh_entry_ir_hours, self.esh_entry_ir_minutes)
                insert_in_entry(row[9], self.esh_entry_bvr_hours, self.esh_entry_bvr_minutes)
            if row[2] == 58:
                insert_in_entry(row[3], self.cat15_entry_volume, self.cat15_entry_volume)
                insert_in_entry(row[4], self.cat15_entry_top_hours, self.cat15_entry_top_minutes)
                insert_in_entry(row[5], self.cat15_entry_trp_hours, self.cat15_entry_trp_minutes)
                insert_in_entry(row[6], self.cat15_entry_krp_hours, self.cat15_entry_krp_minutes)
                insert_in_entry(row[7], self.cat15_entry_rgp_hours, self.cat15_entry_rgp_minutes)
                insert_in_entry(row[8], self.cat15_entry_oip_hours, self.cat15_entry_oip_minutes)
                insert_in_entry(row[9], self.cat15_entry_ir_hours, self.cat15_entry_ir_minutes)
                insert_in_entry(row[9], self.cat15_entry_bvr_hours, self.cat15_entry_bvr_minutes)
            if row[2] == 272:
                insert_in_entry(row[3], self.dt_entry_volume, self.dt_entry_volume)
                insert_in_entry(row[4], self.dt_entry_top_hours, self.dt_entry_top_minutes)
                insert_in_entry(row[5], self.dt_entry_trp_hours, self.dt_entry_trp_minutes)
                insert_in_entry(row[6], self.dt_entry_krp_hours, self.dt_entry_krp_minutes)
                insert_in_entry(row[7], self.dt_entry_rgp_hours, self.dt_entry_rgp_minutes)
                insert_in_entry(row[8], self.dt_entry_oip_hours, self.dt_entry_oip_minutes)
                insert_in_entry(row[9], self.dt_entry_ir_hours, self.dt_entry_ir_minutes)
                insert_in_entry(row[9], self.dt_entry_bvr_hours, self.dt_entry_bvr_minutes)
            if row[2] == 280:
                insert_in_entry(row[3], self.vol25_entry_volume, self.vol25_entry_volume)
                insert_in_entry(row[4], self.vol25_entry_top_hours, self.vol25_entry_top_minutes)
                insert_in_entry(row[5], self.vol25_entry_trp_hours, self.vol25_entry_trp_minutes)
                insert_in_entry(row[6], self.vol25_entry_krp_hours, self.vol25_entry_krp_minutes)
                insert_in_entry(row[7], self.vol25_entry_rgp_hours, self.vol25_entry_rgp_minutes)
                insert_in_entry(row[8], self.vol25_entry_oip_hours, self.vol25_entry_oip_minutes)
                insert_in_entry(row[9], self.vol25_entry_ir_hours, self.vol25_entry_ir_minutes)
                insert_in_entry(row[9], self.vol25_entry_bvr_hours, self.vol25_entry_bvr_minutes)
            if row[2] == 283:
                insert_in_entry(row[3], self.sha26_entry_volume, self.sha26_entry_volume)
                insert_in_entry(row[4], self.sha26_entry_top_hours, self.sha26_entry_top_minutes)
                insert_in_entry(row[5], self.sha26_entry_trp_hours, self.sha26_entry_trp_minutes)
                insert_in_entry(row[6], self.sha26_entry_krp_hours, self.sha26_entry_krp_minutes)
                insert_in_entry(row[7], self.sha26_entry_rgp_hours, self.sha26_entry_rgp_minutes)
                insert_in_entry(row[8], self.sha26_entry_oip_hours, self.sha26_entry_oip_minutes)
                insert_in_entry(row[9], self.sha26_entry_ir_hours, self.sha26_entry_ir_minutes)
                insert_in_entry(row[9], self.sha26_entry_bvr_hours, self.sha26_entry_bvr_minutes)
            if row[2] == 284:
                insert_in_entry(row[3], self.sha30_entry_volume, self.sha30_entry_volume)
                insert_in_entry(row[4], self.sha30_entry_top_hours, self.sha30_entry_top_minutes)
                insert_in_entry(row[5], self.sha30_entry_trp_hours, self.sha30_entry_trp_minutes)
                insert_in_entry(row[6], self.sha30_entry_krp_hours, self.sha30_entry_krp_minutes)
                insert_in_entry(row[7], self.sha30_entry_rgp_hours, self.sha30_entry_rgp_minutes)
                insert_in_entry(row[8], self.sha30_entry_oip_hours, self.sha30_entry_oip_minutes)
                insert_in_entry(row[9], self.sha30_entry_ir_hours, self.sha30_entry_ir_minutes)
                insert_in_entry(row[9], self.sha30_entry_bvr_hours, self.sha30_entry_bvr_minutes)

        def grid_sep_horizontal(element, row):
            return element.grid(row=row, column=0, columnspan=22, sticky=EW, padx=2, pady=2)

        def grid_top_label(element, column):
            return element.grid(row=0, column=column, columnspan=2, padx=2, pady=2, sticky=EW)

        def grid_row_label(element, row):
            return element.grid(row=row, column=0, sticky=NW, columnspan=2, padx=2, pady=2)

        def grid_volume(element, row):
            return element.grid(row=row, column=2, sticky=W, padx=2, pady=2)

        def grid_row_spin(element, row, column):
            return element.grid(row=row, column=column, pady=2)

        self.months_box.grid_forget()
        self.get_button.grid_forget()
        self.months_box_label.grid(row=0, column=0, sticky=E)
        self.months_box_label['text'] = self.months_box.get()
        self.months_box_label['font'] = ('Arial', 12, 'bold')
        self.year_label['font'] = ('Arial', 12, 'bold')
        self.volume.grid(row=0, column=2, padx=2, pady=2)
        grid_top_label(self.top, 3)
        grid_top_label(self.trp, 5)
        grid_top_label(self.krp, 7)
        grid_top_label(self.rgp, 9)
        grid_top_label(self.oip, 11)
        grid_top_label(self.ir, 13)
        grid_top_label(self.bvr, 15)

        grid_sep_horizontal(self.sep_0, 1)

        grid_row_label(self.esh_label, 2)
        grid_volume(self.esh_entry_volume, 2)
        grid_row_spin(self.esh_entry_top_hours, 2, 3)
        grid_row_spin(self.esh_entry_top_minutes, 2, 4)
        grid_row_spin(self.esh_entry_trp_hours, 2, 5)
        grid_row_spin(self.esh_entry_trp_minutes, 2, 6)
        grid_row_spin(self.esh_entry_krp_hours, 2, 7)
        grid_row_spin(self.esh_entry_krp_minutes, 2, 8)
        grid_row_spin(self.esh_entry_rgp_hours, 2, 9)
        grid_row_spin(self.esh_entry_rgp_minutes, 2, 10)
        grid_row_spin(self.esh_entry_oip_hours, 2, 11)
        grid_row_spin(self.esh_entry_oip_minutes, 2, 12)
        grid_row_spin(self.esh_entry_ir_hours, 2, 13)
        grid_row_spin(self.esh_entry_ir_minutes, 2, 14)
        grid_row_spin(self.esh_entry_bvr_hours, 2, 15)
        grid_row_spin(self.esh_entry_bvr_minutes, 2, 16)

        grid_sep_horizontal(self.sep_1, 3)

        grid_row_label(self.bel32_label, 4)
        grid_volume(self.bel32_entry_volume, 4)
        grid_row_spin(self.bel32_entry_top_hours, 4, 3)
        grid_row_spin(self.bel32_entry_top_minutes, 4, 4)
        grid_row_spin(self.bel32_entry_trp_hours, 4, 5)
        grid_row_spin(self.bel32_entry_trp_minutes, 4, 6)
        grid_row_spin(self.bel32_entry_krp_hours, 4, 7)
        grid_row_spin(self.bel32_entry_krp_minutes, 4, 8)
        grid_row_spin(self.bel32_entry_rgp_hours, 4, 9)
        grid_row_spin(self.bel32_entry_rgp_minutes, 4, 10)
        grid_row_spin(self.bel32_entry_oip_hours, 4, 11)
        grid_row_spin(self.bel32_entry_oip_minutes, 4, 12)
        grid_row_spin(self.bel32_entry_ir_hours, 4, 13)
        grid_row_spin(self.bel32_entry_ir_minutes, 4, 14)
        grid_row_spin(self.bel32_entry_bvr_hours, 4, 15)
        grid_row_spin(self.bel32_entry_bvr_minutes, 4, 16)
        grid_row_label(self.bel35_label, 5)
        grid_volume(self.bel35_entry_volume, 5)
        grid_row_spin(self.bel35_entry_top_hours, 5, 3)
        grid_row_spin(self.bel35_entry_top_minutes, 5, 4)
        grid_row_spin(self.bel35_entry_trp_hours, 5, 5)
        grid_row_spin(self.bel35_entry_trp_minutes, 5, 6)
        grid_row_spin(self.bel35_entry_krp_hours, 5, 7)
        grid_row_spin(self.bel35_entry_krp_minutes, 5, 8)
        grid_row_spin(self.bel35_entry_rgp_hours, 5, 9)
        grid_row_spin(self.bel35_entry_rgp_minutes, 5, 10)
        grid_row_spin(self.bel35_entry_oip_hours, 5, 11)
        grid_row_spin(self.bel35_entry_oip_minutes, 5, 12)
        grid_row_spin(self.bel35_entry_ir_hours, 5, 13)
        grid_row_spin(self.bel35_entry_ir_minutes, 5, 14)
        grid_row_spin(self.bel35_entry_bvr_hours, 5, 15)
        grid_row_spin(self.bel35_entry_bvr_minutes, 5, 16)

        grid_sep_horizontal(self.sep_2, 6)

        grid_row_label(self.mtz_label, 7)
        grid_volume(self.mtz_entry_volume, 7)
        grid_row_spin(self.mtz_entry_top_hours, 7, 3)
        grid_row_spin(self.mtz_entry_top_minutes, 7, 4)
        grid_row_spin(self.mtz_entry_trp_hours, 7, 5)
        grid_row_spin(self.mtz_entry_trp_minutes, 7, 6)
        grid_row_spin(self.mtz_entry_krp_hours, 7, 7)
        grid_row_spin(self.mtz_entry_krp_minutes, 7, 8)
        grid_row_spin(self.mtz_entry_rgp_hours, 7, 9)
        grid_row_spin(self.mtz_entry_rgp_minutes, 7, 10)
        grid_row_spin(self.mtz_entry_oip_hours, 7, 11)
        grid_row_spin(self.mtz_entry_oip_minutes, 7, 12)
        grid_row_spin(self.mtz_entry_ir_hours, 7, 13)
        grid_row_spin(self.mtz_entry_ir_minutes, 7, 14)
        grid_row_spin(self.mtz_entry_bvr_hours, 7, 15)
        grid_row_spin(self.mtz_entry_bvr_minutes, 7, 16)
        grid_row_label(self.dt_label, 8)
        grid_volume(self.dt_entry_volume, 8)
        grid_row_spin(self.dt_entry_top_hours, 8, 3)
        grid_row_spin(self.dt_entry_top_minutes, 8, 4)
        grid_row_spin(self.dt_entry_trp_hours, 8, 5)
        grid_row_spin(self.dt_entry_trp_minutes, 8, 6)
        grid_row_spin(self.dt_entry_krp_hours, 8, 7)
        grid_row_spin(self.dt_entry_krp_minutes, 8, 8)
        grid_row_spin(self.dt_entry_rgp_hours, 8, 9)
        grid_row_spin(self.dt_entry_rgp_minutes, 8, 10)
        grid_row_spin(self.dt_entry_oip_hours, 8, 11)
        grid_row_spin(self.dt_entry_oip_minutes, 8, 12)
        grid_row_spin(self.dt_entry_ir_hours, 8, 13)
        grid_row_spin(self.dt_entry_ir_minutes, 8, 14)
        grid_row_spin(self.dt_entry_bvr_hours, 8, 15)
        grid_row_spin(self.dt_entry_bvr_minutes, 8, 16)

        grid_sep_horizontal(self.sep_3, 9)

        grid_row_label(self.cat28_label, 10)
        grid_volume(self.cat28_entry_volume, 10)
        grid_row_spin(self.cat28_entry_top_hours, 10, 3)
        grid_row_spin(self.cat28_entry_top_minutes, 10, 4)
        grid_row_spin(self.cat28_entry_trp_hours, 10, 5)
        grid_row_spin(self.cat28_entry_trp_minutes, 10, 6)
        grid_row_spin(self.cat28_entry_krp_hours, 10, 7)
        grid_row_spin(self.cat28_entry_krp_minutes, 10, 8)
        grid_row_spin(self.cat28_entry_rgp_hours, 10, 9)
        grid_row_spin(self.cat28_entry_rgp_minutes, 10, 10)
        grid_row_spin(self.cat28_entry_oip_hours, 10, 11)
        grid_row_spin(self.cat28_entry_oip_minutes, 10, 12)
        grid_row_spin(self.cat28_entry_ir_hours, 10, 13)
        grid_row_spin(self.cat28_entry_ir_minutes, 10, 14)
        grid_row_spin(self.cat28_entry_bvr_hours, 10, 15)
        grid_row_spin(self.cat28_entry_bvr_minutes, 10, 16)
        grid_row_label(self.cat33_label, 11)
        grid_volume(self.cat33_entry_volume, 11)
        grid_row_spin(self.cat33_entry_top_hours, 11, 3)
        grid_row_spin(self.cat33_entry_top_minutes, 11, 4)
        grid_row_spin(self.cat33_entry_trp_hours, 11, 5)
        grid_row_spin(self.cat33_entry_trp_minutes, 11, 6)
        grid_row_spin(self.cat33_entry_krp_hours, 11, 7)
        grid_row_spin(self.cat33_entry_krp_minutes, 11, 8)
        grid_row_spin(self.cat33_entry_rgp_hours, 11, 9)
        grid_row_spin(self.cat33_entry_rgp_minutes, 11, 10)
        grid_row_spin(self.cat33_entry_oip_hours, 11, 11)
        grid_row_spin(self.cat33_entry_oip_minutes, 11, 12)
        grid_row_spin(self.cat33_entry_ir_hours, 11, 13)
        grid_row_spin(self.cat33_entry_ir_minutes, 11, 14)
        grid_row_spin(self.cat33_entry_bvr_hours, 11, 15)
        grid_row_spin(self.cat33_entry_bvr_minutes, 11, 16)

        grid_sep_horizontal(self.sep_4, 12)
        grid_row_label(self.cat4_label, 13)
        grid_volume(self.cat4_entry_volume, 13)
        grid_row_spin(self.cat4_entry_top_hours, 13, 3)
        grid_row_spin(self.cat4_entry_top_minutes, 13, 4)
        grid_row_spin(self.cat4_entry_trp_hours, 13, 5)
        grid_row_spin(self.cat4_entry_trp_minutes, 13, 6)
        grid_row_spin(self.cat4_entry_krp_hours, 13, 7)
        grid_row_spin(self.cat4_entry_krp_minutes, 13, 8)
        grid_row_spin(self.cat4_entry_rgp_hours, 13, 9)
        grid_row_spin(self.cat4_entry_rgp_minutes, 13, 10)
        grid_row_spin(self.cat4_entry_oip_hours, 13, 11)
        grid_row_spin(self.cat4_entry_oip_minutes, 13, 12)
        grid_row_spin(self.cat4_entry_ir_hours, 13, 13)
        grid_row_spin(self.cat4_entry_ir_minutes, 13, 14)
        grid_row_spin(self.cat4_entry_bvr_hours, 13, 15)
        grid_row_spin(self.cat4_entry_bvr_minutes, 13, 16)
        grid_row_label(self.kom5_label, 14)
        grid_volume(self.kom5_entry_volume, 14)
        grid_row_spin(self.kom5_entry_top_hours, 14, 3)
        grid_row_spin(self.kom5_entry_top_minutes, 14, 4)
        grid_row_spin(self.kom5_entry_trp_hours, 14, 5)
        grid_row_spin(self.kom5_entry_trp_minutes, 14, 6)
        grid_row_spin(self.kom5_entry_krp_hours, 14, 7)
        grid_row_spin(self.kom5_entry_krp_minutes, 14, 8)
        grid_row_spin(self.kom5_entry_rgp_hours, 14, 9)
        grid_row_spin(self.kom5_entry_rgp_minutes, 14, 10)
        grid_row_spin(self.kom5_entry_oip_hours, 14, 11)
        grid_row_spin(self.kom5_entry_oip_minutes, 14, 12)
        grid_row_spin(self.kom5_entry_ir_hours, 14, 13)
        grid_row_spin(self.kom5_entry_ir_minutes, 14, 14)
        grid_row_spin(self.kom5_entry_bvr_hours, 14, 15)
        grid_row_spin(self.kom5_entry_bvr_minutes, 14, 16)

        grid_sep_horizontal(self.sep_5, 15)

        grid_row_label(self.cat15_label, 16)
        grid_volume(self.cat15_entry_volume, 16)
        grid_row_spin(self.cat15_entry_top_hours, 16, 3)
        grid_row_spin(self.cat15_entry_top_minutes, 16, 4)
        grid_row_spin(self.cat15_entry_trp_hours, 16, 5)
        grid_row_spin(self.cat15_entry_trp_minutes, 16, 6)
        grid_row_spin(self.cat15_entry_krp_hours, 16, 7)
        grid_row_spin(self.cat15_entry_krp_minutes, 16, 8)
        grid_row_spin(self.cat15_entry_rgp_hours, 16, 9)
        grid_row_spin(self.cat15_entry_rgp_minutes, 16, 10)
        grid_row_spin(self.cat15_entry_oip_hours, 16, 11)
        grid_row_spin(self.cat15_entry_oip_minutes, 16, 12)
        grid_row_spin(self.cat15_entry_ir_hours, 16, 13)
        grid_row_spin(self.cat15_entry_ir_minutes, 16, 14)
        grid_row_spin(self.cat15_entry_bvr_hours, 16, 15)
        grid_row_spin(self.cat15_entry_bvr_minutes, 16, 16)
        grid_row_label(self.cat16_label, 17)
        grid_volume(self.cat16_entry_volume, 17)
        grid_row_spin(self.cat16_entry_top_hours, 17, 3)
        grid_row_spin(self.cat16_entry_top_minutes, 17, 4)
        grid_row_spin(self.cat16_entry_trp_hours, 17, 5)
        grid_row_spin(self.cat16_entry_trp_minutes, 17, 6)
        grid_row_spin(self.cat16_entry_krp_hours, 17, 7)
        grid_row_spin(self.cat16_entry_krp_minutes, 17, 8)
        grid_row_spin(self.cat16_entry_rgp_hours, 17, 9)
        grid_row_spin(self.cat16_entry_rgp_minutes, 17, 10)
        grid_row_spin(self.cat16_entry_oip_hours, 17, 11)
        grid_row_spin(self.cat16_entry_oip_minutes, 17, 12)
        grid_row_spin(self.cat16_entry_ir_hours, 17, 13)
        grid_row_spin(self.cat16_entry_ir_minutes, 17, 14)
        grid_row_spin(self.cat16_entry_bvr_hours, 17, 15)
        grid_row_spin(self.cat16_entry_bvr_minutes, 17, 16)
        grid_row_label(self.vol25_label, 18)
        grid_volume(self.vol25_entry_volume, 18)
        grid_row_spin(self.vol25_entry_top_hours, 18, 3)
        grid_row_spin(self.vol25_entry_top_minutes, 18, 4)
        grid_row_spin(self.vol25_entry_trp_hours, 18, 5)
        grid_row_spin(self.vol25_entry_trp_minutes, 18, 6)
        grid_row_spin(self.vol25_entry_krp_hours, 18, 7)
        grid_row_spin(self.vol25_entry_krp_minutes, 18, 8)
        grid_row_spin(self.vol25_entry_rgp_hours, 18, 9)
        grid_row_spin(self.vol25_entry_rgp_minutes, 18, 10)
        grid_row_spin(self.vol25_entry_oip_hours, 18, 11)
        grid_row_spin(self.vol25_entry_oip_minutes, 18, 12)
        grid_row_spin(self.vol25_entry_ir_hours, 18, 13)
        grid_row_spin(self.vol25_entry_ir_minutes, 18, 14)
        grid_row_spin(self.vol25_entry_bvr_hours, 18, 15)
        grid_row_spin(self.vol25_entry_bvr_minutes, 18, 16)
        grid_row_label(self.sha26_label, 19)
        grid_volume(self.sha26_entry_volume, 19)
        grid_row_spin(self.sha26_entry_top_hours, 19, 3)
        grid_row_spin(self.sha26_entry_top_minutes, 19, 4)
        grid_row_spin(self.sha26_entry_trp_hours, 19, 5)
        grid_row_spin(self.sha26_entry_trp_minutes, 19, 6)
        grid_row_spin(self.sha26_entry_krp_hours, 19, 7)
        grid_row_spin(self.sha26_entry_krp_minutes, 19, 8)
        grid_row_spin(self.sha26_entry_rgp_hours, 19, 9)
        grid_row_spin(self.sha26_entry_rgp_minutes, 19, 10)
        grid_row_spin(self.sha26_entry_oip_hours, 19, 11)
        grid_row_spin(self.sha26_entry_oip_minutes, 19, 12)
        grid_row_spin(self.sha26_entry_ir_hours, 19, 13)
        grid_row_spin(self.sha26_entry_ir_minutes, 19, 14)
        grid_row_spin(self.sha26_entry_bvr_hours, 19, 15)
        grid_row_spin(self.sha26_entry_bvr_minutes, 19, 16)
        grid_row_label(self.sha30_label, 20)
        grid_volume(self.sha30_entry_volume, 20)
        grid_row_spin(self.sha30_entry_top_hours, 20, 3)
        grid_row_spin(self.sha30_entry_top_minutes, 20, 4)
        grid_row_spin(self.sha30_entry_trp_hours, 20, 5)
        grid_row_spin(self.sha30_entry_trp_minutes, 20, 6)
        grid_row_spin(self.sha30_entry_krp_hours, 20, 7)
        grid_row_spin(self.sha30_entry_krp_minutes, 20, 8)
        grid_row_spin(self.sha30_entry_rgp_hours, 20, 9)
        grid_row_spin(self.sha30_entry_rgp_minutes, 20, 10)
        grid_row_spin(self.sha30_entry_oip_hours, 20, 11)
        grid_row_spin(self.sha30_entry_oip_minutes, 20, 12)
        grid_row_spin(self.sha30_entry_ir_hours, 20, 13)
        grid_row_spin(self.sha30_entry_ir_minutes, 20, 14)
        grid_row_spin(self.sha30_entry_bvr_hours, 20, 15)
        grid_row_spin(self.sha30_entry_bvr_minutes, 20, 16)

        grid_sep_horizontal(self.sep_6, 21)

        self.accept_button.grid(row=22, columnspan=22, sticky=EW)

    def insert_data(self):

        def get_right_time(element1, element2):
            if element1 == '0' and element2 == '0':
                return None
            else:
                res_hours = element1
                lambda_minutes = lambda x: int(round(int(element2) * 1.67, 1)) if element2 != '0' else element2
                return float(f'{res_hours}.{lambda_minutes(element2)}')

        def input_result_data(
                i_truck_id,
                i_volume,
                i_top_h, i_top_m,
                i_trp_h, i_trp_m,
                i_krp_h, i_krp_m,
                i_rgp_h, i_rgp_m,
                i_oip_h, i_oip_m,
                i_ir_h, i_ir_m,
                i_bvr_h, i_bvr_m
        ):
            input_date = f'01.{months.get(self.months_box_label["text"])}.{self.year_label.cget("text")}'
            sql_insert_po6(
                input_date,
                i_truck_id,
                i_volume,
                get_right_time(i_top_h, i_top_m),
                get_right_time(i_trp_h, i_trp_m),
                get_right_time(i_krp_h, i_krp_m),
                get_right_time(i_rgp_h, i_rgp_m),
                get_right_time(i_oip_h, i_oip_m),
                get_right_time(i_ir_h, i_ir_m),
                get_right_time(i_bvr_h, i_bvr_m),
            )

        input_result_data(
            1,
            self.mtz_entry_volume.get(),
            self.mtz_entry_top_hours.get(), self.mtz_entry_top_minutes.get(),
            self.mtz_entry_trp_hours.get(), self.mtz_entry_trp_minutes.get(),
            self.mtz_entry_krp_hours.get(), self.mtz_entry_krp_minutes.get(),
            self.mtz_entry_rgp_hours.get(), self.mtz_entry_rgp_minutes.get(),
            self.mtz_entry_oip_hours.get(), self.mtz_entry_oip_minutes.get(),
            self.mtz_entry_ir_hours.get(), self.mtz_entry_ir_minutes.get(),
            self.mtz_entry_bvr_hours.get(), self.mtz_entry_bvr_minutes.get()
        )
        input_result_data(
            4,
            self.cat4_entry_volume.get(),
            self.cat4_entry_top_hours.get(), self.cat4_entry_top_minutes.get(),
            self.cat4_entry_trp_hours.get(), self.cat4_entry_trp_minutes.get(),
            self.cat4_entry_krp_hours.get(), self.cat4_entry_krp_minutes.get(),
            self.cat4_entry_rgp_hours.get(), self.cat4_entry_rgp_minutes.get(),
            self.cat4_entry_oip_hours.get(), self.cat4_entry_oip_minutes.get(),
            self.cat4_entry_ir_hours.get(), self.cat4_entry_ir_minutes.get(),
            self.cat4_entry_bvr_hours.get(), self.cat4_entry_bvr_minutes.get()
        )
        input_result_data(
            15,
            self.kom5_entry_volume.get(),
            self.kom5_entry_top_hours.get(), self.kom5_entry_top_minutes.get(),
            self.kom5_entry_trp_hours.get(), self.kom5_entry_trp_minutes.get(),
            self.kom5_entry_krp_hours.get(), self.kom5_entry_krp_minutes.get(),
            self.kom5_entry_rgp_hours.get(), self.kom5_entry_rgp_minutes.get(),
            self.kom5_entry_oip_hours.get(), self.kom5_entry_oip_minutes.get(),
            self.kom5_entry_ir_hours.get(), self.kom5_entry_ir_minutes.get(),
            self.kom5_entry_bvr_hours.get(), self.kom5_entry_bvr_minutes.get()
        )
        input_result_data(
            9,
            self.cat28_entry_volume.get(),
            self.cat28_entry_top_hours.get(), self.cat28_entry_top_minutes.get(),
            self.cat28_entry_trp_hours.get(), self.cat28_entry_trp_minutes.get(),
            self.cat28_entry_krp_hours.get(), self.cat28_entry_krp_minutes.get(),
            self.cat28_entry_rgp_hours.get(), self.cat28_entry_rgp_minutes.get(),
            self.cat28_entry_oip_hours.get(), self.cat28_entry_oip_minutes.get(),
            self.cat28_entry_ir_hours.get(), self.cat28_entry_ir_minutes.get(),
            self.cat28_entry_bvr_hours.get(), self.cat28_entry_bvr_minutes.get()
        )
        input_result_data(
            10,
            self.cat33_entry_volume.get(),
            self.cat33_entry_top_hours.get(), self.cat33_entry_top_minutes.get(),
            self.cat33_entry_trp_hours.get(), self.cat33_entry_trp_minutes.get(),
            self.cat33_entry_krp_hours.get(), self.cat33_entry_krp_minutes.get(),
            self.cat33_entry_rgp_hours.get(), self.cat33_entry_rgp_minutes.get(),
            self.cat33_entry_oip_hours.get(), self.cat33_entry_oip_minutes.get(),
            self.cat33_entry_ir_hours.get(), self.cat33_entry_ir_minutes.get(),
            self.cat33_entry_bvr_hours.get(), self.cat33_entry_bvr_minutes.get()
        )
        input_result_data(
            28,
            self.bel32_entry_volume.get(),
            self.bel32_entry_top_hours.get(), self.bel32_entry_top_minutes.get(),
            self.bel32_entry_trp_hours.get(), self.bel32_entry_trp_minutes.get(),
            self.bel32_entry_krp_hours.get(), self.bel32_entry_krp_minutes.get(),
            self.bel32_entry_rgp_hours.get(), self.bel32_entry_rgp_minutes.get(),
            self.bel32_entry_oip_hours.get(), self.bel32_entry_oip_minutes.get(),
            self.bel32_entry_ir_hours.get(), self.bel32_entry_ir_minutes.get(),
            self.bel32_entry_bvr_hours.get(), self.bel32_entry_bvr_minutes.get()
        )
        input_result_data(
            31,
            self.bel35_entry_volume.get(),
            self.bel35_entry_top_hours.get(), self.bel35_entry_top_minutes.get(),
            self.bel35_entry_trp_hours.get(), self.bel35_entry_trp_minutes.get(),
            self.bel35_entry_krp_hours.get(), self.bel35_entry_krp_minutes.get(),
            self.bel35_entry_rgp_hours.get(), self.bel35_entry_rgp_minutes.get(),
            self.bel35_entry_oip_hours.get(), self.bel35_entry_oip_minutes.get(),
            self.bel35_entry_ir_hours.get(), self.bel35_entry_ir_minutes.get(),
            self.bel35_entry_bvr_hours.get(), self.bel35_entry_bvr_minutes.get()
        )
        input_result_data(
            43,
            self.cat16_entry_volume.get(),
            self.cat16_entry_top_hours.get(), self.cat16_entry_top_minutes.get(),
            self.cat16_entry_trp_hours.get(), self.cat16_entry_trp_minutes.get(),
            self.cat16_entry_krp_hours.get(), self.cat16_entry_krp_minutes.get(),
            self.cat16_entry_rgp_hours.get(), self.cat16_entry_rgp_minutes.get(),
            self.cat16_entry_oip_hours.get(), self.cat16_entry_oip_minutes.get(),
            self.cat16_entry_ir_hours.get(), self.cat16_entry_ir_minutes.get(),
            self.cat16_entry_bvr_hours.get(), self.cat16_entry_bvr_minutes.get()
        )
        input_result_data(
            45,
            self.esh_entry_volume.get(),
            self.esh_entry_top_hours.get(), self.esh_entry_top_minutes.get(),
            self.esh_entry_trp_hours.get(), self.esh_entry_trp_minutes.get(),
            self.esh_entry_krp_hours.get(), self.esh_entry_krp_minutes.get(),
            self.esh_entry_rgp_hours.get(), self.esh_entry_rgp_minutes.get(),
            self.esh_entry_oip_hours.get(), self.esh_entry_oip_minutes.get(),
            self.esh_entry_ir_hours.get(), self.esh_entry_ir_minutes.get(),
            self.esh_entry_bvr_hours.get(), self.esh_entry_bvr_minutes.get()
        )
        input_result_data(
            58,
            self.cat15_entry_volume.get(),
            self.cat15_entry_top_hours.get(), self.cat15_entry_top_minutes.get(),
            self.cat15_entry_trp_hours.get(), self.cat15_entry_trp_minutes.get(),
            self.cat15_entry_krp_hours.get(), self.cat15_entry_krp_minutes.get(),
            self.cat15_entry_rgp_hours.get(), self.cat15_entry_rgp_minutes.get(),
            self.cat15_entry_oip_hours.get(), self.cat15_entry_oip_minutes.get(),
            self.cat15_entry_ir_hours.get(), self.cat15_entry_ir_minutes.get(),
            self.cat15_entry_bvr_hours.get(), self.cat15_entry_bvr_minutes.get()
        )
        input_result_data(
            272,
            self.dt_entry_volume.get(),
            self.dt_entry_top_hours.get(), self.dt_entry_top_minutes.get(),
            self.dt_entry_trp_hours.get(), self.dt_entry_trp_minutes.get(),
            self.dt_entry_krp_hours.get(), self.dt_entry_krp_minutes.get(),
            self.dt_entry_rgp_hours.get(), self.dt_entry_rgp_minutes.get(),
            self.dt_entry_oip_hours.get(), self.dt_entry_oip_minutes.get(),
            self.dt_entry_ir_hours.get(), self.dt_entry_ir_minutes.get(),
            self.dt_entry_bvr_hours.get(), self.dt_entry_bvr_minutes.get()
        )
        input_result_data(
            280,
            self.vol25_entry_volume.get(),
            self.vol25_entry_top_hours.get(), self.vol25_entry_top_minutes.get(),
            self.vol25_entry_trp_hours.get(), self.vol25_entry_trp_minutes.get(),
            self.vol25_entry_krp_hours.get(), self.vol25_entry_krp_minutes.get(),
            self.vol25_entry_rgp_hours.get(), self.vol25_entry_rgp_minutes.get(),
            self.vol25_entry_oip_hours.get(), self.vol25_entry_oip_minutes.get(),
            self.vol25_entry_ir_hours.get(), self.vol25_entry_ir_minutes.get(),
            self.vol25_entry_bvr_hours.get(), self.vol25_entry_bvr_minutes.get()
        )
        input_result_data(
            283,
            self.sha26_entry_volume.get(),
            self.sha26_entry_top_hours.get(), self.sha26_entry_top_minutes.get(),
            self.sha26_entry_trp_hours.get(), self.sha26_entry_trp_minutes.get(),
            self.sha26_entry_krp_hours.get(), self.sha26_entry_krp_minutes.get(),
            self.sha26_entry_rgp_hours.get(), self.sha26_entry_rgp_minutes.get(),
            self.sha26_entry_oip_hours.get(), self.sha26_entry_oip_minutes.get(),
            self.sha26_entry_ir_hours.get(), self.sha26_entry_ir_minutes.get(),
            self.sha26_entry_bvr_hours.get(), self.sha26_entry_bvr_minutes.get()
        )
        input_result_data(
            284,
            self.sha30_entry_volume.get(),
            self.sha30_entry_top_hours.get(), self.sha30_entry_top_minutes.get(),
            self.sha30_entry_trp_hours.get(), self.sha30_entry_trp_minutes.get(),
            self.sha30_entry_krp_hours.get(), self.sha30_entry_krp_minutes.get(),
            self.sha30_entry_rgp_hours.get(), self.sha30_entry_rgp_minutes.get(),
            self.sha30_entry_oip_hours.get(), self.sha30_entry_oip_minutes.get(),
            self.sha30_entry_ir_hours.get(), self.sha30_entry_ir_minutes.get(),
            self.sha30_entry_bvr_hours.get(), self.sha30_entry_bvr_minutes.get()
        )
        self.destroy()
        messagebox.showinfo(title='Внесение данных', message='Данные внесены')
