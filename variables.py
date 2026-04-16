import oracledb
import datetime
import os
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

path = os.getenv('PATH_DIR')
report_path = os.getenv('REPORT_DIR')
p_date = datetime.datetime.now().strftime('%d.%m.%Y').split('.')
p_time = datetime.datetime.now().strftime('%H.%M.%S')
oracledb.init_oracle_client(config_dir=os.getenv('CONFIG_DIR'))
connection = oracledb.connect(user=os.getenv('USER'), password=os.getenv('PASSWORD'), dsn=os.getenv('DNS'))
cursor = connection.cursor()

months = {
    'Январь': '01',
    'Февраль': '02',
    'Март': '03',
    'Апрель': '04',
    'Май': '05',
    'Июнь': '06',
    'Июль': '07',
    'Август': '08',
    'Сентябрь': '09',
    'Октябрь': '10',
    'Ноябрь': '11',
    'Декабрь': '12'
}

techs = {
    1: 'Экскаватор ЭШ 6/45 №17',
    2: 'Самосвал БелАЗ №32',
    3: 'Самосвал БелАЗ №35',
    4: 'Трактор МТЗ-82 №25',
    5: 'Будьдозер ДТ-75А №38',
    6: 'Будьдозер CAT-D6R №28',
    7: 'Будьдозер CAT-D9R №33',
    8: 'Автогрейдер CAT-140М №4',
    9: 'Автогрейдер Komatsu GD825А-2 №5',
    10: 'Погрузчик CAT-962 №15',
    11: 'Погрузчик CAT-988H №16',
    12: 'Погрузчик Volvo L120G №25',
    13: 'Погрузчик Shantui SL60W №26',
    14: 'Погрузчик Shantui SL60W №30'
}

colors = {
    'top': '#fce5cd',
    'trp': '#fff2cc',
    'krp': '#d9ead3',
    'rgp': '#d0e0e3',
    'oip': '#cfe2f3',
    'ir': '#d9d2e9',
    'bvr': '#ead1dc'
}


def sql_shovs_year(date):
    cur_shov_year = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_SHOVSY_NEW', [date, cur_shov_year])
    return cur_shov_year


def sql_veh_year(date):
    cur_veh_year = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_VEHY_NEW', [date, cur_veh_year])
    return cur_veh_year


def sql_bul_year(date):
    cur_bul_year = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_BULY_NEW', [date, cur_bul_year])
    return cur_bul_year


def sql_pogr_year(date):
    cur_pogr_year = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_POGRY_NEW', [date, cur_pogr_year])
    return cur_pogr_year


def sql_shovs_month(date):
    cur_shov_month = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_SHOVSM_NEW', [date, cur_shov_month])
    return cur_shov_month


def sql_veh_month(date):
    cur_veh_month = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_VEHM_NEW', [date, cur_veh_month])
    return cur_veh_month


def sql_bul_month(date):
    cur_bul_month = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_BULM_NEW', [date, cur_bul_month])
    return cur_bul_month


def sql_pogr_month(date):
    cur_pogr_month = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_POGRM_NEW', [date, cur_pogr_month])
    return cur_pogr_month


def sql_shovs_day(date, shift):
    cur_shovs_day = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_SHOVS_NEW', [date, shift, cur_shovs_day])
    return cur_shovs_day


def sql_veh_day(date, shift):
    cur_veh_day = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_VEH_NEW', [date, shift, cur_veh_day])
    return cur_veh_day


def sql_bul_day(date, shift):
    cur_bul_day = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_BUL_NEW', [date, shift, cur_bul_day])
    return cur_bul_day


def sql_pogr_day(date, shift):
    cur_pogr_day = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.REPORTS.GET_POGR_NEW', [date, shift, cur_pogr_day])
    return cur_pogr_day


def sql_get_po6(date):
    cur_get_po6 = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.EXPORTFULLREPORT.GET_PO6', [date, cur_get_po6])
    return cur_get_po6


def sql_insert_po6(p_month, p_truck_key, p_volume, p_top, p_trp, p_krp, p_rgp, p_oip, p_ir, p_bvr):
    # cur_get_po6 = cursor.var(oracledb.CURSOR)
    cursor.callproc('ZSU.INSERT_PO6', [p_month, p_truck_key, p_volume, p_top, p_trp, p_krp, p_rgp, p_oip, p_ir, p_bvr])
    # return cur_get_po6


def insert_technics(type_tech, ws, row_tech, df):
    if type_tech == 'shov' and len(df) == 27:
        ws[f'F{row_tech}'] = df[1]  # 6 ДВС
        ws[f'K{row_tech}'] = df[2]  # 11 объем работ план
        ws[f'L{row_tech}'] = df[3]  # 12 объем работ факт
        ws[f'N{row_tech}'] = df[4]  # 14 факт ТО
        ws[f'P{row_tech}'] = df[5]  # 16 факт ТР
        ws[f'R{row_tech}'] = df[6]  # 18 фатк КР
        ws[f'T{row_tech}'] = df[7]  # 20 факт регламент
        ws[f'V{row_tech}'] = df[8]  # 22 факт обед
        ws[f'X{row_tech}'] = df[9]  # 24 факт прием/передача
        ws[f'Z{row_tech}'] = df[10]  # 26 факт забой
        ws[f'AB{row_tech}'] = df[11]  # 28 факт БВР
        ws[f'AC{row_tech}'] = df[12]  # 29 ДВС
        ws[f'AD{row_tech}'] = df[13]  # 30 трансмиссия
        ws[f'AE{row_tech}'] = df[14]  # 31 ходовая
        ws[f'AF{row_tech}'] = df[15]  # 32 навесное
        ws[f'AG{row_tech}'] = df[16]  # 33 электро
        ws[f'AH{row_tech}'] = df[17]  # 34 гибравлика
        ws[f'AI{row_tech}'] = df[18]  # 35 прочие
        ws[f'AJ{row_tech}'] = df[19]  # 36 автошины
        ws[f'AK{row_tech}'] = df[20]  # 37 фронт работ
        ws[f'AL{row_tech}'] = df[21]  # 38 зап.части
        ws[f'AM{row_tech}'] = df[22]  # 39 ГСМ
        ws[f'AN{row_tech}'] = df[23]  # 40 персонал
        ws[f'AO{row_tech}'] = df[24]  # 41 метеоусловия
        ws[f'AP{row_tech}'] = df[25]  # 42 прочие
        ws[f'AT{row_tech}'] = float(df[26])  # 46 фонд времени
    if type_tech == 'veh' and len(df) == 28:
        ws[f'F{row_tech}'] = df[1]  # 6 ДВС
        ws[f'H{row_tech}'] = df[2]  # 8 пробег общий
        ws[f'I{row_tech}'] = df[3]  # 9 пробег с грузом
        ws[f'K{row_tech}'] = df[4]  # 11 объем работ план
        ws[f'L{row_tech}'] = df[5]  # 12 объем работ факт
        ws[f'N{row_tech}'] = df[6]  # 14 факт ТО
        ws[f'P{row_tech}'] = df[7]  # 16 факт ТР
        ws[f'R{row_tech}'] = df[8]  # 18 фатк КР
        ws[f'T{row_tech}'] = df[9]  # 20 факт регламент
        ws[f'V{row_tech}'] = df[10]  # 22 факт обед
        ws[f'X{row_tech}'] = df[11]  # 24 факт прием/передача
        ws[f'AB{row_tech}'] = df[12]  # 28 факт БВР
        ws[f'AC{row_tech}'] = df[13]  # 29 ДВС
        ws[f'AD{row_tech}'] = df[14]  # 30 трансмиссия
        ws[f'AE{row_tech}'] = df[15]  # 31 ходовая
        ws[f'AF{row_tech}'] = df[16]  # 32 навесное
        ws[f'AG{row_tech}'] = df[17]  # 33 электро
        ws[f'AH{row_tech}'] = df[18]  # 34 гибравлика
        ws[f'AI{row_tech}'] = df[19]  # 35 прочие
        ws[f'AJ{row_tech}'] = df[20]  # 36 автошины
        ws[f'AK{row_tech}'] = df[21]  # 37 фронт работ
        ws[f'AL{row_tech}'] = df[22]  # 38 зап.части
        ws[f'AM{row_tech}'] = df[23]  # 39 ГСМ
        ws[f'AN{row_tech}'] = df[24]  # 40 персонал
        ws[f'AO{row_tech}'] = df[25]  # 41 метеоусловия
        ws[f'AP{row_tech}'] = df[26]  # 42 прочие
        ws[f'AT{row_tech}'] = float(df[27])  # 46 фонд времени
    if type_tech in ['bul', 'pogr'] and len(df) == 24:
        ws[f'F{row_tech}'] = df[1]  # 6 ДВС
        ws[f'N{row_tech}'] = df[2]  # 14 факт ТО
        ws[f'P{row_tech}'] = df[3]  # 16 факт ТР
        ws[f'R{row_tech}'] = df[4]  # 18 фатк КР
        ws[f'T{row_tech}'] = df[5]  # 20 факт регламент
        ws[f'V{row_tech}'] = df[6]  # 22 факт обед
        ws[f'X{row_tech}'] = df[7]  # 24 факт прием/передача
        ws[f'AB{row_tech}'] = df[8]  # 28 факт БВР
        ws[f'AC{row_tech}'] = df[9]  # 29 ДВС
        ws[f'AD{row_tech}'] = df[10]  # 30 трансмиссия
        ws[f'AE{row_tech}'] = df[11]  # 31 ходовая
        ws[f'AF{row_tech}'] = df[12]  # 32 навесное
        ws[f'AG{row_tech}'] = df[13]  # 33 электро
        ws[f'AH{row_tech}'] = df[14]  # 34 гибравлика
        ws[f'AI{row_tech}'] = df[15]  # 35 прочие
        ws[f'AJ{row_tech}'] = df[16]  # 36 автошины
        ws[f'AK{row_tech}'] = df[17]  # 37 фронт работ
        ws[f'AL{row_tech}'] = df[18]  # 38 зап.части
        ws[f'AM{row_tech}'] = df[19]  # 39 ГСМ
        ws[f'AN{row_tech}'] = df[20]  # 40 персонал
        ws[f'AO{row_tech}'] = df[21]  # 41 метеоусловия
        ws[f'AP{row_tech}'] = df[22]  # 42 прочие
        ws[f'AT{row_tech}'] = float(df[23])  # 46 фонд времени
