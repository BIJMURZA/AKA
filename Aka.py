import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Задаем пути к файлам и названия листов
excel_file_path = r'C:\Users\a.kostyunin\Desktop\New_Oct_Copy\For_New_Tests.xlsx'
sheet_name = 'Реестр договров'
forma_sheet_name = 'Narabotka_Test'
excel_file_save = r'C:\Users\a.kostyunin\Desktop\New_Oct_Copy\Форма_Для_Наработки-Выручки.xlsx'

# Функции для работы с данными
def load_data(excel_file_path, sheet_name):
    return pd.read_excel(excel_file_path, sheet_name=sheet_name)

def filter_data(df, new_abonent_value):
    return df[df['Новый абонент'] == new_abonent_value]

def create_pivot_table(filtered_df, values_to_sum):
    return pd.pivot_table(filtered_df, values=values_to_sum, index=['Филиал'], columns=['Новый абонент'], aggfunc='sum', fill_value=0, margins=True, margins_name='Total')

def calculate_total_sum(pivot_table):
    return pivot_table.sum(axis=1)

def create_total_sum_df(index, total_sum):
    return pd.DataFrame({'Филиал': index, 'Total Sum': total_sum}).reset_index(drop=True)

def merge_data(original_df, total_sum_df):
    return pd.merge(original_df, total_sum_df, on='Филиал', how='left')

def update_excel_sheet(excel_file_save, sheet_name, data):
    wb = openpyxl.load_workbook(excel_file_save)
    ws = wb[sheet_name]
    for r_idx, row in enumerate(dataframe_to_rows(data, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(excel_file_save)

# Основной процесс
def process_data_to_excel(excel_file_path, sheet_name, values_to_sum, target_sheet, start_row):
    df = load_data(excel_file_path, sheet_name)
    for new_abonent_value in ['да', 'нет']:
        filtered_df = filter_data(df, new_abonent_value)
        pivot_table = create_pivot_table(filtered_df, values_to_sum)
        total_sum = calculate_total_sum(pivot_table)
        total_sum_df = create_total_sum_df(pivot_table.index, total_sum)
        merged_df = merge_data(df, total_sum_df)
        final_data = pd.concat([total_sum_df['Филиал'], total_sum_df['Total Sum']], axis=1)
        update_excel_sheet(excel_file_save, target_sheet, final_data)

# Значения для суммирования (предположим, что они такие же)
values_to_sum_1 = [                                                                                                        
    'Подключение', 'Точка доступа', 'Подключение Дозор', 'Продажа IPTV-приставки', 'Подключение Wi-Fi', 'Продажа роутера', 
    'Продажа камеры', 'Платный выезд', 'ivi', 'Смена тарифа', 'Обещанный платеж', 'Добровольная блокировка', 'Тест-драйв', 
    'Годовой абонемент', 'Перетяжка', 'Возврат ДС', 'Дебиторка', 'Остальные расходы'                                       
]                                                                                                                          
                                                                                                                           
values_to_sum_2 = [                                                                                                        
    'Рассрочка подключение, наработка', 'Рассрочка IPTV-приставка, наработка старая', 'Рассрочка IPTV-приставка, наработка 
    'Рассрочка IPTV-приставка, наработка неопределенная', 'Рассрочка Wi-Fi, наработка', 'Рассрочка подключение Дозор, нараб
    'Тариф ШПД, наработка', 'ТВ-пакеты, наработка', 'Аренда терминала, наработка', 'Аренда Wi-Fi, наработка в тарифе', 'Аре
    'Аренда Wi-Fi, наработка неопределенная', 'Аренда IPTV-приставки, наработка в тарифе', 'Аренда IPTV-приставки, наработк
    'Аренда IPTV-приставки, наработка неопределенная', 'Белый IP, наработка', 'Тариф Дозор, наработка', 'Подписки, наработк
]                                                                                                                          

# Выполнение обработки данных и обновление Excel
process_data_to_excel(excel_file_path, sheet_name, values_to_sum_1, forma_sheet_name, 4)
process_data_to_excel(excel_file_path, sheet_name, values_to_sum_2, forma_sheet_name, 29)
