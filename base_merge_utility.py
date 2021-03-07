"""
ПОДГОТОВКА ФАЙЛОВ
1. все файлы баз и прайсов должны быть закрыты!
2. файл главной базы должен соответствовать маске [*_ad???@?adis???rg_*.xlsx]
3. файлы новых прайсов
    - содержать имя поставщика например "surgaz" (в любой раскладке)
    - файл прайс от одного вендора должен быть только 1!
"""
import openpyxl
import sys
import glob

# ================================================================
# 0=SETTINGS
mask_file_base = "*_ad???@?adis???rg_*.xlsx"
file_opened_startwith_symbols = "~$"

vendor_dict = {"surgaz": {"file_mask": "*surgaz*.xlsx", "file_found_if_one": None,
                          "column_article_int": 1, "data_article": dict()}}

column_name_code_base = "Код"
column_name_art_base = "Артикул"
column_values_code_base_null_set = {column_name_code_base, None, ""}

# ================================================================
# 1=DETECT FILES
# ----------------------------------------------------------------
# 1=FILE - MAIN BASE
print("*"*80)
print(f"START finding files MainBASE")
print()

file_name_main_base_found_list = glob.glob(mask_file_base)
print("Найдены файлы главной базы:", file_name_main_base_found_list)

def files_found_several(type_txt):
    print(f"ОШИБКА: найдено несколько файлов {type_txt}! уберите лишние! и перезапустите приложение", file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

def files_found_opened(type_txt):
    print(f"ОШИБКА: найден открытый файл {type_txt}! закройте его и перезапустите приложение", file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

if len(file_name_main_base_found_list) == 1:
    file_name_main_base = file_name_main_base_found_list[0]
elif len(file_name_main_base_found_list) > 1:
    for file in file_name_main_base_found_list:
        if file.startswith(file_opened_startwith_symbols):
            files_found_opened("главной базы")
    else:
        files_found_several("главной базы")

# ----------------------------------------------------------------
# 2=FILES - VENDORS
print("*"*80)
print(f"START finding files VENDORS")
print()

for vendor in vendor_dict:
    print("-" * 80)
    print(f"VENDOR [{vendor}]")

    vendor_data_dict = vendor_dict[vendor]
    mask_vendor_file = vendor_data_dict["file_mask"]
    file_name_vendor_found_list = glob.glob(mask_vendor_file)
    print(f"Найдены файлы вендора [{vendor}]: {file_name_vendor_found_list}")

    if len(file_name_vendor_found_list) == 1:
        vendor_data_dict["file_found_if_one"] = file_name_vendor_found_list[0]
    elif len(file_name_vendor_found_list) > 1:
        for file in file_name_vendor_found_list:
            if file.startswith(file_opened_startwith_symbols):
                files_found_opened(f"вендора [{vendor}]")
        else:
            files_found_several(f"вендора [{vendor}]")

# ================================================================
# 2=MainBASE=WORK
# --------------------------
# 1=file INIT
wb_base = openpyxl.load_workbook(file_name_main_base)
ws_base = wb_base.active

# --------------------------
# 2=DETECT COLUMNS IN HEADER
row_title = ws_base[1]
column_index_base_from_1 = 1
for cell in row_title:
    cell_value = cell.value
    if cell_value == column_name_code_base:
        column_index_base_code = column_index_base_from_1
        print(f"Номер колонки [{column_name_code_base}] = [{column_index_base_code}]")
    elif cell_value == column_name_art_base:
        column_index_base_art = column_index_base_from_1
        print(f"Номер колонки [{column_name_art_base}] = [{column_index_base_art}]")
    column_index_base_from_1 += 1

# --------------------------
# 3=load DATA - ColumnCODE
column_values_code_base_all_dict = dict()
column_values_code_base_repeated_set = set()

print("*"*80)
print(f"load data from file: [{file_name_main_base}]")
print("*"*80)
column_values_code_base_iter = ws_base.iter_cols(min_col=column_index_base_code, max_col=column_index_base_code)
for column_tuple in column_values_code_base_iter:
    for cell_obj in column_tuple:
        cell_value = cell_obj.value
        if cell_value not in column_values_code_base_all_dict:
            # print("+", cell_value)
            column_values_code_base_all_dict.update({cell_value: {"cell_obj_list": [cell_obj, ]}})
        else:
            print(f'{"-" * 10} found repeated value: [{cell_value}]')
            column_values_code_base_all_dict[cell_value]["cell_obj_list"].append(cell_obj)
            column_values_code_base_repeated_set.update({cell_value})

# --------------------------
# 4=print loadRESULTS
count_column_values_code_base_all_dict = len(column_values_code_base_all_dict)
count_column_values_code_base_repeated_set = len(column_values_code_base_repeated_set)

print("*"*80)
# print("from file:", file_name_main_base)
print("count_column_values_code_base_all_dict:", count_column_values_code_base_all_dict)
print("count_column_values_code_base_repeated_set:", count_column_values_code_base_repeated_set)

print("column_values_code_base_repeated_set:", column_values_code_base_repeated_set)
print("*"*80)

# ================================================================
# 3=VENDOR FILES=WORK
# --------------------------
# 1=file INIT
wb_vendor = openpyxl.load_workbook(file_name_vendor)
ws_vendor = wb_vendor.active

# --------------------------
# 2=DETECT COLUMNS IN HEADER
column_index_vendor_article = 1

# --------------------------
# 3=load DATA - ColumnCODE
column_values_article_vendor_all_dict = dict()
column_values_article_vendor_repeated_set = set()

print("*"*80)
print(f"load data from file: [{file_name_vendor}]")
print("*"*80)
column_values_art_price_iter = ws_price_new.iter_cols(min_col=column_index_vendor_article, max_col=column_index_vendor_article)
for column_tuple in column_values_price_iter:
    for cell_obj in column_tuple:
        cell_value = cell_obj.value
        if cell_value not in column_values_code_base_all_dict:
            # print("+", cell_value)
            column_values_code_base_all_dict.update({cell_value: {"cell_obj_list": [cell_obj, ]}})
        else:
            print(f'{"-" * 10} found repeated value: [{cell_value}]')
            column_values_code_base_all_dict[cell_value]["cell_obj_list"].append(cell_obj)
            column_values_code_base_repeated_set.update({cell_value})

# --------------------------
# 4=print loadRESULTS
count_column_values_code_base_all_dict = len(column_values_code_base_all_dict)
count_column_values_code_base_repeated_set = len(column_values_code_base_repeated_set)

print("*"*80)
# print("from file:", file_name_main_base)
print("count_column_values_code_base_all_dict:", count_column_values_code_base_all_dict)
print("count_column_values_code_base_repeated_set:", count_column_values_code_base_repeated_set)

print("column_values_code_base_repeated_set:", column_values_code_base_repeated_set)
print("*"*80)
