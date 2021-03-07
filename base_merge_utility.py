import openpyxl
import sys
import glob

# ================================================================
# 0=SETTINGS
mask_file_base = "*_ad???@?adis???rg_*.xlsx"
mask_file_price = "*Прайс*.xlsx"
file_opend_startwith_symbols = "~$"

column_name_code_base = "Код"
column_values_code_base_null_set = {column_name_code_base, None, ""}


# ================================================================
# 1=DETECT FILES
# ----------------------------------------------------------------
# 1=FILE - MAIN BASE
file_name_main_base_found_list = glob.glob(mask_file_base)
print("Найдены файлы главной базы:", file_name_main_base_found_list)

def found_several_bases():
    print("ОШИБКА: найдено несколько файлов главной базы! уберите лишние! и перезапустите приложение", file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

if len(file_name_main_base_found_list) == 1:
    file_name_main_base = file_name_main_base_found_list[0]
elif len(file_name_main_base_found_list) == 2:
    for file in file_name_main_base_found_list:
        if file.startswith(file_opend_startwith_symbols):
            print("ОШИБКА: файл главной базы открыт! закройте его и перезапустите приложение!", file=sys.stderr)
            input("Нажмите Enter для выхода")
            sys.exit()
    else:
        found_several_bases()
else:
    found_several_bases()
# ----------------------------------------------------------------
# 2=FILES - NEW PRICE
file_name_price_new_found_list = glob.glob(mask_file_price)
print("Найдены файлы новых цен:", file_name_price_new_found_list)

if len(file_name_price_new_found_list) > 1:
    print("ОШИБКА: пока утилита работает только с одним файлом новых цен! оставьте только один файл и перезапустите приложение!",
          file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

# пока безсмысленный код поиска открытых!!!
for file in file_name_price_new_found_list:
    if file.startswith(file_opend_startwith_symbols):
        print("ОШИБКА: найдены открытые файлы новых цен! закройте их и перезапустите приложение!", file=sys.stderr)
        input("Нажмите Enter для выхода")
        sys.exit()

file_name_price_new = file_name_price_new_found_list[0]
# ================================================================
# 2=INIT FILES
wb_base = openpyxl.load_workbook(file_name_main_base)
ws_base = wb_base.active

wb_price_new = openpyxl.load_workbook(file_name_price_new)
ws_price_new = wb_price_new.active

# ================================================================
# 3=DATA LOAD
# ----------------------------------------------------------------
# 1=MainBASE
# --------------------------
# 1=DETECT COLUMNS IN HEADER
row_title = ws_base[1]
column_index_base_from_1 = 1
for cell in row_title:
    cell_value = cell.value
    if cell_value == column_name_code_base:
        column_index_base_code = column_index_base_from_1
        print(f"Колонка Код = {column_index_base_code}")
    elif cell_value == "Артикул":
        column_index_base_art = column_index_base_from_1
        print(f"Колонка Артикул = {column_index_base_art}")
    column_index_base_from_1 += 1

# --------------------------
# 2=DATA - load ColumnCODE
column_values_code_base_all_dict = dict()
column_values_code_base_repeated_set = set()

print("*"*80)
column_values_code_base_iter = ws_base.iter_cols(min_col=column_index_base_code, max_col=column_index_base_code)
for column_tuple in column_values_code_base_iter:
    for cell_obj in column_tuple:
        cell_value = cell_obj.value
        if cell_value not in column_values_code_base_all_dict:
            print("+", cell_value)
            column_values_code_base_all_dict.update({cell_value: {"cell_obj_list": [cell_obj, ]}})
        else:
            print("-"*10, cell_value)
            column_values_code_base_all_dict[cell_value]["cell_obj_list"].append(cell_obj)
            column_values_code_base_repeated_set.update({cell_value})
print("*"*80)
print("column_values_code_base_repeated_set:", column_values_code_base_repeated_set)
print("*"*80)

# ----------------------------------------------------------------
# 2=PRICE NEW
# --------------------------
# 1=DETECT COLUMNS IN HEADER
column_index_price_code = 1
