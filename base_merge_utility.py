import openpyxl
import sys
import glob

# ================================================================
# 1=DETECT FILES
# ----------------------------------------------------------------
# 1=FILE - MAIN BASE
file_name_main_base_found_list = glob.glob("*_admin@madisonorg_*.xlsx")
print("Найдены файлы главной базы:", file_name_main_base_found_list)

def found_several_bases():
    print("ОШИБКА: найдено несколько файлов главной базы! уберите лишние! и перезапустите приложение", file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

if len(file_name_main_base_found_list) == 1:
    file_name_main_base = file_name_main_base_found_list[0]
elif len(file_name_main_base_found_list) == 2:
    for file in file_name_main_base_found_list:
        if file.startswith("~$"):
            print("ОШИБКА: файл главной базы открыт! закройте его и перезапустите приложение!", file=sys.stderr)
            input("Нажмите Enter для выхода")
            sys.exit()
    else:
        found_several_bases()
else:
    found_several_bases()
# ----------------------------------------------------------------
# 2=FILES - NEW PRICE
file_name_price_new_found_list = glob.glob("*Прайс*.xlsx")
print("Найдены файлы новых цен:", file_name_price_new_found_list)

if len(file_name_price_new_found_list) > 1:
    print("ОШИБКА: пока утилита работает только с одним файлом новых цен! оставьте только один файл и перезапустите приложение!",
          file=sys.stderr)
    input("Нажмите Enter для выхода")
    sys.exit()

# пока безсмысленный код поиска открытых!!!
for file in file_name_price_new_found_list:
    if file.startswith("~$"):
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
row_title = ws[1]
column_index_from_1 = 1
for cell in row_title:
    cell_value = cell.value
    if cell_value == "Код":
        column_index_code = column_index_from_1
        print(f"Колонка Код = {column_index_code}")
    elif cell_value == "Артикул":
        column_index_art = column_index_from_1
        print(f"Колонка Артикул = {column_index_art}")
    column_index_from_1 += 1

# --------------------------
# 2=DATA - load ColumnCODE
column_values_code_iter = ws.iter_cols(min_row=2, min_col=column_index_code, max_col=column_index_code, values_only=True)
column_values_code_list = list(column_values_code_iter)[0]      #[(1, 4, 7)][0] = (1, 4, 7)
print(column_values_code_list)

# ----------------------------------------------------------------
# 2=PRICE NEW
