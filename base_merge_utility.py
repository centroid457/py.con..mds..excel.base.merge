import openpyxl
import sys
import glob

# ================================================================
# 1=DETECT MAIN BASE FILE
file_name_main_base_found = glob.glob("*_admin@madisonorg_*.xlsx")
print("найдены файлы:", file_name_main_base_found)

def found_several_bases():
    print("ОШИБКА: найдено несколько файлов главной базы! уберите лишние! и перезапустите приложение", file=sys.stderr)
    input("нажмите Enter для выхода")

if len(file_name_main_base_found) == 1:
    file_name_main_base = file_name_main_base_found[0]
elif len(file_name_main_base_found) == 2:
    if any((file_name_main_base_found[0].startswith("~$") and file_name_main_base_found[0][2:] == file_name_main_base_found[1],
    file_name_main_base_found[1].startswith("~$") and file_name_main_base_found[1][2:] == file_name_main_base_found[0])):
        print("ОШИБКА: файл главной базы открыт! закройте его и перезапустите приложение!", file=sys.stderr)
        input("нажмите Enter для выхода")
    else:
        found_several_bases()
else:
    found_several_bases()

# ================================================================
# 2=WORK WITH MAIN BASE
wb = openpyxl.load_workbook(file_name_main_base)
ws = wb.active

# ----------------------------------------------------------------
# DETECT COLUMNS IN HEADER
row_title = ws[1]
column_index_from_1 = 1
for cell in row_title:
    cell_value = cell.value
    if cell_value == "Код":
        column_index_code = column_index_from_1
        print(f"колонка Код = {column_index_code}")
    elif cell_value == "Артикул":
        column_index_art = column_index_from_1
        print(f"колонка Артикул = {column_index_art}")
    column_index_from_1 += 1

# ----------------------------------------------------------------
# make DATA COLUMN CODE
column_values_code_iter = ws.iter_cols(min_row=2, min_col=column_index_code, max_col=column_index_code, values_only=True)
column_values_code_list = list(column_values_code_iter)[0]      #[(1, 4, 7)][0] = (1, 4, 7)
print(column_values_code_list)

