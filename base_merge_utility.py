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

vendor_dict = {"surgaz": {"file_mask": "*surgaz*.xlsx",
                          "file_found_if_one": None,
                          "article_blank_set": {None, ""},
                          "mask_article_blank_set": {"Артикул", "Материал", "жизни", },
                          "column_article_int": 1,
                          "column_price1_int": 4,
                          "data_article_all_dict": dict(), # большой
                          "data_article_repeated_set": set()},
               }

column_name_base_code = "Код"
column_name_base_art = "Артикул"
column_name_base_price1 = "Цена: Цена продажи"
column_name_base_price2 = "Цена: РРЦ"
column_name_base_price3 = "Закупочная цена"

column_values_code_base_null_set = {column_name_base_code, None, ""}

# ================================================================
# 1=DETECT FILES
# ----------------------------------------------------------------
# 1=FILE - MAIN BASE
print("-"*80)
print(f"START finding files MainBASE")

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
print("-"*80)
print(f"START finding files VENDORS")

for vendor in vendor_dict:
    print("-" * 40)
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
print("*"*80)
print(f"START MainBASE WORK")

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
    if cell_value == column_name_base_code:
        column_index_base_code = column_index_base_from_1
        print(f"Номер колонки [{column_name_base_code}] = [{column_index_base_code}]")
    elif cell_value == column_name_base_art:
        column_index_base_art = column_index_base_from_1
        print(f"Номер колонки [{column_name_base_art}] = [{column_index_base_art}]")
    elif cell_value == column_name_base_price1:
        column_index_base_price1 = column_index_base_from_1
        print(f"Номер колонки [{column_name_base_price1}] = [{column_index_base_price1}]")
    elif cell_value == column_name_base_price2:
        column_index_base_price2 = column_index_base_from_1
        print(f"Номер колонки [{column_name_base_price2}] = [{column_index_base_price2}]")
    elif cell_value == column_name_base_price3:
        column_index_base_price3 = column_index_base_from_1
        print(f"Номер колонки [{column_name_base_price3}] = [{column_index_base_price3}]")
    column_index_base_from_1 += 1

# --------------------------
# 3=load DATA - ColumnCODE
column_values_code_base_all_dict = dict()
column_values_code_base_repeated_set = set()

print("-"*80)
print(f"load data from file: [{file_name_main_base}]")
print("-"*40)
column_values_code_base_iter = ws_base.iter_cols(min_col=column_index_base_code, max_col=column_index_base_code)
for column_tuple in column_values_code_base_iter:
    for cell_obj in column_tuple:
        cell_value = cell_obj.value
        if cell_value not in column_values_code_base_all_dict:
            # print("+", cell_value)
            column_values_code_base_all_dict.update({cell_value: {"cell_obj_list": [cell_obj, ],
                                                                  "marker": None,
                                                                  "price1": None,
                                                                  "price2": None,
                                                                  "price3": None,
                                                                  }})
            """
            MARKER
            0=NULL ARTICLE
            1=INFO LINE
            13=INCORRECT DATA=exists different price for one article
            100=OK
            """
        else:
            print(f'found repeated value: [{cell_value}]')
            column_values_code_base_all_dict[cell_value]["cell_obj_list"].append(cell_obj)
            column_values_code_base_repeated_set.update({cell_value})

# --------------------------
# 4=print loadRESULTS
count_column_values_code_base_all_dict = len(column_values_code_base_all_dict)
count_column_values_code_base_repeated_set = len(column_values_code_base_repeated_set)

print("-"*40)
# print("from file:", file_name_main_base)
print("count_column_values_code_base_all_dict:", count_column_values_code_base_all_dict)
print("count_column_values_code_base_repeated_set:", count_column_values_code_base_repeated_set)

print("column_values_code_base_repeated_set:", column_values_code_base_repeated_set)
print("-"*80)

# ================================================================
# 3=VENDOR FILES=WORK
print("*"*80)
print(f"START VENDOR WORK")

for vendor in vendor_dict:
    vendor_data_dict = vendor_dict[vendor]

    file_name_vendor = vendor_data_dict["file_found_if_one"]
    if file_name_vendor is None:
        continue

    # --------------------------
    # 1=file INIT
    wb_vendor = openpyxl.load_workbook(file_name_vendor)
    ws_vendor = wb_vendor.active

    # --------------------------
    # 2=DETECT COLUMNS IN HEADER
    column_index_vendor_article = vendor_data_dict["column_article_int"]
    column_index_vendor_price1 = vendor_data_dict["column_price1_int"]

    # --------------------------
    # 3=load DATA - ColumnCODE
    column_values_vendor_article_all_dict = vendor_data_dict["data_article_all_dict"]
    column_values_vendor_article_repeated_set = vendor_data_dict["data_article_repeated_set"]

    print("-"*80)
    print(f"load data from file: [{file_name_vendor}]")
    print("-"*40)
    column_values_vendor_article_iter = ws_vendor.iter_cols(min_col=column_index_vendor_article, max_col=column_index_vendor_article)
    for column_tuple in column_values_vendor_article_iter:
        for cell_obj in column_tuple:
            cell_value = cell_obj.value

            if cell_value not in column_values_vendor_article_all_dict:
                # print("+", cell_value)
                column_values_vendor_article_all_dict.update({cell_value: {"cell_obj_list": [cell_obj, ],
                                                                           "marker": None,
                                                                           "price1": None,
                                                                           "price2": None,
                                                                           "price3": None,
                                                                           }})

                cell_value_dict = column_values_vendor_article_all_dict[cell_value]
                cell_obj_price1 = ws_vendor.cell(row=cell_obj.row, column=column_index_vendor_price1).value

                if cell_value in vendor_data_dict["article_blank_set"]:
                    cell_value_dict["marker"] = 0   # NULL ARTICLE
                elif all([mask in cell_value for mask in vendor_data_dict["mask_article_blank_set"]]):
                    cell_value_dict["marker"] = 1   # INFO LINE

                elif cell_obj_price1 is not None:
                    cell_value_dict["price1"] = cell_obj_price1
                    cell_value_dict["marker"] = 100   # OK
                elif cell_obj_price1 is None and column_values_code_base_all_dict.get(cell_value, None) is None:
                    cell_value_dict["marker"] = 1   # INFO LINE

            else:
                cell_value_dict = column_values_vendor_article_all_dict[cell_value]
                cell_obj_list = cell_value_dict["cell_obj_list"]
                cell_obj_list.append(cell_obj)

                # check indent price! may be it was several articles but indent price! it is OK!
                cell_obj_price1_last = ws_vendor.cell(row=cell_obj_list[-1].row, column=column_index_vendor_price1).value
                cell_obj_price1_prev = ws_vendor.cell(row=cell_obj_list[-2].row, column=column_index_vendor_price1).value
                if cell_obj_price1_last != cell_obj_price1_prev:
                    column_values_vendor_article_repeated_set.update({cell_value})
                    count_repeated = len(column_values_vendor_article_all_dict[cell_value]["cell_obj_list"])
                    print(f'found repeated value: [{cell_value}] \tby [{count_repeated}]times')
                    cell_value_dict["marker"] = 13   # INCORRECT DATA=exists different price for one article

    # --------------------------
    # 4=print loadRESULTS
    count_column_values_vendor_article_all_dict = len(column_values_vendor_article_all_dict)
    count_column_values_vendor_article_repeated_set = len(column_values_vendor_article_repeated_set)

    print("-"*40)
    # print("from file:", file_name_main_base)
    print("count_column_values_vendor_article_all_dict:", count_column_values_vendor_article_all_dict)
    print("count_column_values_vendor_article_repeated_set:", count_column_values_vendor_article_repeated_set)

    print("column_values_vendor_article_repeated_set:", column_values_vendor_article_repeated_set)
    print("*"*80)

    for cell_velue in column_values_vendor_article_all_dict:
        data_dict = column_values_vendor_article_all_dict[cell_velue]
        article_value = cell_velue
        article_price1 = data_dict["price1"]
        article_mark = data_dict["marker"]
        print(f"{article_mark}=[{article_value}]={article_price1}")
