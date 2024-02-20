from openpyxl import load_workbook

MAX_ROWS_TO_PROCESS = 11
PRODUCT_ID_COLUMN = "L"
PRODUCT_ID_COLUMN2 = "J"

def remove_matching_rows(file1_path, file2_path, start_row=5):
    wb1 = load_workbook(file1_path)
    sheet1 = wb1.active

    wb2 = load_workbook(file2_path, keep_vba=True)
    sheet2 = wb2.active

    values_to_check = [sheet1[f"{PRODUCT_ID_COLUMN}{row}"].value for row in range(start_row, sheet1.max_row + 1) if sheet1[f"{PRODUCT_ID_COLUMN}{row}"].value is not None]

    rows_to_delete = []

    for row in range(1, sheet2.max_row + 1):
        cell_value = sheet2[f"{PRODUCT_ID_COLUMN2}{row}"].value
        if cell_value in values_to_check:
            rows_to_delete.append(row)

    rows_to_delete.sort(reverse=True)

    for row in rows_to_delete:
        sheet2.delete_rows(row)

    wb2.save(file2_path)
    print(f"Rows removed and {file2_path} saved.")

file1_path = "allegro.xlsx"
file2_path = "allegro2.xlsm"
remove_matching_rows(file1_path, file2_path)