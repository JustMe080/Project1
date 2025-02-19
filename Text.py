import pandas as pd
import win32com.client
import numpy as np
import xlwings as xw
import shutil
import datetime

current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S").replace(":", "-")

data1 = pd.read_excel('Addresses.xlsx', sheet_name='Sheet1')

lot = np.array([1, 2, 4, 6, 13, 14, 16, 17, 18, 22, 23, 27, 32, 36, 39, 40, 41, 42, 43, 44, 45], dtype=int)
File = f'Receipt_{current_time}.xlsx' 

original_file = 'C:/project1/template.xlsx'
copy_file = f"C:/project1/{File}" 

shutil.copy(original_file, copy_file)

file_path = f"C:/project1/{File}" 
wb = xw.Book(file_path)

def normalize_text(text):
    if isinstance(text, str):
        return text.strip().lower() 
    return text

for i in range(len(data1)):  
    original_region = str(data1.loc[i, 'Region']).strip()
    original_division = str(data1.loc[i, 'Division']).strip()
    address = str(data1.loc[i, 'Address']).strip()

    region = normalize_text(original_region)
    division = normalize_text(original_division)

    data2 = pd.read_excel('quantity_bases.xlsx', sheet_name=original_region, header=[2, 3])  
    data2.columns = ['_'.join(map(str, col)).strip() for col in data2.columns] 

    data2_normalized = data2.copy()
    data2_normalized.columns = [normalize_text(col) for col in data2.columns]

    data2_temp = pd.read_excel('quantity_bases.xlsx', sheet_name=original_region, header=3)  
    data2_temp_header = [normalize_text(col) for col in data2_temp.columns]

    data2 = data2.drop(columns=["Unnamed: 0_level_0_Unnamed: 0_level_1"], errors="ignore") 
    data2.iloc[:, 2:6] = data2.iloc[:, 2:6].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    counting = ""
    for j in data2_temp_header:
        if division == j:  
            if region == 'ncr':
                counting = f"NATIONAL CAPITAL REGION ({original_region.upper()})_{original_division}"
            elif region == 'car':
                counting = f"CORDILLERA ADMINISTRATIVE REGION ({original_region.upper()})_{original_division}"
            elif region == 'region iv-b':
                counting = f"MIMAROPA_{original_division}"
            else:
                counting = f"{original_region.upper()}_{original_division}"

    if not counting:
        continue  

    counting_normalized = normalize_text(counting)
    if counting_normalized in data2_normalized.columns:
        counting = data2.columns[data2_normalized.columns.get_loc(counting_normalized)]  

    data2 = data2[~data2.apply(lambda row: row.astype(str).str.contains('total', case=False).any(), axis=1)]
    Selected_data = data2[data2[counting] > 1].reset_index(drop=True)
    
    for k in range(len(Selected_data)):
        lot_no = str(Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1']).strip()

        if int(lot_no) in lot:
            data3 = pd.read_excel('Lot_Items.xlsx', sheet_name=lot_no)
            data3_normalized = data3.copy()
            data3_normalized.columns = [normalize_text(col) for col in data3.columns]

            source_sheet = wb.sheets["Sheet1"] 
            source_sheet.api.Copy(After=wb.sheets[-1].api)

            new_sheet = wb.sheets[-1]
            sheet_name = f"{original_region}_{original_division}_{lot_no}"[:30]
            counter = 1
            while sheet_name in [sh.name for sh in wb.sheets]:
                truncated_base = sheet_name[:27]
                sheet_name = f"{truncated_base}_{counter}"
                counter += 1

            new_sheet.name = sheet_name
            new_sheet.range("B15").options(transpose=True).value = data3.iloc[:, 0].tolist() 
            new_sheet.range("C15").options(transpose=True).value = data3.iloc[:, 1].tolist()
            new_sheet.range("B14").value = "Lot No. " + str(Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1']) + " " + Selected_data.loc[k]['QUALIFICATION TITLE/PROGRAM_Unnamed: 2_level_1']
            new_sheet.range("B8").value = f"TESDA {original_region} - {original_division}"
            new_sheet.range("B9").value = address

            # Adjust font size in merged B9:D9
            cell_b9 = new_sheet.range("B9")
            merge_width = cell_b9.merge_area.width
            font_size = cell_b9.api.Font.Size
            text_length = len(str(cell_b9.value)) * 6  # Approximate text width
            while text_length > merge_width and font_size > 8:
                font_size -= 1
                cell_b9.api.Font.Size = font_size
                text_length = len(str(cell_b9.value)) * 6

            for row in range(14, 30):
                cell = new_sheet.range(f"B{row}")
                while cell.column_width < len(str(cell.value)) * 0.8 and cell.api.Font.Size > 8:
                    cell.api.Font.Size -= 1

            wb.app.calculate()
            print("✅ Sheet updated successfully with formatted text!")

wb.save(file_path)
xw.App(visible=False)
wb.close()

print("✅ All sheets updated successfully!")
