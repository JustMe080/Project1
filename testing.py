import pandas as pd
import win32com.client
import numpy as np
import xlwings as xw

data = pd.read_excel('Addresses.xlsx', sheet_name='Sheet1')

print(data.columns)
lot = np.array([1, 2, 4, 6, 13, 14, 16, 17, 18, 22, 23, 27, 32, 36, 39, 40, 41, 42, 43, 44, 45], dtype=int)
counting = ""
data_startloop = len(data) - 1

for i in range(data_startloop):
    data2 = pd.read_excel('quantity_bases.xlsx', sheet_name = data.loc[i]['Region'], header=[2,3])
    data2.columns = ['_'.join(map(str, col)).strip() for col in data2.columns]
    data2_temp = pd.read_excel('quantity_bases.xlsx', sheet_name = data.loc[i]['Region'], header=3)
    data2_temp_header = data2_temp.columns
    data2 = data2.drop(columns=["Unnamed: 0_level_0_Unnamed: 0_level_1"], errors="ignore") 
    data2.iloc[:, 2:6] = data2.iloc[:, 2:6].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    for j in data2_temp_header:
        if data.loc[i]['Division'] == j:
            if data.loc[i]['Region'] == 'NCR':
                counting = str("NATIONAL CAPITAL REGION ("+data.loc[i]['Region'].upper()+")_"+data.loc[i]['Division'])
            elif data.loc[i]['Region'] == 'CAR':
                counting = str('CORDILLERA ADMINISTRATIVE REGION (' + data.loc[i]['Region'].upper() +')_'+data.loc[i]['Division'])
            elif data.loc[i]['Region'] == 'Region IV-B':
                counting = str('MIMAROPA_' + data.loc[i]['Division'])
            else:
                counting = str(data.loc[i]['Region'].upper() +'_'+ data.loc[i]['Division'])
                
    Selected_data = data2[data2[counting] > 1]
    Selected_data = Selected_data.reset_index(drop=True)
    for k in range(len(Selected_data)-1):
        for l in lot:
            if Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1'] == l:
                data3 = pd.read_excel('Lot_Items.xlsx', sheet_name = str(Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1']))
                print(f"Tesda {data.loc[i]['Region']} - {data.loc[i]['Division']}")
                print(f"Address: {data.loc[i]['Address']}")
                print(f"Quantity: {Selected_data.loc[k][counting]}")
                print(f"items: {data3.iloc[:,0]}")
                file_path = "C:/project1/template.xlsx"
                wb = xw.Book(file_path)
                source_sheet = wb.sheets["Sheet1_Copied"]
                source_sheet.api.Copy(After=wb.sheets[-1].api)
                new_sheet = wb.sheets[-1]
                new_sheet.name = "Sheet1_"+ str(i)
                new_sheet.range("B15").value = data3.iloc[:,0] 
                new_sheet.range("B8").value = f"Tesda {data.loc[i]['Region']} - {data.loc[i]['Division']}"
                new_sheet.range("B9").value = data.loc[i]['Address']
                new_sheet.range("A14").value = Selected_data.loc[k][counting]
                wb.save(file_path)
                wb.close()
                print("✅ Sheet duplicated successfully with images!")
    










