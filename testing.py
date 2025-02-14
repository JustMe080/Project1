import pandas as pd
import win32com.client

data = pd.read_excel('Addresses.xlsx', sheet_name='Sheet1')

print(data.columns)
selected_Data = data[data['Region'] == 'Region I']
data_div = selected_Data[selected_Data['Division'] == 'Ilocos Norte']

for i in data['Region']:
    data2 = pd.read_excel('quantity_bases.xlsx', sheet_name = i, header=[2,3])
    data2.columns = ['_'.join(map(str, col)).strip() for col in data2.columns] 
    data2 = data2.drop(columns=["Unnamed: 0_level_0_Unnamed: 0_level_1"], errors="ignore")
    data2.iloc[:, 2:6] = data2.iloc[:, 2:6].fillna(0).astype(int)
    print(data2)
    selected = data[data['Region']==i]
    for j in selected['Division']:
        counting = str(i.upper() +'_'+ j)
        counted_data = data2.loc[data2[counting] >= 1]
        print(counted_data)

print(selected_Data)
print(data_div)



counting = str(i.upper() +'_'+ j)
print(counting)
selected1 = data2.loc[data2[counting] >= 1]
print(selected1)
lotno = selected1.iloc[0]['LOT NO._Unnamed: 1_level_1']
data3 = pd.read_excel('Lot_Items.xlsx', sheet_name= str(lotno))
print(data3)
print((data3.iloc[:,-1]*3).sum())

    

