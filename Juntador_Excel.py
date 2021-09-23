import os
import pandas as pd
import openpyxl

# Remember to put all the xlsx files into your project root
path = os.path.abspath('XLSX path')
arquivos = os.listdir(path)  # It will locate your archieves here
df = pd.DataFrame()
for file in arquivos:
    # It will select only xlsx files
    if file.endswith('.xlsx'):
        df = df.append(pd.read_excel(file), ignore_index=True)

# It will create your merged excel file
df.to_excel('Combined_file.xlsx')
