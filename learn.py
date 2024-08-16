import openpyxl.workbook
import pandas as pd
import numpy  as np
import openpyxl
from openpyxl.styles import PatternFill

rows, cols = 100, 100
data=np.random.randint(-100,100,size=(rows,cols))
df=pd.DataFrame(data)
df.to_excel('random_data.xlsx',index=False)

df['Row_mean'] = df.mean(axis=1)
df.loc['Col_mean'] = df.mean(axis=0)

df.to_excel('mean_data.xlsx', index=False)

wb=openpyxl.load_workbook('mean_data.xlsx')
ws=wb.active

for row in range(2,rows+2):
    cell=ws.cell(row=row, column=cols+1)
    if cell.value <0:
        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    else:
        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

for col in range(1, cols+2):
    cell=ws.cell(row=cols+2, column=col)
    if cell.value <0:
       cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    else:
        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

wb.save('mean_data_colored.xlsx')