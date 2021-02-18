# 5 Minute Python Scripts - Automate Multiple Sheet Excel Reporting - Full Code Along Walkthrough
# Derrick Sherrill
# https://youtu.be/ZRwiMcGUf-Y

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

xls1 = 'ExcelData/1_shift-data.xlsx'
xls2 = 'ExcelData/1_shift-third-data.xlsx'

# Input Data
df_sheet_1 = pd.read_excel(xls1, sheet_name='first')
df_sheet_2 = pd.read_excel(xls1, sheet_name='second')
df_sheet_3 = pd.read_excel(xls2)

# df_sheet_1 
# df_sheet_1['Product']


# Combining all data
df_all = pd.concat([df_sheet_1, df_sheet_2, df_sheet_3])

#df_all


# Calculations: average(promedio)
pivot = df_all.groupby(['Shift']).mean()
shift_productivity = pivot.loc[:,"Production Run Time (Min)":"Products Produced (Units)"]

shift_productivity
shift_productivity.plot(kind='bar')
#shift_productivity.show()

# Output Data
df_all.to_excel("ExcelData/1_shift-output.xlsx", sheet_name='data')
shift_productivity.to_excel("ExcelData/1_shift-output.xlsx", sheet_name='average')