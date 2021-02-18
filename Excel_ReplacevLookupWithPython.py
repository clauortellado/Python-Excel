# Derrick
# https://www.youtube.com/watch?v=cRELNmDpaks+
# Replace Excel VLookUp with Python

import pandas as pd
import numpy as np

xls_Leg= '\\ExcelData\\A_Legajos.xlsx'
xls_pf= 'ExcelData/A_PFs_Legajos.xlsx'
xls_output= 'ExcelData/A_Output.xlsx'

# XLS_1
df_leg= pd.read_excel(xls_Leg)
#Renombro la columna clave
df_leg.rename(columns={'FICHA':'LEGAJO'}, inplace=True)
df_leg.head()

# XLS_2
df_pf= pd.read_excel(xls_pf)
#Renombro la columna clave
df_pf.rename(columns={'FICHA':'LEGAJO'}, inplace=True)
df_pf.head()

# XLS_3
df3= pd.merge(df_pf, df_leg, on='LEGAJO', how='left')
df3.head()

df3_to_excel(xls_output, index=False)