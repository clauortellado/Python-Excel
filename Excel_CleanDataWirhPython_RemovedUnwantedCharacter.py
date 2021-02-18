# PYTHON - DERRICK SHERRILL
# Clean Excel Data with Python - Removed Unwanted Character
# https://www.youtube.com/watch?v=HU0re8UJViM


import pandas as pd

xls='ExcelData/Excel_ToClean.xlsx'
df= pd.read_excel(xls)

df.head()

# replace Everything is Not a Number and not a Letter with an Empty String
# in Column NAME
#df['NAME'] = df['NAME'].str.replace(r'\W',"")
#print(df)

# replace Everything is Not a Number and not a Letter with an Empty String
# loop for All Data on the All Columns

for column in df.columns:
    df[column]= df[column].str.replace(r'\W',"")
    
df.to_excel("ExcelData/Excel_ToClean2.xlsx