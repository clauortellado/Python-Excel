# Advanced Excel Filtering Across Workbooks Tutorial - Excel Automation with Python Series
# Derrick Sherrill
# https://www.youtube.com/watch?v=Jiyz5RLdV2s

import pandas as pd

xls1= 'ExcelData/2_Workbook_1.xlsx'
xls2= 'ExcelData/2_Workbook_2.xlsx'

df1= pd.read_excel(xls1)
df2= pd.read_excel(xls2)

#Dataframes columns
print(df1.columns)
print(df2.columns)

#Records in boths dataframes
print(df1['Name'].isin(df2['Name']))


# Todos los NAMES del XLS1 que estan en el XLS2
df1_filtered = df1.loc[df1['Name'].isin(df2['Name'])]
#print(df1_filtered)

# Todos los NAMES del XLS1 que NO estan en el XLS2
df2_filtered = df1.loc[~(df1['Name'].isin(df2['Name']))]
#print(df2_filtered)

# Todos los NAMES del XLS1 que estan en el XLS2 AND InterviewScore >4
df3_filtered = df1.loc[df1['Name'].isin(df2['Name']) & (df1['Interview Score']>4)]
#print(df3_filtered)

# Todos los NAMES del XLS1 que estan en el XLS2 OR InterviewScore >4
df4_filtered = df1.loc[df1['Name'].isin(df2['Name']) | (df1['Interview Score']>4)]
#print(df4_filtered)

# Dataframes Merge - All information from DF - Connected by Name - Nan= None
df_all = pd.merge(df1, df2, how='outer', on="Name")
# Complete DF: values
#print(df_all)
# Complete DF: columns
#print(df_all.columns)

# Dataframe All Merge - Filter: 'YR Experience'<5 AND'Group Interview Score'>4
df_all_filter = df_all.loc[(df_all['YR Experience']<5) & (df_all['Group Interview Score']>4)]
print(df_all_filter)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('ExcelData/2_Workbook_Outout.xlsx')

# Convert the dataframe to an XlsxWriter Excel object.
df_all_filter.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
writer.close()