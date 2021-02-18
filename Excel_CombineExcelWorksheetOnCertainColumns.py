# Programmatically Combine Excel Worksheets on Certain Columns - Five Minute Python Scripts
# Derrick Sherrill
# https://www.youtube.com/watch?v=MqA0ZNP0EDQ

import os
import pandas as pd

root_path = "/ExcelData"

for file_name in os.listdir(root_path):

    # /path/to/directory/myfile.txt
    full_path_to_file = os.path.join(root_path, file_name)


data_location = "ExcelData/1/"

desired_headings = ["Clave"]
df_total = pd.DataFrame(columns=desired_headings)

for file in os.listdir(data_location):
    df_file = pd.read_excel(data_location + file)
    selected_columns = df_file.loc[:, desired_headings]
    df_total = pd.concat([selected_columns, df_total], ignore_index=True)

df_total