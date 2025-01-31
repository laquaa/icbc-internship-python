import pandas as pd
import os

dictionaries = ['2021', '2022', '2023', '2024-06']
types = ['0002_0001_0001_00000045', '0002_0001_0002_00000006', '0002_0001_0003_00000179']

for dictionary in dictionaries:
    for type in types:
        folder_name = dictionary
        file_name = type + '.xlsx'
        full_path = os.path.join(folder_name, file_name)
        if os.path.exists(full_path):
            df = pd.read_excel(full_path)
            csv_file_name = type + '.csv'
            csv_folder_name = dictionary + '_csv'
            csv_file_path = os.path.join(csv_folder_name, csv_file_name)
            df.to_csv(csv_file_path, index=False)