import pandas as pd
import os

dictionaries = ['2021', '2022', '2023', '2024-06']
types = ['0002_0001_0001_00000045', '0002_0001_0002_00000006', '0002_0001_0003_00000179']

for dictionary in dictionaries:
    folder_name = dictionary
    for type in types:

        model_name = type + '.csv'
        model_path = os.path.join('csv模板',model_name)
        model = pd.read_csv(model_path, encoding='latin1')

        file_name = type + '.xlsx'
        full_path = os.path.join('审计报告',folder_name, file_name)

        if os.path.exists(full_path):
            df = pd.read_excel(full_path)
            row_number = len(df) - 4
            for i in range(row_number):
                data = df.iloc[i+4, 3]
                model.iloc[i+4, 3] = data
            for col in model.columns:
                if "Unnamed" in col:
                    model = model.rename(columns={col: None})

            csv_file_path = os.path.join('csv版（结果）', dictionary, model_name)
            model.to_csv(csv_file_path, index=False, encoding='latin1')



