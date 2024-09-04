import os
import pandas as pd


def process_excel_files(folder_path):
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        num_columns = df.shape[1]
        df = df.dropna(subset=[df.columns[num_columns - 1]])
        df.to_excel(file_path, index=False)
        print(f"Processed {file} successfully.")


folder_path = r"E:\WorkingWord\马缕_新闻搞采集(8.19-8.23)\互联网新闻信息稿源单位名单\山西\山西新闻网"
process_excel_files(folder_path)
