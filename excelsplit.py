import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook


# 拆分 Excel 文件的函数
def split_excel_by_column(input_file_path, column_to_filter, output_folder_path):
    # 读取原始 Excel 文件
    df = pd.read_excel(input_file_path)

    # 获取筛选的列中的唯一值
    unique_values = df[column_to_filter].unique()

    for value in unique_values:
        # 对筛选结果进行拆分
        filtered_df = df[df[column_to_filter] == value]

        # 创建一个新的 Excel 文件并将筛选结果保存在其中
        wb = Workbook()
        ws = wb.active

        # Append a customized 1st row
        customized_row = ["Hostname", "BU","Reporting Status","Class","IP","Azure Status","CMDB Status"]  # Customize the values
        ws.append(customized_row)

        for row in filtered_df.itertuples(index=False):
            ws.append(row)

        # Get the current date in YYYY-MM-DD format
        current_date = datetime.now().strftime("%Y-%m-%d")

        # Save the Excel file with the date added to the name
        output_file_name = f"GIS Sonar Non-compliance_{value}_{current_date}.xlsx"
        output_file_path = os.path.join(output_folder_path, output_file_name)
        wb.save(output_file_path)


if __name__ == "__main__":
    input_excel_path = "D:\\pythonProject\\ExcelSplit\\excel\\sonar.xlsx"  # 替换为您的输入 Excel 文件路径
    column_to_filter = "BU"  # 替换为要筛选的列名
    output_folder_path = "D:\\pythonProject\\ExcelSplit\\excel\\output"  # 替换为输出文件夹路径

    os.makedirs(output_folder_path, exist_ok=True)
    split_excel_by_column(input_excel_path, column_to_filter, output_folder_path)
    print("拆分完成。")
