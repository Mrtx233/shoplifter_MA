# import os
# import pandas as pd
#
#
# def create_combined_excel(folder_path):
#     # 获取文件夹中的所有 Excel 文件
#     excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]
#
#     # 获取文件夹的名称
#     folder_name = os.path.basename(folder_path.rstrip("\\/"))
#
#     # 输出文件路径
#     output_file = os.path.join(folder_path, f"{folder_name}.xlsx")
#
#     # 创建一个新的 Excel 文件
#     with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#         for file in excel_files:
#             file_path = os.path.join(folder_path, file)
#
#             # 读取 Excel 文件
#             df = pd.read_excel(file_path)
#
#             # 获取文件名（去除扩展名）
#             sheet_name = os.path.splitext(file)[0]
#
#             # 将数据写入新的 Excel 文件中的对应工作表
#             df.to_excel(writer, sheet_name=sheet_name, index=False)
#
#             print(f"Added {sheet_name} to {output_file} successfully.")
#
#
# # 直接在代码中指定文件夹路径
# folder_path = r"E:\WorkingWord\马缕_新闻搞采集(8.12-8.16)\互联网新闻信息稿源单位名单\河北\衡水新闻网"
# create_combined_excel(folder_path)


import os
import pandas as pd
import logging
import re

# 配置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)


def create_combined_excel(folder_path):
    # 获取文件夹中的所有 Excel 文件
    excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx') and not file.startswith('~$')]

    # 获取文件夹的名称
    folder_name = os.path.basename(folder_path.rstrip("\\/"))

    # 输出文件路径
    output_file = os.path.join(folder_path, f"{folder_name}.xlsx")

    # 创建一个新的 Excel 文件
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            try:
                # 读取 Excel 文件
                df = pd.read_excel(file_path)

                # 删除第六列（如果存在）
                if df.shape[1] >= 6:
                    df.drop(df.columns[5], axis=1, inplace=True)

                # 修改表头
                # 确保表头的列数和现有列数匹配
                new_headers = ['标题', '地址', '目录', '发布时间', '来源']
                if len(new_headers) == df.shape[1]:
                    df.columns = new_headers
                else:
                    logging.warning(
                        f"Column count mismatch for file {file}. Expected {len(new_headers)} headers but found {df.shape[1]} columns.")

                # 获取文件名（去除扩展名），并检查是否符合 Excel 工作表名称规则
                sheet_name = os.path.splitext(file)[0]
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                if not sheet_name.isidentifier():
                    sheet_name = re.sub(r'[\\/*?:"<>|]', '_', sheet_name)

                # 将数据写入新的 Excel 文件中的对应工作表
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                logging.info(f"Added {sheet_name} to {output_file} successfully.")
            except Exception as e:
                logging.error(f"Error processing file {file}: {e}")

    logging.info(f"All files have been combined into {output_file}")


# 直接在代码中指定文件夹路径
folder_path = r"E:\WorkingWord\马缕_新闻搞采集(8.12-8.16)\互联网新闻信息稿源单位名单\河北\河北日报"
create_combined_excel(folder_path)

