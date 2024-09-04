# import pandas as pd
# import os
# import re
# import logging
#
# # 配置日志记录
# logging.basicConfig(
#     level=logging.INFO,
#     format='%(asctime)s - %(levelname)s - %(message)s',
# )
#
# # 用户需要提供的文件夹路径
# excel_dir = rf'E:\WorkingWord\马缕_新闻搞采集(8.12-8.16)\互联网新闻信息稿源单位名单\河北\衡水新闻网'
#
#
# def sanitize_filename(filename):
#     # 移除所有非法字符，包括换行符
#     return re.sub(r'[\\/*?:"<>|\r\n]', "_", filename)
#
#
#
# def print_red(message):
#     logging.error(message)
#
#
# # 遍历文件夹中的所有 Excel 文件
# for filename in os.listdir(excel_dir):
#     if filename.endswith('.xlsx'):
#         excel_path = os.path.join(excel_dir, filename)
#         logging.info(f'处理 Excel 文件: {excel_path}')
#
#         if not os.path.exists(excel_path):
#             logging.error(f'文件不存在: {excel_path}')
#             continue
#
#         # 为每个 Excel 文件创建同名子文件夹
#         folder_name = os.path.splitext(filename)[0]
#         save_dir = os.path.join(excel_dir, folder_name)
#
#         if not os.path.exists(save_dir):
#             os.makedirs(save_dir, exist_ok=True)
#             logging.info(f'创建保存目录: {save_dir}')
#
#         df = pd.read_excel(excel_path)
#
#         logging.info("DataFrame 列信息：")
#         logging.info(df.columns)
#
#         # 遍历 DataFrame 的每一行
#         for index, row in df.iterrows():
#             try:
#                 field1 = row['p1']
#                 field6 = row['p6']
#
#                 if pd.isna(field1) or pd.isna(field6):
#                     print_red(f'跳过文档生成: {field1}（p1标题或p6文本为空）')
#                     continue
#
#                 field2 = row['p2'] if 'p2' in row else ""
#                 field3 = row['p3'] if 'p3' in row else ""
#                 field4 = row['p4'] if 'p4' in row else ""
#                 field5 = row['p5'] if 'p5' in row else ""
#
#                 field_text = f"{field1}\n{field2}\n{field3}\t{field4}\n{field5}\n{field6}"
#                 field_text = field_text.replace('nan', '')
#
#                 doc_name = sanitize_filename(str(field1)) + ".docx"
#                 save_path = os.path.join(save_dir, doc_name)
#
#                 with open(save_path, 'w', encoding="utf-8") as f:
#                     f.write(field_text)
#
#                 logging.info(f'生成文档: {doc_name}')
#             except Exception as e:
#                 logging.error(f'错误: {str(e)}')
#                 continue
#
# logging.info("所有Word文档已生成")
#
#


import pandas as pd
import os
import re
import logging

# 配置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

# 用户需要提供的文件夹路径
excel_dir = rf'E:\WorkingWord\马缕_新闻搞采集(8.12-8.16)\互联网新闻信息稿源单位名单\河北\沧州日报'


def sanitize_filename(filename):
    # 移除所有非法字符，包括换行符
    return re.sub(r'[\\/*?:"<>|\r\n]', "_", filename)


def print_red(message):
    logging.error(message)


def file_size_check(file_path, size_limit_kb):
    # 获取文件大小（以字节为单位）
    size_bytes = os.path.getsize(file_path)
    # 将大小限制从 KB 转换为字节
    size_limit_bytes = size_limit_kb * 1024
    # 如果文件小于限制，删除文件
    if size_bytes < size_limit_bytes:
        os.remove(file_path)
        logging.info(f'文件 {file_path} 小于 {size_limit_kb} KB，已删除')
        return False
    return True


# 遍历文件夹中的所有 Excel 文件
for filename in os.listdir(excel_dir):
    if filename.endswith('.xlsx') and not filename.startswith('~$'):
        excel_path = os.path.join(excel_dir, filename)
        logging.info(f'处理 Excel 文件: {excel_path}')

        if not os.path.exists(excel_path):
            logging.error(f'文件不存在: {excel_path}')
            continue

        # 为每个 Excel 文件创建同名子文件夹
        folder_name = os.path.splitext(filename)[0]
        save_dir = os.path.join(excel_dir, folder_name)

        if not os.path.exists(save_dir):
            os.makedirs(save_dir, exist_ok=True)
            logging.info(f'创建保存目录: {save_dir}')

        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            logging.error(f'无法读取 Excel 文件 {excel_path}: {e}')
            continue

        logging.info("DataFrame 列信息：")
        logging.info(df.columns)

        # 遍历 DataFrame 的每一行
        for index, row in df.iterrows():
            try:
                field1 = row['p1']
                field6 = row['p6']

                if pd.isna(field1) or pd.isna(field6):
                    print_red(f'跳过文档生成: {field1}（p1标题或p6文本为空）')
                    continue

                field2 = row['p2'] if 'p2' in row else ""
                field3 = row['p3'] if 'p3' in row else ""
                field4 = row['p4'] if 'p4' in row else ""
                field5 = row['p5'] if 'p5' in row else ""

                field_text = f"{field1}\n{field2}\n{field3}\t{field4}\n{field5}\n{field6}"
                field_text = field_text.replace('nan', '')

                doc_name = sanitize_filename(str(field1)) + ".docx"
                save_path = os.path.join(save_dir, doc_name)

                with open(save_path, 'w', encoding="utf-8") as f:
                    f.write(field_text)

                logging.info(f'生成文档: {doc_name}')

                # 检查文件大小，并根据条件删除文件
                if not file_size_check(save_path, size_limit_kb=1):
                    continue  # 如果文件已删除，则跳过后续操作

            except Exception as e:
                logging.error(f'错误: {str(e)}')
                continue

logging.info("所有Word文档处理完成")




