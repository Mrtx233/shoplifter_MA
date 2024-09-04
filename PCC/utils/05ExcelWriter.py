import os
import pandas as pd
import logging
import re

# 配置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)


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


def process_excel_files(folder_path):
    # 遍历文件夹中的所有 Excel 文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            excel_path = os.path.join(folder_path, filename)
            logging.info(f'处理 Excel 文件: {excel_path}')

            if not os.path.exists(excel_path):
                logging.error(f'文件不存在: {excel_path}')
                continue

            # 为每个 Excel 文件创建同名子文件夹
            folder_name = os.path.splitext(filename)[0]
            save_dir = os.path.join(folder_path, folder_name)

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

            num_columns = df.shape[1]
            df = df.dropna(subset=[df.columns[num_columns - 1]])

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

            df.to_excel(excel_path, index=False)
            logging.info(f"已处理并保存 Excel 文件: {excel_path}")

    logging.info("所有Word文档处理完成")


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
folder_path = r"E:\WorkingWord\马缕_新闻稿数据(8.19-8.23)\马缕_数据(8.23)\互联网新闻信息稿源单位名单\山西\山西广播电视台"

# 首先处理 Excel 文件，生成 Word 文档
process_excel_files(folder_path)

# 然后将所有处理后的 Excel 文件合并成一个 Excel 文件
create_combined_excel(folder_path)
