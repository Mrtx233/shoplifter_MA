import os

def get_folder_details(folder_path):
    folder_details = {
        'total_subfolders': 0,
        'subfolder_file_counts': {},
        'total_size_MB': 0,
        'total_files_in_subfolders': 0
    }

    for root, dirs, files in os.walk(folder_path):
        # 计算子文件夹的数量
        if root == folder_path:
            folder_details['total_subfolders'] = len(dirs)

        # 计算每个子文件夹中的文件数量
        for subfolder in dirs:
            subfolder_path = os.path.join(root, subfolder)
            subfolder_file_count = sum([len(files) for _, _, files in os.walk(subfolder_path)])
            folder_details['subfolder_file_counts'][subfolder] = subfolder_file_count
            folder_details['total_files_in_subfolders'] += subfolder_file_count

        # 计算总大小
        for file in files:
            file_path = os.path.join(root, file)
            folder_details['total_size_MB'] += os.path.getsize(file_path) / (1024 * 1024)

    return folder_details

# 示例使用
folder_path = r'E:\WorkingWord\马缕_新闻搞采集(8.5-8.9)\中央和国家机关、群团组织等政务发布平台\中国民航局'  # 替换为你的文件夹路径
details = get_folder_details(folder_path)
print(f"总子文件夹数量: {details['total_subfolders']}")
print("每个子文件夹中的文件数量:")
for subfolder, file_count in details['subfolder_file_counts'].items():
    print(f"  {subfolder}: {file_count} 个文件")
print(f"所有子文件夹中的文件总数: {details['total_files_in_subfolders']} 个文件")
print(f"文件夹总大小: {details['total_size_MB']:.2f} MB")
