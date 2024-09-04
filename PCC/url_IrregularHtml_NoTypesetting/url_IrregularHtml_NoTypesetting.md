

```chatinput
pip install random re os time lxml pandas requests colorama beautifulsoup4 python-docx openpyxl logging
pip install lxml pandas requests colorama beautifulsoup4 python-docx openpyxl

```

### 函数逻辑说明

1. **`write_to_excel(file_path, dataframe, new_sheet)`**:

   - **功能**: 将数据写入Excel文件。如果文件存在且工作表已存在，则追加数据；否则，创建新文件和工作表。

   - 步骤

     :

     - 检查文件是否存在。如果不存在，则创建新文件并写入数据。
     - 如果文件存在，加载现有工作簿并检查指定的工作表是否存在。
     - 如果工作表存在，读取现有数据并与新数据合并。
     - 使用`pd.ExcelWriter`以追加模式写入数据到工作表中。

2. **主程序**:

   - **功能**: 抓取网页内容，解析数据，并将数据保存到Word文档和Excel文件中。

   - 步骤

     :

     - 设置基础URL和文件路径。
     - 创建存储文件的目录（如果不存在）。
     - 遍历分页URL，获取每一页的数据。
     - 对每个页面链接，提取标题、目录、发布时间和来源，并保存为Word文档。
     - 检查文件大小，删除0KB的文件。
     - 将数据追加到DataFrame中。
     - 最后，将所有数据写入Excel文件。