

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

2. **`print_tags(tag, level=0, doc=None)`**:

   - **功能**: 将HTML标签及其内容打印到Word文档。

   - 步骤

     :

     - 查找标签的直接子标签。
     - 如果没有子标签，提取文本并添加到Word文档中。
     - 递归处理子标签。

3. **`main(url, target_tag, target_attr=None, output_path='output.docx')`**:

   - **功能**: 抓取网页内容并将内容保存到Word文档。

   - 步骤

     :

     - 发送GET请求获取网页内容。
     - 使用BeautifulSoup解析HTML。
     - 查找目标标签并将内容打印到Word文档中。
     - 保存Word文档到指定路径。

4. **主程序**:

   - **功能**: 遍历分页，抓取每个页面的数据，将内容保存到Word文档，并将数据记录到Excel文件中。

   - 步骤

     :

     - 设置基础URL和文件路径。
     - 创建存储文件的目录（如果不存在）。
     - 遍历分页，获取每一页的数据。
     - 对每个页面链接，提取标题、目录、发布时间和来源，并保存为Word文档。
     - 检查文件大小，删除小于35KB的文件。
     - 将数据追加到DataFrame中。
     - 最后，将所有数据写入Excel文件。