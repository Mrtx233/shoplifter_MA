

```chatinput
pip install random re os time lxml pandas requests colorama beautifulsoup4 python-docx openpyxl logging
pip install lxml pandas requests colorama beautifulsoup4 python-docx openpyxl

```



### `write_to_excel(file_path, dataframe, new_sheet)`

这个函数的作用是将一个 `pandas` 数据帧 `dataframe` 写入到指定的 Excel 文件 `file_path` 中的工作表 `new_sheet`。

#### 参数说明：

- `file_path`：字符串类型，表示 Excel 文件的路径。
- `dataframe`：`pandas` 数据帧，包含需要写入的数据。
- `new_sheet`：字符串类型，表示要写入的工作表名称。

#### 函数逻辑：

1. 检查指定路径下的文件是否存在。
2. 如果文件不存在，直接将 `dataframe` 写入一个新的 Excel 文件。
3. 如果文件存在，加载 Excel 文件。
4. 检查工作表 `new_sheet` 是否存在。
5. 如果工作表存在，读取该工作表的数据并与 `dataframe` 合并。
6. 将合并后的数据写入工作表 `new_sheet`。
7. 如果工作表不存在，直接将 `dataframe` 写入新的工作表。

### `main(url, target_tag, target_attr=None, output_path='output.docx')`

这个函数的作用是从指定的 URL 中提取 HTML 元素，并将其内容保存到一个 Word 文档中。

#### 参数说明：

- `url`：字符串类型，表示要抓取的网页 URL。
- `target_tag`：字符串类型，表示要提取的 HTML 标签。
- `target_attr`：字典类型，表示 HTML 标签的属性（默认为 `None`）。
- `output_path`：字符串类型，表示保存提取内容的 Word 文档路径（默认为 `'output.docx'`）。

#### 函数逻辑：

1. 发送 HTTP GET 请求获取指定 URL 的响应。
2. 设置响应的编码为 `utf-8`。
3. 检查响应状态码是否为 200（成功）。
4. 使用 `BeautifulSoup` 解析响应的 HTML 内容。
5. 根据 `target_tag` 和 `target_attr` 提取指定的 HTML 元素。
6. 创建一个新的 Word 文档。
7. 遍历提取的 HTML 元素，并调用 `print_tags` 函数将元素内容写入文档。
8. 创建输出目录（如果不存在）。
9. 将文档保存到指定路径。

### `print_tags(tag, level=0, doc=None)`

这个函数的作用是递归地遍历 HTML 元素及其子元素，并将文本内容写入 Word 文档。

#### 参数说明：

- `tag`：`BeautifulSoup` 元素，表示要处理的 HTML 元素。
- `level`：整数类型，表示递归层次（默认为 `0`）。
- `doc`：`Document` 对象，表示要写入的 Word 文档（默认为 `None`）。

#### 函数逻辑：

1. 获取当前元素的所有子元素。
2. 如果没有子元素，获取元素的文本内容，并写入 Word 文档。
3. 如果有子元素，递归调用 `print_tags` 处理每个子元素。
4. 对于每个非空文本段落，创建一个新的段落，并设置字体和字号。

### 主体代码解析

主体代码主要完成网页数据抓取、处理和保存。逻辑如下：

1. 设置基础 URL 和路径等参数。
2. 定义 XPath 路径以提取网页内容。
3. 创建存储文件的目录。
4. 初始化 Excel 文件路径和数据存储。
5. 循环抓取分页数据：
   - 生成分页 URL。
   - 发送 HTTP GET 请求获取页面内容。
   - 解析 HTML 文档，提取页面列表中的链接。
   - 遍历每个链接，提取页面内容并保存到 Word 文档。
   - 将提取的内容保存到 Excel 文件中。

### 具体步骤：

1. **循环分页抓取数据**：通过遍历页面的 URL，抓取分页数据。
2. **请求和解析页面**：发送请求并解析页面内容，提取所需的链接。
3. **处理每个链接**：对于每个链接，再次发送请求并解析页面内容，提取标题、导航、发布时间、来源和正文内容，保存到 Word 文档中。
4. **检查文件大小**：确保文件内容不为空，如果文件大小为 0，则删除文件。
5. **保存数据到 Excel**：将提取的数据保存到指定的 Excel 文件和工作表中。

通过这些步骤，代码能够自动化地从多个网页中抓取内容，并保存到本地的 Word 文档和 Excel 文件中。