# Sku-details-fetch
fetch the sku details, and sort as the standard format show in template。xlsx

# SKU Data Processor

该项目提供了一个Python脚本，它读取两个Excel文件中的数据，然后处理和整合这些数据，并将结果写入一个新的Excel文件。

## 功能

- 从`mlds-list.xlsx`读取SKU、包装尺寸和重量数据。
- 将尺寸从毫米转换为厘米。
- 避免将重复的SKU数据写入目标Excel文件。
- 将处理后的数据写入`template.xlsx`的`SKUs`和`Total`工作表。

## 使用方法

1. 确保已经安装了`openpyxl`库。
2. 将要处理的数据放入`mlds-list.xlsx`。
3. 运行脚本。
4. 检查`processed_template1.xlsx`以查看处理后的数据。

## 代码结构

- `convert_mm_to_cm(value)`: 将给定的毫米值转换为厘米。
- 主逻辑:
  - 从`mlds-list.xlsx`中读取数据。
  - 处理数据。
  - 将数据写入`template.xlsx`。

## 出现的问题

如果在`mlds-list.xlsx`中有无效或格式不正确的数据，脚本可能会抛出错误。
