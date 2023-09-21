import openpyxl

# 读取两个xlsx文件
template_wb = openpyxl.load_workbook('template.xlsx')
rasson_wb = openpyxl.load_workbook('work.xlsx')

template_ws = template_wb.active
rasson_ws = rasson_wb.active

# 获取work.xlsx中的必要列索引
sku_col = 1
packing_size_col = 2
weight_col = 3

# 从work.xlsx中创建SKU到详情的映射
sku_details = {}

for row in range(2, rasson_ws.max_row + 1):
    sku = rasson_ws.cell(row=row, column=sku_col).value
    packing_size_raw = rasson_ws.cell(row=row, column=packing_size_col).value
    weight = rasson_ws.cell(row=row, column=weight_col).value

    # 检查 packing_size_raw 是否有效
    if packing_size_raw is None or not isinstance(packing_size_raw, str):
        continue

    if sku not in sku_details:
        sku_details[sku] = []

    try:
        packing_size = packing_size_raw.replace('m', '').split('=')[0].strip()
        lengths = packing_size.split('*')
        if len(lengths) != 3:
            raise ValueError(f"Expected 3 values separated by *, but got {len(lengths)} in {packing_size}")

        length, width, height = map(float, lengths)
        sku_details[sku].append((weight, length*100, width*100, height*100))
    except Exception as e:
        print(f"在第{row}行遇到了问题: {e}")

# 更新template.xlsx中的数据
for row in range(2, template_ws.max_row + 1):
    sku = template_ws.cell(row=row, column=1).value
    details_list = sku_details.get(sku, [])

    if details_list:
        for idx, (weight, length, width, height) in enumerate(details_list, start=1):
            template_ws.cell(row=row, column=1 + (idx-1)*5 + 1).value = weight
            template_ws.cell(row=row, column=1 + (idx-1)*5 + 2).value = length
            template_ws.cell(row=row, column=1 + (idx-1)*5 + 3).value = width
            template_ws.cell(row=row, column=1 + (idx-1)*5 + 4).value = height
            template_ws.cell(row=row, column=1 + (idx-1)*5 + 5).value = 1
    else:
        print(f"没有找到SKU {sku} 对应的数据")

template_wb.save('processed_template.xlsx')

