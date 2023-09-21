import openpyxl

# 读取两个xlsx文件
template_wb = openpyxl.load_workbook('template.xlsx')
rasson_wb = openpyxl.load_workbook('Rasson包装.xlsx')

template_ws = template_wb.active
rasson_ws = rasson_wb.active

# 获取Rasson包装.xlsx中的必要列索引
sku_col = 2
packing_size_col = 6
weight_col = 8

# 从Rasson包装.xlsx中创建SKU到详情的映射
sku_details = {}

for row in range(2, rasson_ws.max_row + 1):
    sku = rasson_ws.cell(row=row, column=sku_col).value
    packing_size_raw = rasson_ws.cell(row=row, column=packing_size_col).value

    # 进行检查
    if not packing_size_raw or '*' not in packing_size_raw:
        print(f"在第{row}行遇到了问题: {packing_size_raw}")
        continue

    try:
        # 删除m并根据=分割字符串
        packing_size = packing_size_raw.replace('m', '').split('=')[0]
        length, width, height = map(float, packing_size.split('*'))
        weight = rasson_ws.cell(row=row, column=weight_col).value
        if weight is None:
            weight = 0
        sku_details[sku] = (weight, length*100, width*100, height*100)
    except Exception as e:
        print(f"在第{row}行遇到了问题: {e}")

# 更新template.xlsx中的数据
for row in range(2, template_ws.max_row + 1):
    sku = template_ws.cell(row=row, column=1).value
    if sku in sku_details:
        weight, length, width, height = sku_details[sku]
        template_ws.cell(row=row, column=3).value = weight
        template_ws.cell(row=row, column=4).value = length
        template_ws.cell(row=row, column=5).value = width
        template_ws.cell(row=row, column=6).value = height
        template_ws.cell(row=row, column=7).value = 1

template_wb.save('processed_template.xlsx')

