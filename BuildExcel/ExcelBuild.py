import xlwings as xw
import re


def convert_to_excel_column(n):
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + n % 26) + result
        n //= 26
    return result


app = xw.App(visible=False, add_book=False)
# 使用xlwings打开当前目录名字为mxr的Excel文件
wbmxr = xw.Book('mxr.xlsx')
# 获取当前名字为包装箱包装信息的sheet
shtInformation = wbmxr.sheets['包装箱包装信息']
# 获取当前名字为inventory的sheet
shtInventory = wbmxr.sheets['inventory']

wbresult = app.books.add()
resultSheet = wbresult.sheets[0]
# 修改resultSheet的名称
resultSheet.name = '发票样本'
# 生成表头
resultSheet.range('A1:Q1').value = ['ITEM NO.', 'FBA标签号', '数量', 'Unit Price', 'Total Price', '箱数', '净重',
                                    '毛重', 'CTN.Size', '中文品名', '英文品名', '材质-中文', '材质-英文', '图片',
                                    'ASIN', 'URL/ 商品链接']
rngInformation = shtInformation.range('M3').value
# 标签号前缀
# 打开文件并读取内容
file_path = "mxr.txt"
with open(file_path, 'r') as file:
    content = file.read()
FBA = content
s = shtInformation.range('A3').value
match = re.search(r'：(.*?)（', s)
if match:
    num = match.group(1)

start_cell = shtInformation.range('M6')
shtInventoryAA = shtInventory.range('A:A')
# 获取导出excel最后一行的下一行
last_row = resultSheet.range('A' + str(resultSheet.cells.last_cell.row)).end('up').row + 1
number = "1"
itemNo = '1'
for _ in range(int(rngInformation)):
    for cell in shtInformation.range(
            convert_to_excel_column(start_cell.column) + str(start_cell.row) + ':' + convert_to_excel_column(
                start_cell.column) + str(6 + int(num) - 1)):
        if cell.value is not None:
            SKU = shtInformation.range('A' + str(cell.row)).value
            asin = shtInformation.range('D' + str(cell.row)).value
            for shtInventoryCell in shtInventoryAA:
                if shtInventoryCell.value == SKU:
                    price = shtInventory.range('F' + str(shtInventoryCell.row)).value
                    break
            data = [itemNo, FBA + itemNo.zfill(6), cell.value, price, int(cell.value) * int(price), 1, None,
                    None, None, None,
                    'keyboard case', '90% EVA+10% 尼龙', 'EVA 90%+Nylon 10%', None, asin,
                    'https://www.amazon.co.jp/dp/' + asin]
            number = str(int(number) + 1)
            resultSheet.range('A' + str(last_row)).value = data
            last_row = last_row + 1
    start_cell = start_cell.offset(0, 1)
    itemNo = str(int(itemNo) + 1)

# 获取 A 列数据
column_a = resultSheet.range('A2:A' + number).value

# 初始化变量
current_value = None
start_row = None
end_row = None

# 遍历 A 列数据
for row_num, value in enumerate(column_a, start=2):
    if value == current_value:
        end_row = row_num
    else:
        if start_row and end_row:
            # 合并相同值的单元格区域
            resultSheet.range(f'A{start_row}:A{end_row}').merge()
            resultSheet.range(f'F{start_row}:F{end_row}').merge()
            resultSheet.range(f'B{start_row}:B{end_row}').merge()
        current_value = value
        start_row = row_num
        end_row = row_num

# 如果最后一组相同值的单元格未合并，则进行合并
if start_row and end_row:
    resultSheet.range(f'A{start_row}:A{end_row}').merge()
    resultSheet.range(f'F{start_row}:F{end_row}').merge()
    resultSheet.range(f'B{start_row}:B{end_row}').merge()

# 添加总和计算
resultSheet.range('C' + str(last_row)).value = sum(resultSheet.range('C2:C' + str(last_row - 1)).value)
resultSheet.range('E' + str(last_row)).value = sum(resultSheet.range('E2:E' + str(last_row - 1)).value)
resultSheet.range('F' + str(last_row)).value = rngInformation
wbresult.save('result.xlsx')
wbresult.close()
wbmxr.close()
app.quit()
