from playwright.sync_api import sync_playwright
import xlwings as xw
import os


def delete_sheet_by_name(workbook, sheet_name):
    """
    删除指定名称的工作表，如果存在。
    """
    try:
        sheet_to_delete = workbook.sheets[sheet_name]
        if len(workbook.sheets) > 1:
            sheet_to_delete.delete()
            print(f"Sheet '{sheet_name}' has been deleted.")
        else:
            print(f"Cannot delete '{sheet_name}' because it is the only sheet in the workbook.")
    except Exception as e:
        print(f"Sheet '{sheet_name}' not found or cannot be deleted: {e}")


# 获取当前目录
current_directory = os.getcwd()
if not os.path.exists('temp'):
    os.mkdir('temp')
# 定义保存路径为当前目录下的截图.xlsx
excel_file_path = os.path.join(current_directory, "截图.xlsx")
txt_file_path = os.path.join(current_directory, "parameters.txt")
if os.path.exists(txt_file_path):
    # 创建一个字典用于存储读取的参数
    parameters = {}
    # 打开并读取txt文件
    with open(txt_file_path, 'r') as file:
        for line in file:
            # 去掉每行的换行符并检查是否为空行
            line = line.strip()
            if line:
                # 解析键值对
                key, value = line.split('=')
                parameters[key.strip()] = int(value.strip())
    # 你可以通过字典访问各个参数
    w1 = parameters['w1']
    h1 = parameters['h1']
    w2 = parameters['w2']
    h2 = parameters['h2']
else:
    w1 = 1300
    h1 = 2000
    w2 = 1300
    h2 = 1800
# 创建一个新的exce
app = xw.App(visible=False)
wb = app.books.add()
wb.save(excel_file_path)
new_wb = app.books.open(excel_file_path)

with sync_playwright() as p:
    # 启动浏览器
    browser = p.chromium.launch(headless=False, proxy={"server": "http://127.0.0.1:7897"})
    page = browser.new_page()

    excel = xw.Book('result.xlsx')
    sheet = excel.sheets['发票样本']

    # 循环P2到P列最后的数据
    for i in range(2, sheet.range('P1').end('down').row + 1):
        url = sheet.range('P' + str(i)).value
        asin = sheet.range('O' + str(i)).value

        # 打开网页
        page.goto("https://sellercentral-japan.amazon.com/revcalpublic?ref=RC1&lang=ja-JP")

        # 点击按钮
        selector = '.spacing-top-small > button'
        page.wait_for_selector(selector)
        element = page.locator(selector)
        element.click()

        # 填写输入框
        page.fill('input[id="katal-id-4"]', asin)
        page.click('kat-button[label="検索"]')

        # 等待 close 图标加载
        close_icon_selector = '.program-card-header-icon kat-icon[name="close"]'
        page.wait_for_selector(close_icon_selector)

        # 点击
        close_icon = page.locator(close_icon_selector)
        close_icon.nth(1).click()

        # 截图
        screen_1 = page.screenshot(clip={"x": 0, "y": 0, "width": w1, "height": h1}, full_page=True)

        # 打开网页
        page.goto(url)
        page.fill('input[id="twotabsearchtextbox"]', asin)
        page.mouse.click(100, 300);
        # page.click('span[id="a-autoid-36-announce"]')
        # selector = '.a-button-input'
        # page.wait_for_selector(selector)
        # element = page.locator(selector).nth(1)
        # element.click()
        # 截图
        screen_2 = page.screenshot(clip={"x": 0, "y": 0, "width": w2, "height": h2}, full_page=True)

        # 将截图保存到当前目录
        screenshot_1_path = os.path.join(current_directory + '\\temp', f"{asin}_screenshot1.png")
        screenshot_2_path = os.path.join(current_directory + '\\temp', f"{asin}_screenshot2.png")

        with open(screenshot_1_path, "wb") as f:
            f.write(screen_2)

        with open(screenshot_2_path, "wb") as f:
            f.write(screen_1)

        # 添加图片到Excel
        sheet1 = new_wb.sheets.add(name=asin + '_1', after=new_wb.sheets[-1])
        sheet2 = new_wb.sheets.add(name=asin + '_2', after=new_wb.sheets[-1])
        sheet1.pictures.add(screenshot_1_path, name='Screenshot2', update=True, left=0, top=0)
        sheet2.pictures.add(screenshot_2_path, name='Screenshot1', update=True, left=0, top=0)

    # 关闭浏览器
    browser.close()

delete_sheet_by_name(wb, "Sheet1")
# 保存并关闭Excel
new_wb.save(excel_file_path)
new_wb.close()
app.quit()
