from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.common.exceptions
import selenium.common
from selenium.webdriver.common.by import By
import requests
from PIL import Image
from io import BytesIO
import time
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from bs4 import BeautifulSoup
from pyzbar.pyzbar import decode
import pyqrcode
import threading
import os
from msedge.selenium_tools import EdgeOptions
from msedge.selenium_tools import Edge

web_browser = input("请输入浏览器类型：[0] Chrome [1] Windows Edge:\n")

if web_browser == '0':
    # 配置Chrome WebDriver
    chromedriver_path = "./chromedriver-win64/chromedriver.exe"
    # service = Service(chromedriver_path)
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    #driver = webdriver.Chrome(service = ChromeService(ChromeDriverManager().install()), options=options)
    driver = webdriver.Chrome(service=ChromeService(chromedriver_path), options=options)
elif web_browser == '1':
    Edgedriver_path = "./edgedriver_win64/msedgedriver.exe"
    edge_options = EdgeOptions()
    edge_options.use_chromium = True
    # 设置无界面模式，也可以添加其它设置
    edge_options.add_argument('headless')
    driver = Edge(options=edge_options, executable_path=Edgedriver_path)
else:
    print("输入错误")
    os._exit(0)

# 打开目标页面
target_url = "https://jwgl.ustb.edu.cn/"
driver.get(target_url)

# 找到目标图片
# image_element = driver.find_element(By.XPATH,"//div[@id='bjkjdx_qrcode_view']//img[@id='img']")
iframe = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//iframe"))
)
driver.switch_to.frame(iframe)

image_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "qrimg"))
)
image_url = image_element.get_attribute("src")

driver.switch_to.default_content()

# 在后台线程中显示图片
response = requests.get(image_url)
img = Image.open(BytesIO(response.content))

decodeQR = decode(img)
print(decodeQR[0].data.decode("utf-8"))
qr = pyqrcode.create(decodeQR[0].data.decode("utf-8"))
print(qr.terminal(module_color="black", background="white"))
print("请扫码>>>>")


def thead_print(stop_event):
    for i in range(60):
        print("\r", "等待扫码>>>>  {}s/60s".format(60 - i), end="", flush=True)
        time.sleep(1)
        if stop_event.is_set():
            return
    print("扫码超时 >> 请重新运行程序")
    os._exit(0)


stop_event = threading.Event()
thread = threading.Thread(target=thead_print, args=(stop_event,))
thread.start()


def thread_get_alert():
    for i in range(16):
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except selenium.common.exceptions.NoAlertPresentException:
            pass
        time.sleep(0.5)


thread1 = threading.Thread(target=thread_get_alert)
thread1.start()
# WebDriverWait (driver, timeout=True, poll_frequency=10).until (EC.presence_of_element_located ((By.CLASS_NAME, "topic-item")))

click_1 = WebDriverWait(driver, 100).until(
    EC.presence_of_element_located((By.XPATH, "//img[@data-src='nav2']"))
)
click_1.click()
print("登录成功>>>>")
stop_event.set()

click_2 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//li[@data-code='NEW_XSD_XJCJ_WDCJ_KCCJCX']")
    )
)
click_2.click()

iframe = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[@id='mainIframe']"))
)
driver.switch_to.frame(iframe)

iframe2 = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[@id='cjcx_query_frm']"))
)
driver.switch_to.frame(iframe2)
click_iframe = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@onclick='queryKscj()']"))
)
click_iframe.click()

driver.switch_to.default_content()
driver.switch_to.frame(iframe)

iframe = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[@id='cjcx_list_frm']"))
)
driver.switch_to.frame(iframe)

html_srouce = driver.page_source

soup = BeautifulSoup(html_srouce, "lxml")

# 找到表格元素
table = soup.find("table", id="dataList")

# 初始化一个空列表来存储表格数据
table_data = []

# 找到所有的行（<tr> 标签），跳过第一行（表头）
rows = table.find_all("tr")

# 遍历每一行，提取单元格内容
for row in rows:
    # 找到行中的所有单元格（<td> 标签）
    cells = row.find_all("td")
    # 提取每个单元格的文本内容，并添加到列表中
    row_data = [cell.text for cell in cells]
    table_data.append(row_data)

# average_table = [row for row in table_data if ((row[4] != "0") and (row[8] != "专业选修") and (row[8] != "素质拓展") and (row[8] != "专业拓展"))]

average_table = []
credits = 0
scores = 0
for row in table_data:
    if row != []:
        if (
            (row[5] != "0")             # 学分为0
            and (row[4] != '0')         # 成绩为0
            and (row[9] != "专业选修")
            and (row[9] != "素质拓展")
            and (row[9] != "专业拓展")
            and (row[11] != "辅修")
            and (row[9] != "国防公益")

        ):
            credits = credits + float(row[5])
            scores = scores + float(row[5]) * float(row[4])
            average_table.append(row)


df = pd.DataFrame(average_table[0:])

excel_file = "data.xlsx"
df.to_excel(excel_file, index=False, engine="openpyxl")

print("已生成>>")
print("Your Average Scores is: >>  ", str(scores / credits), "  <<")


driver.quit()
