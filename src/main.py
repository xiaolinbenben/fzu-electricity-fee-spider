import time
from openpyxl import Workbook
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# 设置 ChromeDriver 的路径
def fee(driver, b_num, d_num):
    driver.get(
        'https://cloud.life.ccb.com/index_u.jhtml?param=4C4A647F371B43F11DBA91BA7799EE58C59D79B9E834A07DA96C44E250D2E7C78781BBF44E457D35A237FF682F4BD607D1AABB8CEA42084E4CE8C5ADAC6473BDC47749CBBD30B2A06C82607E4D2CF04BD9ADC269DEFD0FE7D61058311BDAABD73F7CAFA0C143779B602AF2A9BE9743569C714A5ABD64A6560B60B92AFEA11F2C0D75995186952173'
    )

    dorm_element = driver.find_element(By.XPATH, "//*[@id='inputTab']/tbody/tr[1]/td[2]/span[2]")
    dorm_element.click()

    # 选择楼栋
    # 等待<ul>元素加载完成
    wait = WebDriverWait(driver, 10)
    ul_element = wait.until(EC.presence_of_element_located((By.ID, "letter#")))

    # 定义要查找的楼号
    target_building = str(b_num) + "号楼"

    # 查找所有<li>元素
    li_elements = ul_element.find_elements(By.CLASS_NAME, "letter_content")

    # 遍历<li>元素，并点击匹配的楼号
    for li in li_elements:
        if li.text == target_building:
            driver.execute_script("arguments[0].scrollIntoView(true);", li)
            li.click()
            # print(f"已点击{target_building}")
            break
    else:
        print(f"{target_building} 不存在")

    # 定位到输入框并清空
    ele = driver.find_element(By.ID, '0000333441')
    ele.clear()  # 清空输入框
    ele.send_keys(d_num)  # 输入新的宿舍号

    # 点击查询按钮
    dorm_element = driver.find_element(By.XPATH, "//*[@id='chaxunbtn']")
    dorm_element.click()
    time.sleep(3)
    # 点击确认按钮
    try:
        confirm_button = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "singleBtn") and text()="确定"]'))
        )
        confirm_button.click()
    except Exception as e:
        pass

    # 直接去表2保存所有数据
    try:
        driver.find_element(By.ID, "tab2")
        lishidingdan = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/span[2]")
        lishidingdan.click()
        # 定位<ul>元素
        ul_element = driver.find_element(By.ID, "tab2")
        # 定位<ul>内的所有<li>元素
        li_elements = ul_element.find_elements(By.TAG_NAME, "li")
        # 存储所有数据的列表
        data_list = []

        # 遍历所有<li>元素
        for li in li_elements:
            # 提取日期信息
            date = li.find_element(By.CLASS_NAME, "font1").text

            # 提取金额信息
            amount = li.find_element(By.CSS_SELECTOR, "input.font3").get_attribute("value")

            # 存储数据到字典
            each_data = {"b_num": b_num, "d_num": d_num, "date": date, "amount": amount}
            data_list.append(each_data)


        return data_list
    except:
        return [{"b_num": b_num, "d_num": d_num, "date": "未查询到该宿舍数据",
                "amount": "未查询到该宿舍数据"}]

def run_one_building(b_num):
    #新建一个xlsx文档，存在的话覆盖
    # 创建一个工作簿对象
    workbook = Workbook()

    # 获取默认的工作表
    sheet = workbook.active
    sheet.title = "Sheet1"

    # 保存工作簿
    file_name = '{}号楼.xlsx'.format(b_num)
    workbook.save(file_name)

    # 初始化一次浏览器
    option = ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    driver = webdriver.Chrome(chrome_options=option)
    d_num_list = []
    for i in range(1, 7):
        for j in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17',
                  '18', '19', '20']:
            d_num_list.append(int(str(i) + j))
    for d_num in d_num_list:
        data = fee(driver, b_num, d_num)
        if data is not None:
            pd_data = pd.DataFrame(data)
            file_path = file_name
            # 使用 ExcelWriter 追加数据
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                # 假设工作表名称为 'Sheet1'
                # 获取最大行数，找到空行位置以追加数据
                pd_data.to_excel(writer, sheet_name='Sheet1', index=False, header=False,
                                startrow=writer.sheets['Sheet1'].max_row)

    driver.quit()


if __name__ == '__main__':
    building_left = [7, 8, 9, 48, 49,
                     5, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59,
                     6, 60, 61, 62, 63]
    for each_num in building_left:
        run_one_building(each_num)