import pandas as pd
from openpyxl import Workbook
pd_data1 = pd.read_excel('数据完整的宿舍号.xlsx')
pd_data2 = pd.read_excel('被筛选的数据.xlsx')


#创建一个文件
workbook = Workbook()

# 获取默认的工作表
sheet = workbook.active
sheet.title = "Sheet1"

# 保存工作簿
file_name = 'new_file'
workbook.save(file_name)

#创建内容
new_list = []
for tar_dorm in pd_data1.iloc[:, 0]:
    for row in pd_data2.itertuples(index=True):
        if row[1] == tar_dorm:
            new_list.append(row)

# 将格式化好的行创建为 DataFrame
df = pd.DataFrame(new_list)

# 或者写入 Excel 文件
df.to_excel('output.xlsx', index=False)


