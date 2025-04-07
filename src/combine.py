import pandas as pd


file_num_list = []
for i in range(1,64):
    file_num_list.append(i)
file_num_list.remove(2)
file_num_list.remove(17)
file_num_list.remove(18)
file_num_list.remove(19)


file_name_list = []
for num in file_num_list:
    file_name_list.append(str(num)+"号楼.xlsx")
    
# 读取多个表格
df_list = []
for file_name in file_name_list:
    df_list.append(pd.read_excel(file_name))


# 合并表格
combined_df = pd.concat(df_list, ignore_index=True)

# 保存合并后的表格
combined_df.to_excel('水电历史账单总表.xlsx', index=False)
