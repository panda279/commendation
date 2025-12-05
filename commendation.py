import pandas as pd
# 设置显示选项，取消列宽限制
pd.set_option('display.max_columns', None)  # 显示所有列
pd.set_option('display.max_rows', None)     # 显示所有行
pd.set_option('display.width', None)        # 取消宽度限制
pd.set_option('display.max_colwidth', None) # 取消列宽限制，完整显示长文本

file_path =  r"C:\Users\30384\Desktop\志愿者名单（标红为组长）.xlsx" 
df= pd.read_excel(file_path)
columns=df.columns.tolist()
user_input1=input("请输入查找到列标题")
user_input2=input("请输入查找到列标题")
user_input3=input("请输入查找到列标题")
user_input4=input("请输入查找到列标题")
data1=[]
data2=[]
data3=[]
data4=[]
if user_input1 in columns  and user_input2 in columns and user_input3 in columns and user_input4 in columns:
    data1=df[user_input1]
    data2=df[user_input2]
    data3=df[user_input3]
    data4=df[user_input4] 
else:
    if user_input1 not in columns:
       print(user_input1,"不存在")
    if user_input2 not in columns:
       print(user_input2,"不存在")
    if user_input3 not in columns:
       print(user_input3,"不存在")
    else:
        print(user_input4,"不存在")
data1=data1.dropna()
data2=data2.dropna()
data3=data3.dropna()
data4=data4.dropna()
college=data1.drop_duplicates()
name=data2.drop_duplicates()
Class=data3.drop_duplicates()
number=data4.drop_duplicates()
print(college.to_string(index=False))
print(name.to_string(index=False))
print(Class.to_string(index=False))
print(number.to_string(index=False))


