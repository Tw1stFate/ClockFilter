import pandas as pd
import random

# 定义员工名字列表
employees = ['张三', '李四', '王五', '赵六', '钱七', '孙八', '周九', '吴十', '郑十一', '冯十二',
             '陈十三', '楚十四', '魏十五', '蒋十六', '沈十七', '韩十八', '杨十九', '朱二十', '秦二十一',
             '尤二十二', '许二十三', '何二十四', '吕二十五', '施二十六', '张三十', '李三十一', '王三十二',
             '赵三十三', '钱三十四', '孙三十五', '周三十六', '吴三十七', '郑三十八', '冯三十九', '陈四十',
             '楚四十一', '魏四十二', '蒋四十三', '沈四十四', '韩四十五', '杨四十六', '朱四十七', '秦四十八',
             '尤四十九', '许五十']

# 随机生成员工的每日打卡时间
data = []
for i in range(len(employees)):
    row = []
    for j in range(30):
        punches = []
        for k in range(random.randint(0, 4)):
            hour = random.randint(7, 19)
            minute = random.randint(0, 59)
            punches.append('{:02d}:{:02d}'.format(hour, minute))
        punches.sort()
        row.append('\n'.join(punches))
    data.append(row)

# 创建 DataFrame 对象
df = pd.DataFrame(data, index=employees, columns=range(1, 31))

# 将 DataFrame 写入 Excel 文件
with pd.ExcelWriter('clock.xlsx') as writer:
    df.to_excel(writer)