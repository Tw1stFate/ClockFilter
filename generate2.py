import pandas as pd
import random

# 定义每个员工每天的打卡时间
def generate_punch_time():
    num_punches = 0
    if random.random() < 0.8:  # 80%的概率有两个时间
        num_punches = 2
    elif random.random() < 0.9:  # 10%的概率有3~4个时间
        num_punches = random.randint(3, 4)
    elif random.random() < 0.93:  # 3%的概率有一个时间
        num_punches = 1
    else:  # 2%的概率没有时间
        num_punches = 0
    punch_times = []
    for i in range(num_punches):
        hour = random.randint(7, 19)  # 随机生成小时数
        minute = random.randint(0, 59)  # 随机生成分钟数
        punch_time = f'{hour:02d}:{minute:02d}'
        punch_times.append(punch_time)
    punch_times.sort()  # 将打卡时间按从早到晚排序
    if punch_times:  # 判断 punch_times 是否为空列表
        if random.random() < 0.8:  # 第一个打卡时间在 9:40 至 10:00 之间的概率为 80%
            first_punch_time = f'09:{random.randint(40, 59):02d}' if punch_times[0][:2] == '09' else f'10:{random.randint(0, 0):02d}' # 如果第一个打卡时间已经是10:xx，则将第一个打卡时间设置为10:00
        else:
            first_punch_time = f'10:{random.randint(0, 30):02d}'  # 第一个打卡时间在 10:00 之后的概率为 20%
        punch_times[0] = first_punch_time
        
        if random.random() < 0.9:  # 最后一个打卡时间在 19:30 至 20:30 之间的概率为 90%
            last_punch_time = f'{random.randint(19, 20):02d}:{random.randint(30, 59):02d}'
        else:
            last_punch_time = f'{random.randint(19, 19):02d}:{random.randint(0, 30):02d}'  # 最后一个打卡时间在 19:00 至 19:30 之间的概率为 10%
        punch_times[-1] = last_punch_time
    return '\n'.join(punch_times)

# 生成数据
names = [f'员工{i+1}' for i in range(50)]
dates = [f'{i}号' for i in range(1, 31)]
data = [[generate_punch_time() for _ in dates] for _ in names]

# 创建 DataFrame 对象
df = pd.DataFrame(data, index=names, columns=dates)

# 将 DataFrame 对象写入 Excel 文件
df.to_excel('clock.xlsx', index=True)