'''
帮我生成一个月度员工打卡时间统计表, 表名clock.xlsx, 要求如下:

1. 行索引为员工姓名, 数量为50个(例如: 李飞, 乔峰, 慕容复).
2. 列索引为日期, 范围1~30(例如1, 2, 3).
3. 每个单元格内容为时间字符串, 代表员工每日打卡时间, 范围07:00 ~ 20:00。可以有0~4个, 当有多个时间时, 时间通过换行分割, 顺序排序, 且第一个时间范围为09:30到10:20之间, 最后一个时间范围为18:50到20:00之间.
4. 多个时间的情况, 第一个时间80%的概率在9:40 ~ 10:00的区间, 20%概率在10:00之后. 最后一个时间90%概率在19:30 ~ 20:30之间, 10%概率在19:00 ~ 19:30之间.
5. 80%的概率有两个时间, 10%的概率有3~4个时间, 8%的概率有一个时间, 2%的概率没有时间.

第二步: 你已经生成了样本数据, 现在请你对这些数据进行一下处理:

1. 读取clock.xlsx, 找出单元格中的最早时间和最晚时间, 判断最早时间是否超过10:00(迟到), 最晚时间是否早于19:30(早退), 将这些单元格背景颜色修改为红色.
2. 追加两列数据, 索引为'迟到次数'和'早退次数', 值为第1步中该员工这一整行数据统计的'迟到'和'早退'的次数.
3. 追加两行数据, 索引为'迟到人数'和'早退人数', 值为每一天的迟到和早退的员工数量.
4. 输出clock_processed.xlsx文件.
'''

import pandas as pd
import random

# 定义每个员工每天的打卡时间
def generate_punch_time():
    punch_time = []
    # 80%的概率有两个时间
    if random.random() < 0.8:
        punch_time.append(generate_one_punch_time())
        punch_time.append(generate_one_punch_time())
    # 10%的概率有3~4个时间
    elif random.random() < 0.9:
        punch_time.append(generate_one_punch_time())

def generate_one_punch_time():
    # 80%的概率在9:40 ~ 10:00的区间
    if random.random() < 0.8:
        return '09:40'
    # 20%概率在10:00之后
    else:
        return '10:00'

def generate_clock():
    # 生成员工姓名
    names = ['李飞', '乔峰', '慕容复']
    # 生成日期
    days = [i for i in range(1, 31)]
    # 生成打卡时间
    clock = pd.DataFrame(columns=days, index=names)
    for name in names:
        for day in days:
            clock.loc[name, day] = generate_punch_time()
    # 保存数据
    clock.to_excel('clock.xlsx')

def process_clock():
    '''
    1. 读取clock.xlsx, 找出单元格中的最早时间和最晚时间, 判断最早时间是否超过10:00(迟到), 最晚时间是否早于19:30(早退), 将这些单元格背景颜色修改为红色.
    2. 追加两列数据, 索引为'迟到次数'和'早退次数', 值为第1步中该员工这一整行数据统计的'迟到'和'早退'的次数.
    3. 追加两行数据, 索引为'迟到人数'和'早退人数', 值为每一天的迟到和早退的员工数量.
    4. 输出clock_processed.xlsx文件.
    '''
    # 读取数据
    clock = pd.read_excel('clock.xlsx', index_col=0)
    # 找出单元格中的最早时间和最晚时间, 判断最早时间是否超过10:00(迟到), 最晚时间是否早于19:30(早退), 将这些单元格背景颜色修改为红色.
    # clock = clock.style.applymap(lambda x: 'background-color: red' if x < '10:00' or x > '19:30' else '')
    # 追加两列数据, 索引为'迟到次数'和'早退次数', 值为第1步中该员工这一整行数据统计的'迟到'和'早退'的次数.
    clock['迟到次数'] = clock.apply(lambda x: x[x < '10:00'].count(), axis=1)
    clock['早退次数'] = clock.apply(lambda x: x[x > '19:30'].count(), axis=1)
    # clock['迟到次数'] = clock.applymap(lambda x: 1 if x < '10:00' else 0).sum(axis=1)
    # clock['早退次数'] = clock.applymap(lambda x: 1 if x > '19:30' else 0).sum(axis=1)
    # 追加两行数据, 索引为'迟到人数'和'早退人数', 值为每一天的迟到和早退的员工数量.
    clock.loc['迟到人数'] = clock.apply(lambda x: x[x < '10:00'].count(), axis=0)
    clock.loc['早退人数'] = clock.apply(lambda x: x[x > '19:30'].count(), axis=0)
    # 输出数据
    clock.to_excel('clock_processed.xlsx')

if __name__ == '__main__':
    # 生成样本数据
    # generate_clock()
    # 处理样本数据
    process_clock()