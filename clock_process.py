import pandas as pd

# 定义迟到和早退的时间阈值
late_time = '10:01'
early_time = '19:29'

# 读取 Excel 文件
df = pd.read_excel('clock.xlsx', index_col=0)

# 创建一个与 DataFrame 相同形状的空白 DataFrame，用于存储背景颜色
color_df = pd.DataFrame('', index=df.index, columns=df.columns)

# 遍历 DataFrame 中的每个单元格，找出最早时间和最晚时间，并判断是否迟到或早退
for i in range(len(df)):
    for j in range(len(df.columns)):
        cell = df.iloc[i, j]
        if isinstance(cell, str):
            punches = cell.split('\n')
            earliest_time = punches[0]
            latest_time = punches[-1]
            if earliest_time > late_time and latest_time < early_time:
                color_df.iloc[i, j] = 'background-color: red'
            else:
                if earliest_time > late_time:
                    color_df.iloc[i, j] = 'background-color: orange'
                if latest_time < early_time:
                    color_df.iloc[i, j] = 'background-color: purple'

# 追加两列数据, 索引为'迟到次数'和'早退次数', 值为每行数据统计的red和blue的次数.
# TypeError: 'Styler' object does not support item assignment, 不能直接修改styled_df的值, 应该先转换为DataFrame, 然后再修改.
df['迟到次数'] = color_df.apply(lambda x: x.str.contains('red|orange').sum(), axis=1)
df['早退次数'] = color_df.apply(lambda x: x.str.contains('red|purple').sum(), axis=1)

# 追加两行数据, 索引为'迟到人数'和'早退人数', 值为每一天的迟到和早退的员工数量.
df.loc['迟到人数'] = color_df.apply(lambda x: x.str.contains('red|orange').sum(), axis=0)
df.loc['早退人数'] = color_df.apply(lambda x: x.str.contains('red|purple').sum(), axis=0)

# 将背景颜色 DataFrame 与原 DataFrame 合并，并生成带有样式的 Excel 文件
styled_df = df.style.apply(lambda x: color_df, axis=None).set_table_attributes('border="1"')

styled_df.to_excel('clock_processed.xlsx', engine='openpyxl', index=True)