import pandas as pd

# 读取文件
df = pd.read_excel("clock.xlsx", index_col=0, dtype=str)

# 将时间字符串转换为时间类型
df = df.applymap(lambda x: [pd.to_datetime(t) for t in x.split()] if isinstance(x, str) else [])

print(df)

# 定义颜色映射
color_map = {
    "迟到": "red",
    "早退": "red",
    "正常": "white"
}

# 计算迟到和早退的次数
df["迟到次数"] = df.apply(lambda row: sum([1 for t in row if all([t_i is not None and t_i.time() > pd.to_datetime("10:00").time() for t_i in t])]), axis=1)
df["早退次数"] = df.apply(lambda row: sum([1 for t in row if all([t_i is not None and t_i.time() < pd.to_datetime("19:30").time() for t_i in t])]), axis=1)

# 计算每一天的迟到和早退的人数
df.loc["迟到人数"] = (df.apply(lambda row: [all([t_i is not None and t_i.time() > pd.to_datetime("10:00").time() for t_i in t]) for t in row], axis=1)).sum()
df.loc["早退人数"] = (df.apply(lambda row: [all([t_i is not None and t_i.time() < pd.to_datetime("19:30").time() for t_i in t]) for t in row], axis=1)).sum()

# 修改单元格颜色
def highlight(row):
    return ['background-color: %s' % color_map["迟到"] if all([cell is not None and cell.time() > pd.to_datetime("10:00").time() for cell in row])
            else 'background-color: %s' % color_map["早退"] if all([cell is not None and cell.time() < pd.to_datetime("19:30").time() for cell in row])
            else '' for row in row]

df = df.style.apply(highlight, axis=1)

# 将 DataFrame 写入 Excel 文件
df.to_excel("clock_processed.xlsx", engine='openpyxl')