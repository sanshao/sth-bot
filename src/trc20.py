import pandas as pd

# 读取 CSV 文件
df = pd.read_csv('./src/Transfers_20260515.csv')

# 只保留 Token Symbol 为 USDT 的交易
df = df[df['Token Symbol'] == 'USDT'].copy()

# 确保时间列为 datetime 类型
df['Time(UTC)'] = pd.to_datetime(df['Time(UTC)'])

# 按时间排序
df = df.sort_values('Time(UTC)')

# 生成“净变动”列：转入为正，转出为负
df['NetAmount'] = df.apply(
    lambda row: float(row['Amount/TokenID']) if row['To'].lower() == row['To'].lower() else -float(row['Amount/TokenID']),
    axis=1
)

# 计算累计余额
df['Balance'] = df['NetAmount'].cumsum()

# 提取年月
df['YearMonth'] = df['Time(UTC)'].dt.to_period('M')

# 获取每月最后一笔交易的余额作为月末余额
monthly_balance = df.groupby('YearMonth').last()['Balance'].reset_index()
monthly_balance.columns = ['Month', 'EndOfMonthBalance']

# 输出结果
print(monthly_balance)

# 保存到 CSV
monthly_balance.to_csv('monthly_end_balance.csv', index=False)