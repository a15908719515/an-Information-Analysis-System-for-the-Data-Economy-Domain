import pandas as pd
import matplotlib.pyplot as plt
import os
import re

# 创建 influence_class 文件夹，如果不存在的话
output_folder = 'influence_class'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 读取 Excel 文件
file_path = 'Google_sicence.xlsx'
sheet_names = pd.ExcelFile(file_path).sheet_names  # 获取所有标签页的名称

# 创建一个字典来存储每个表的年份与引用次数的汇总数据
citation_data = {}

# 循环遍历每个标签页
for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)

    # 提取年份
    df['Year'] = df['gs_a'].apply(
        lambda x: int(re.findall(r'\b\d{4}\b', str(x))[0]) if re.findall(r'\b\d{4}\b', str(x)) else None)

    # 舍弃1960年之前和2024年之后的年份
    df = df[(df['Year'] >= 1960) & (df['Year'] <= 2023)]

    df = df.dropna(subset=['Year'])  # 去掉没有年份的行

    # 提取引用次数
    df['Citation'] = df['gs_fl'].apply(
        lambda x: int(re.findall(r'\d+', str(x))[0]) if re.findall(r'\d+', str(x)) else 0)

    # 按年份汇总引用次数
    yearly_citations = df.groupby('Year')['Citation'].sum()

    # 确保所有年份都有值，填充缺失年份为0
    full_years = pd.Series(0, index=range(int(df['Year'].min()), int(df['Year'].max()) + 1))
    yearly_citations = full_years.add(yearly_citations, fill_value=0)

    # 将数据存储在字典中
    citation_data[sheet] = yearly_citations

# 计算总引用量
total_citations = pd.Series(dtype=float)

for data in citation_data.values():
    total_citations = total_citations.add(data, fill_value=0)

# 绘制所有表的引用次数随年份变化的曲线图
plt.figure(figsize=(12, 8))

for sheet, data in citation_data.items():
    plt.plot(data.index.to_numpy(), data.values, marker='o', label=sheet)

# 绘制总引用量的曲线
plt.plot(total_citations.index.to_numpy(), total_citations.values, marker='o', linestyle='--', color='black',
         label='Total Citations')

plt.xlabel('Year')
plt.ylabel('Total Citations')
plt.title('Total Citations per Year by Category')
plt.legend(title='Category')
plt.grid(True)
plt.tight_layout()

# 保存图表到当前工作目录中
output_path = 'total_citations_by_year.png'
plt.savefig(output_path)

# 显示图表
plt.show()
