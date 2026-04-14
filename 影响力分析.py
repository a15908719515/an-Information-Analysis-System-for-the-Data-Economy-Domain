import pandas as pd
import matplotlib.pyplot as plt
import os
import re

# 创建 influence 文件夹，如果不存在的话
output_folder = 'influence'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 读取 Excel 文件
file_path = 'Google_sicence.xlsx'
sheet_names = pd.ExcelFile(file_path).sheet_names  # 获取所有标签页的名称

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

    # 舍弃引用量不超过200的文章
    df = df[df['Citation'] > 200]

    # 为每篇文章编号
    df['Article_Number'] = range(1, len(df) + 1)

    # 获取编号、年份、引用次数
    numbers = df['Article_Number']
    years = df['Year']
    citations = df['Citation']

    # 生成气泡图
    plt.figure(figsize=(12, 8))
    scatter = plt.scatter(years, numbers, s=citations, alpha=0.5, color='skyblue')

    # 添加引用次数的数字标注，使用较小的字体
    for i in range(len(numbers)):
        plt.text(years.iloc[i], numbers.iloc[i], f'{citations.iloc[i]} ({numbers.iloc[i]})',
                 fontsize=8, ha='center', va='center',
                 color='darkblue')  # 使用深蓝色与气泡颜色相匹配

    plt.xlabel('Year')
    plt.ylabel('Article Number')
    plt.title(f'Influence Analysis by Citations in {sheet}')
    plt.grid(True)
    plt.tight_layout()

    # 保存图表到 influence 文件夹中
    output_path = os.path.join(output_folder, f'{sheet}_influence_analysis_bubble.png')
    plt.savefig(output_path)

    # 显示图表
    plt.show()

    # 保存编号与文章标题的对应关系到 CSV 文件
    mapping_output_path = os.path.join(output_folder, f'{sheet}_article_mapping.csv')
    df[['Article_Number', df.columns[0]]].to_csv(mapping_output_path, index=False, header=['Article Number', 'Title'])
