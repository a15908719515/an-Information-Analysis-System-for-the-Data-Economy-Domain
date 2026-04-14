import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import re
import os


# 定义函数：生成单个表格的作者关系图
def generate_author_graph(df, sheet_name, output_folder='author'):
    """
    根据给定的数据框生成作者关系图，并保存为图片，并输出合作数量前5名的作者组合及其合作次数。

    参数：
    - df: pandas DataFrame，包含作者信息的表格数据
    - sheet_name: 字符串，当前处理的表格名称
    - output_folder: 字符串，保存生成图片的文件夹名称
    """
    # 提取作者信息，只保留 ' - ' 或 '…' 之前的部分
    df = df.dropna(subset=['gs_a'])  # 删除gs_a列中为空的行
    author_lists = df['gs_a'].apply(lambda x: re.split(r' - |…', x)[0].strip())  # 去除前后空格

    # 进一步清理数据：只保留字母、空格和逗号，确保逗号后跟随空格
    author_lists = author_lists.apply(lambda x: re.sub(r'[^a-zA-Z ,]', '', x))

    # 分割作者列表，确保按正确的逗号分割，并过滤掉只有逗号的情况
    author_lists = author_lists.apply(
        lambda x: [author.strip() for author in x.split(',') if author.strip() != "," and author.strip()])

    # 统计每个作者的出现次数
    author_counter = {}
    for authors in author_lists:
        for author in authors:
            if author in author_counter:
                author_counter[author] += 1
            else:
                author_counter[author] = 1

    # 筛选出出现次数大于1次的作者
    filtered_author_lists = author_lists.apply(
        lambda authors: [author for author in authors if author_counter.get(author, 0) > 1]
    )

    # 创建一个空的无向图
    G = nx.Graph()

    # 遍历每篇论文的作者列表
    for authors in filtered_author_lists:
        # 更新每个作者的出现次数（节点权重）
        for author in authors:
            if G.has_node(author):
                G.nodes[author]['count'] += 1
            else:
                G.add_node(author, count=1)

        # 更新作者之间的合作次数（边权重）
        for i in range(len(authors)):
            for j in range(i + 1, len(authors)):
                author_a = authors[i]
                author_b = authors[j]
                if G.has_edge(author_a, author_b):
                    G[author_a][author_b]['weight'] += 1
                else:
                    G.add_edge(author_a, author_b, weight=1)

    # 输出合作数量前5名的作者组合及其合作次数
    top_5_edges = sorted(G.edges(data=True), key=lambda x: x[2]['weight'], reverse=True)[:5]
    print(f"在表格 {sheet_name} 中，合作数量前5名的作者组合及其合作次数为：")
    for u, v, data in top_5_edges:
        print(f"{u} 和 {v} 合作次数：{data['weight']}")

    # 根据作者出现次数设置节点大小
    node_sizes = [G.nodes[author]['count'] * 100 for author in G.nodes()]

    # 找到出现次数最多的作者
    max_author = max(G.nodes(), key=lambda author: G.nodes[author]['count'])
    max_count = G.nodes[max_author]['count']
    print(f"在表格 {sheet_name} 中，出现次数最多的作者是：{max_author}，出现次数为：{max_count}")

    # 根据合作次数设置边宽度
    edge_widths = [G[u][v]['weight'] for u, v in G.edges()]

    # 设置绘图布局，并增大图形尺寸和节点之间的距离
    plt.figure(figsize=(15, 15))  # 增大图形的尺寸
    pos = nx.spring_layout(G, k=1.0, seed=42)  # 调整k值以增大节点间距离

    # 绘制节点，并加深颜色
    nx.draw_networkx_nodes(G, pos, node_size=node_sizes, node_color='royalblue', alpha=0.85,
                           label="Nodes: Author's Appearance")

    # 绘制边
    nx.draw_networkx_edges(G, pos, width=edge_widths, edge_color='grey', alpha=0.5, label="Edges: Co-authorship Count")

    # 绘制节点标签，添加出现次数并设置背景框
    labels = {author: f"{author} ({G.nodes[author]['count']})" for author in G.nodes()}
    nx.draw_networkx_labels(G, pos, labels, font_size=10, font_family='sans-serif',
                            bbox=dict(facecolor='white', edgecolor='none', alpha=0.7))

    # 设置标题
    plt.title(f"{sheet_name} Author Relationship Graph", fontsize=15)

    # 添加图例
    plt.legend(loc="upper right")

    # 去除坐标轴
    plt.axis('off')

    # 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 保存图片
    plt.savefig(os.path.join(output_folder, f"{sheet_name}_author_relationship_graph.png"), dpi=300)
    plt.close()
    print(f"{sheet_name} 的作者关系图已成功保存！")


# 主函数：读取Excel文件并生成所有表格的作者关系图
def main():
    # 定义Excel文件路径
    file_path = 'Google_sicence.xlsx'  # 确保该文件与代码在同一目录下

    # 读取Excel文件中的所有表格
    xls = pd.ExcelFile(file_path)

    # 遍历每个表格并生成作者关系图
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        generate_author_graph(df, sheet_name)

    print("所有作者关系图已生成并保存至 author 文件夹！")


# 执行主函数
if __name__ == '__main__':
    main()
