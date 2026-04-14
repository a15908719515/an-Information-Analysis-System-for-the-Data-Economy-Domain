import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import DBSCAN
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
import numpy as np
import os

# 创建 new_class 文件夹
output_folder = 'new_class'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 读取 Excel 文件
file_path = 'Google_sicence.xlsx'
sheet_names = pd.ExcelFile(file_path).sheet_names  # 获取所有标签页的名称

# 合并所有表的数据
combined_df = pd.DataFrame()

for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)
    df['Sheet'] = sheet  # 添加一个列标识所属表
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# 提取标题
titles = combined_df.iloc[:, 0].dropna().values.astype('U')  # 转换为 Unicode 字符串

# 使用 TfidfVectorizer 将标题转换为 TF-IDF 特征矩阵
vectorizer = TfidfVectorizer(stop_words='english')
X = vectorizer.fit_transform(titles)

# 标准化数据
X_scaled = StandardScaler(with_mean=False).fit_transform(X)

# 使用 PCA 将高维数据降维到二维
pca = PCA(n_components=2)
X_pca = pca.fit_transform(X_scaled.toarray())

# 使用 DBSCAN 进行聚类，调整参数以减少簇的数量
dbscan = DBSCAN(eps=0.855, min_samples=3, metric='cosine')
clusters = dbscan.fit_predict(X_scaled)

# 统计噪声点数量 (Cluster -1)
noise_points = np.sum(clusters == -1)
print(f"噪声点的数量: {noise_points}")

# 将聚类结果添加到数据框
combined_df = combined_df.loc[combined_df.iloc[:, 0].notna()]  # 保留有标题的行
combined_df['Cluster'] = clusters

# 按簇将数据保存到不同的 Excel 表中
unique_clusters = np.unique(clusters)
for cluster in unique_clusters:
    if cluster != -1:  # 忽略噪声点
        cluster_df = combined_df[combined_df['Cluster'] == cluster]
        cluster_df.to_excel(os.path.join(output_folder, f'cluster_{cluster}.xlsx'), index=False)

# 输出聚类结果到一个CSV文件中
output_path = 'combined_dbscan_clusters.csv'
combined_df.to_csv(output_path, index=False)

# 可视化聚类结果
plt.figure(figsize=(10, 7))
unique_labels = np.unique(clusters)
for label in unique_labels:
    plt.scatter(X_pca[clusters == label, 0], X_pca[clusters == label, 1], label=f'Cluster {label}')

plt.title('Clustering of Titles from Combined Sheets (PCA + DBSCAN)')
plt.xlabel('PCA Feature 1')
plt.ylabel('PCA Feature 2')
plt.legend()
plt.grid(True)
plt.tight_layout()

# 保存聚类图
plt.savefig('combined_dbscan_cluster_plot.png')
plt.show()
