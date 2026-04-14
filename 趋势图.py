import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import re
from sklearn.preprocessing import MinMaxScaler
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense, LSTM

# 创建qushi文件夹，如果不存在的话
output_folder = 'qushi'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 读取Excel文件
file_path = 'Google_sicence.xlsx'
sheet_names = pd.ExcelFile(file_path).sheet_names  # 获取所有标签页的名称

# LSTM 模型定义
def build_lstm_model(input_shape):
    model = Sequential()
    model.add(LSTM(50, return_sequences=True, input_shape=input_shape))
    model.add(LSTM(50, return_sequences=False))
    model.add(Dense(25))
    model.add(Dense(1))
    model.compile(optimizer='adam', loss='mean_squared_error')
    return model

# 循环遍历每个标签页
for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)

    # 提取 'gs_a' 列中的年份信息
    df['Year'] = df['gs_a'].apply(lambda x: re.findall(r'\b\d{4}\b', str(x)))

    # 只保留第一个找到的年份
    df['Year'] = df['Year'].apply(lambda x: x[0] if x else None)

    # 去掉没有年份的行
    df = df.dropna(subset=['Year'])

    # 将年份列转换为整数类型
    df['Year'] = df['Year'].astype(int)

    # 舍弃1960年之前和2024年之后的年份
    df = df[(df['Year'] >= 1960) & (df['Year'] <= 2023)]

    # 统计每一年的文章数量
    yearly_counts = df['Year'].value_counts().sort_index()

    # 确保年份连续，即没有文章的年份数量为0
    full_years = pd.Series(0, index=range(yearly_counts.index.min(), yearly_counts.index.max() + 1))
    yearly_counts = full_years.add(yearly_counts, fill_value=0)

    # 数据归一化处理
    scaler = MinMaxScaler(feature_range=(0, 1))
    scaled_data = scaler.fit_transform(yearly_counts.values.reshape(-1, 1))

    # 准备训练数据
    X_train, y_train = [], []
    for i in range(5, len(scaled_data)):
        X_train.append(scaled_data[i-5:i, 0])
        y_train.append(scaled_data[i, 0])

    X_train, y_train = np.array(X_train), np.array(y_train)
    X_train = np.reshape(X_train, (X_train.shape[0], X_train.shape[1], 1))

    # 构建LSTM模型
    model = build_lstm_model((X_train.shape[1], 1))

    # 训练LSTM模型
    model.fit(X_train, y_train, batch_size=1, epochs=100)

    # 预测未来五年
    last_sequence = scaled_data[-5:].reshape(1, -1, 1)
    future_predictions = []
    for _ in range(5):
        pred = model.predict(last_sequence)
        future_predictions.append(pred[0, 0])
        last_sequence = np.append(last_sequence[:, 1:, :], pred.reshape(1, 1, 1), axis=1)

    future_predictions = scaler.inverse_transform(np.array(future_predictions).reshape(-1, 1)).flatten()

    # 生成预测数据
    future_years = range(yearly_counts.index.max() + 1, yearly_counts.index.max() + 6)
    future_counts = pd.Series(future_predictions, index=future_years)

    # 生成折线图和柱状图
    plt.figure(figsize=(12, 6))

    # 生成当前数据的柱状图
    plt.bar(yearly_counts.index.to_numpy(), yearly_counts.values, color='lightblue', alpha=0.7,
            label='Current Bar Chart')

    # 生成当前数据的折线图
    plt.plot(yearly_counts.index.to_numpy(), yearly_counts.values, marker='o', color='b', label='Current Line Chart')

    # 生成预测数据的柱状图
    plt.bar(future_counts.index.to_numpy(), future_counts.values, color='lightgreen', alpha=0.7,
            label='Predicted Bar Chart')

    # 生成预测数据的折线图
    plt.plot(future_counts.index.to_numpy(), future_counts.values, marker='o', color='g', linestyle='--',
             label='Predicted Line Chart')

    plt.title(f'Number of Articles per Year in {sheet} with 5-Year Prediction (LSTM)')
    plt.xlabel('Year')
    plt.ylabel('Number of Articles')
    plt.grid(True)
    plt.xticks(list(yearly_counts.index.to_numpy()) + list(future_counts.index.to_numpy()), rotation=45)
    plt.legend()
    plt.tight_layout()

    # 保存图表到qushi文件夹中
    output_path = os.path.join(output_folder, f'{sheet}_trend_with_prediction_lstm.png')
    plt.savefig(output_path)

    # 显示图表
    plt.show()
