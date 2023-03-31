import requests
import pandas as pd

# 确认 API 密钥和基础 URL
api_key = 'DA302G5VVA1TL5QE'
base_url = 'https://www.alphavantage.co/query'

# 获取要查询的股票代码和数据类型
symbols = input('请输入要查询的股票代码（多个代码用逗号分隔）：').split(',')
datatype = input('请输入要查询的数据类型（1：当日数据；2：多天数据）：')

# 根据数据类型获取数据
if datatype == '1':
    # 查询当日数据
    data = {}
    for symbol in symbols:
        # 构造 API 请求 URL
        url = f'{base_url}?function=GLOBAL_QUOTE&symbol={symbol}&apikey={api_key}'

        # 发送 API 请求并获取响应
        response = requests.get(url)
        data[symbol] = response.json()

    # 提取收盘价
    closing_prices = []
    for symbol in symbols:
        if 'Global Quote' in data[symbol]:
            closing_price = float(data[symbol]['Global Quote']['05. price'])
            closing_prices.append(closing_price)
        else:
            closing_prices.append(None)

    # 构造 DataFrame 并输出收盘价
    df = pd.DataFrame({'股票代码': symbols, '收盘价': closing_prices})
    print(df)

elif datatype == '2':
    # 查询多天数据
    num_days = int(input('请输入要查询的天数：'))

    # 查询数据并提取收盘价
    dates = []
    closing_prices_dict = {}
    for symbol in symbols:
        # 构造 API 请求 URL
        url = f'{base_url}?function=TIME_SERIES_DAILY_ADJUSTED&symbol={symbol}&apikey={api_key}'

        # 发送 API 请求并获取响应
        response = requests.get(url)
        data = response.json()

        # 提取收盘价
        closing_prices = []
        for date in sorted(data['Time Series (Daily)'], reverse=True)[:num_days]:
            closing_price = float(data['Time Series (Daily)'][date]['4. close'])
            closing_prices.append(closing_price)
            if symbol not in closing_prices_dict:
                closing_prices_dict[symbol] = []
            closing_prices_dict[symbol].append(closing_price)
            if symbol == symbols[0]:
                dates.append(date)

    # 构造 DataFrame 并输出收盘价
    closing_prices_matrix = []
    for symbol in symbols:
        closing_prices_matrix.append(closing_prices_dict[symbol])
    df = pd.DataFrame(closing_prices_matrix, index=symbols, columns=dates)
    print(df)

else:
    print('请选择有效的数据类型。')

# 是否导出为 Excel
export_to_excel = input('是否导出为 Excel 文件？（y/n）')

if export_to_excel.lower() == 'y':
    file_path = input('请输入文件保存路径和文件名（例如：C:/Users/username/Desktop/stock_data.xlsx）：')
    # c:/users/asus/desktop/stock.xlsx
    df.to_excel(file_path)
    print('已将数据保存为 Excel 文件。')
elif export_to_excel.lower() == 'n':
    print('已选择不导出为 Excel 文件。')
else:
    print('无效的选项，已选择不导出为 Excel 文件。')