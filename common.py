import os
import pandas as pd
import yaml
import datetime


# 使用 'with open' 语句打开文件 'config.yml'，并以只读模式 ('r') 读取文件内容
# 'with' 语句确保文件在使用后会自动关闭，即使发生异常
with open('config.yml', 'r') as f:
    # 使用 yaml.safe_load 函数将文件内容解析为 Python 数据结构（通常是字典）
    # safe_load 是安全的 YAML 解析方法，避免执行任意代码的风险
    config = yaml.safe_load(f)

    # 从解析后的配置字典中获取键为 'ori' 的值，并将其赋值给变量 o
    # 假设 'ori' 是配置文件中的一个配置项
    o = config['ori']

    # 从解析后的配置字典中获取键为 'filePath' 的值，并将其赋值给变量 f
    # 假设 'filePath' 是配置文件中的一个配置项，通常表示文件路径
    f = config['filePath']


# 载入配置文件和无效词
def load_config():
    with open(f['invalid'], 'r', encoding='utf-8') as file:
        terms = [line.strip() for line in file.readlines()]
    return config, terms


# 载入已处理文件
def load_processed_files(file_path):
    processed_files = set()
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            for line in file:
                processed_files.add(line.strip())  # 移除行尾的换行符
    return processed_files


# 保存已处理文件
def save_processed_files(file_path, processed_files):
    now = datetime.datetime.now()
    now = '👆导入时间：' + str(now.strftime('%Y-%m-%d %H:%M:%S'))
    with open(file_path, 'a') as file:
        for processed_file in processed_files:
            file.write(processed_file + '\n')
    with open(file_path, 'a') as file:
        file.write(now + '\n')


def save_counts_to_excel(counts_dict, filename="file.xlsx", sheet_name="tagnum"):
    """将匹配计数保存到指定工作表的Excel文件"""
    # 将字典转换为DataFrame
    df = pd.DataFrame(list(counts_dict.items()), columns=['Item', 'Count'])

    # 写入Excel的指定工作表
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        # 检查工作表是否已存在，不存在则创建
        if sheet_name not in writer.book.sheetnames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # 工作表已存在，需要先读取现有数据，合并后再写入
            existing_df = pd.read_excel(filename, sheet_name=sheet_name)
            combined_df = pd.concat([existing_df, df]).drop_duplicates(subset='Item', keep='last')
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)


# 把列名中和数据中首尾的空格都去掉
def strip_in_data(data):
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.map(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data


# 加入月份列和收支正负
def preprocessing(data):
    for index in range(len(data.iloc[:, o['time']])):
        # 月份
        time = data.iloc[index, o['time']]
        data.iloc[index, o['month']] = time.month  # 访问月份属性的值，赋给这月份列

        # 来源
        if '2367' in str(data.iloc[index, o['payWay']]):
            data.iloc[index, o['from']] = '中国银行2367'
        elif '8672' in str(data.iloc[index, o['payWay']]):
            data.iloc[index, o['from']] = '农业银行8672'
        elif '4382' in str(data.iloc[index, o['payWay']]):
            data.iloc[index, o['from']] = '招商银行4382'

    # 逻辑收支
    # for index in range(len(data.iloc[:, o['inOrex']])):  # 遍历第3列的值，判断为收入，则改'逻辑1'为1
    #     if data.iloc[index, o['inOrex']] == '支出':
    #         data.iloc[index, o['money']] = -abs(data.iloc[index, o['money']])

    return data


# 自动分类
def classification(data):
    data.insert(o['class'], '分类', "", allow_duplicates=True)  # 插入列，默认值为 null
    data.insert(o['tag'], '标记', "", allow_duplicates=True)  # 插入列，默认值为 null

    excel_data = pd.read_excel(f['tag'])

    # 获取 Excel 中的分类名和对应的字符
    categories = {}
    for col in excel_data.columns:
        categories[col] = set(excel_data[col].dropna().values)

    # 遍历 CSV 数据，将匹配到的行写入分类名
    for index, row in data.iterrows():
        goods_value = str(row.iloc[o['goods']])
        counterparty_value = str(row.iloc[o['counterparty']])
        value = goods_value + counterparty_value

        if_hit = False
        for category, items in categories.items():
            for item in items:
                if item in value:
                    print(str(index) + "-" + str(category) + "-" + str(item) + "-" + str(value))
                    data.iloc[index, o['class']] = str(category).replace(".1", "")
                    data.iloc[index, o['tag']] = str(item)
                    if_hit = True
                    break
            if if_hit:
                break

    return data
