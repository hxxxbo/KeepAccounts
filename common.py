import yaml

with open('config.yml', 'r') as f:
    config = yaml.safe_load(f)


def load_config():
    with open('invalid.txt', 'r', encoding='utf-8') as file:
        terms = [line.strip() for line in file.readlines()]
    return config, terms


def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.map(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data


def preprocessing(data):  # 预处理
    o = config['ori']

    # 月份
    for index in range(len(data.iloc[:, o['time']])):
        time = data.iloc[index, o['time']]
        data.iloc[index, o['month']] = time.month  # 访问月份属性的值，赋给这月份列

    # 逻辑收支
    for index in range(len(data.iloc[:, o['inOrex']])):  # 遍历第3列的值，判断为收入，则改'逻辑1'为1
        if data.iloc[index, o['inOrex']] == '支出':
            data.iloc[index, o['money']] = -abs(data.iloc[index, o['money']])

    return data

def classification(data):

    return data
