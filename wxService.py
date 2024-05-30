import pandas as pd
import common


def read_data(path):  # 获取微信数据
    config, terms = common.load_config()

    o = config['ori']
    c = config['wx']
    d = pd.read_csv(path, header=c['header'], skipfooter=c['footer'], encoding='utf-8')  # 数据获取，微信

    d = d.iloc[:, [c['time'], c['inOrex'], c['payWay'], c['state'], c['type'], c['counterparty'], c['goods'],
                   c['amount']]]  # 按顺序提取所需列
    d = common.strip_in_data(d)  # 去除列名与数值中的空格。
    d.iloc[:, o['time']] = d.iloc[:, o['time']].astype('datetime64[ns]')  # 数据类型更改
    d.insert(o['month'], '月份', 0, allow_duplicates=True)  # 插入列，默认值为0
    d.insert(o['from'], '来源', "微信", allow_duplicates=True)  # 添加微信来源标识
    d.iloc[:, o['money']] = d.iloc[:, o['money']].astype('float64')  # 数据类型更改

    d.rename(columns={'当前状态': '支付状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称

    filter_pattern = '|'.join(terms)

    rm1 = d[d['收/支'] == '/']  # 保存'收/支'为'/'的行
    rm2 = d[d['支付状态'].str.contains(filter_pattern)]

    rm = pd.concat([rm1, rm2])
    d = d.drop(rm.index)  # 删除'收/支'为'/'的行
    len1 = len(d)
    print("成功读取 " + str(len1) + " 条「微信」账单数据\n")
    return d, rm
