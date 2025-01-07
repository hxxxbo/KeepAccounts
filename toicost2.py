import os
import shutil

import openpyxl
import pandas as pd
import common
from datetime import datetime

if __name__ == '__main__':
    config, terms = common.load_config()
    f = config['filePath']
    o = config['ori']

    out_dir = f['out_dir']  # 替换为你的目录路径
    processed_file_list_path = f['processed_file']  # 存储已处理文件的文本文件路径

    # 加载已处理文件集合
    processed_files = common.load_processed_files(processed_file_list_path)

    processed_count = 0

    # 遍历目录及其子目录，获取所有数据
    for dir_path, dirs, files in os.walk(out_dir):
        for file_name in files:
            file_path = os.path.join(dir_path, file_name)

            # 如果文件尚未处理
            if "$" not in file_path and "账单" in file_path and file_path not in processed_files:
                print(f"Processing file: {file_path}")
                # 处理文件

                # 保存文件
                current_date = datetime.now().strftime('%Y%m%d_%H%M')
                now_file_path = f'{f['out_dir']}/{current_date}_icost.xlsx'

                # 复制模板文件到新文件路径
                shutil.copyfile(f['icost_template'], now_file_path)

                df = pd.read_excel(file_path)

                # 合并数据
                df['备注'] = df.iloc[:, o['counterparty']].astype(str) + df.iloc[:, o['goods']].astype(str)

                # 删除不需要的列
                columns_to_delete_idx = [o['month'], o['payWay'], o['state'], o['type'], o['counterparty'],
                                         o['goods']]  # 要删除的列的索引
                df = df.drop(columns=df.columns[columns_to_delete_idx])

                # 修改列名
                new_column_names = {
                    '交易时间': '时间',
                    '来源': '账户1',
                    '收/支': '类型',
                    '分类': '一级分类',
                    '标记': '二级分类'
                }
                df = df.rename(columns=new_column_names)

                # 重排列的顺序
                new_column_order = ['时间', '类型', '金额', '一级分类', '二级分类', '账户1', '备注']
                df = df[new_column_order]

                df.insert(6, '账户2', "", allow_duplicates=True)
                df.insert(8, '货币', "", allow_duplicates=True)
                df.insert(9, '标签', "", allow_duplicates=True)

                # 打印修改后的列名和列索引
                print("修改后的列名：", df.columns.tolist())
                print("修改后的列索引：", df.columns)

                df_list = df.values.tolist()  # 格式转换，DataFrame->List


                # 加载新创建的Excel文件
                workbook = openpyxl.load_workbook(now_file_path)
                sheet = workbook.active

                for row in df_list:
                    sheet.append(row)  # openpyxl写文件

                workbook.save(now_file_path)

                print(f"修改后的文件已保存到 {now_file_path}")

                # 标记文件为已处理
                processed_files.add(file_path)

    # 输出需要处理的文件数量
    print(f"Found {processed_count} files to process.\n")

    common.save_processed_files(processed_file_list_path, processed_files)

    # 检查是否处理了新文件
    if processed_count == 0:
        print("No new files to process. Exiting the program.\n")
        exit(0)
