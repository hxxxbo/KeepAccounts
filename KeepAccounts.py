import os

import pandas as pd
import openpyxl
import wxService as wx
import zfbService as zfb
import common

if __name__ == '__main__':

    config, terms = common.load_config()
    f = config['filePath']

    root_dir = f['root_dir']  # 替换为你的目录路径
    processed_file_list_path = f['processed_file']  # 存储已处理文件的文本文件路径

    # 加载已处理文件集合
    processed_files = common.load_processed_files(processed_file_list_path)

    path_write = f['path_write']

    data_merge = pd.DataFrame()
    rm_merge = pd.DataFrame()

    processed_count = 0

    # 遍历目录及其子目录，获取所有数据
    for dir_path, dirs, files in os.walk(root_dir):
        for file_name in files:
            file_path = os.path.join(dir_path, file_name)

            # 如果文件尚未处理
            if file_path not in processed_files:
                # 处理文件
                if "alipay" in file_path:
                    processed_count += 1
                    data, rm = zfb.read_data(file_path)  # 读数据
                elif "微信" in file_path:
                    processed_count += 1
                    data, rm = wx.read_data(file_path)  # 读数据
                else:
                    continue

                data_merge = pd.concat([data_merge, data], axis=0, ignore_index=True)  # 上下拼接合并表格
                rm_merge = pd.concat([rm_merge, rm], axis=0)  # 上下拼接合并表格

                # 标记文件为已处理
                processed_files.add(file_path)

    # 输出需要处理的文件数量
    print(f"Found {processed_count} files to process.\n")

    # 检查是否处理了新文件
    if processed_count == 0:
        print("No new files to process. Exiting the program.\n")
        exit(0)

    data_merge = common.preprocessing(data_merge)
    data_merge = common.classification(data_merge)

    merge_list = data_merge.values.tolist()  # 格式转换，DataFrame->List
    workbook = openpyxl.load_workbook(path_write)  # openpyxl读取账本文件
    sheet = workbook['明细']

    maxRow = sheet.max_row  # 获取最大行

    # for row in reversed(range(2, maxRow + 1)):
    #     sheet.delete_rows(row)
    #     maxRow = sheet.max_row  # 获取最大行
    # print('\n「明细」 sheet 页原有 ' + str(maxRow) + ' 行数据，将在抹除后写入数据')

    for row in merge_list:
        sheet.append(row)  # openpyxl写文件

    # 保存已各种文件
    common.save_processed_files(processed_file_list_path, processed_files)
    with pd.ExcelWriter(f['invalid_file']) as writer:
        rm_merge.to_excel(writer, sheet_name="无效数据", index=False)
    workbook.save(path_write)

    print("\n成功将数据写入到 " + path_write)
    print("\n运行成功！write successfully!")
    exit(0)
