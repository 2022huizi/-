# This is a sample Python script.
import pandas as pd
import numpy as np
import os
import argparse


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


# 读取Excel中的语料
def readFile(in_path, out_dir):
    df = pd.read_excel(in_path, header=None)
    # 语料的条数
    sen_nums = df.shape[0]
    # print(df)
    # print(df.iloc[2, 0])
    # 定义结果集二维list
    my_matrix = [[0 for i in range(sen_nums)] for i in range(sen_nums)]
    # 依次两两计算相似度
    for i in range(sen_nums):
        for j in range(sen_nums):
            my_rate = compute(df.iloc[i, 0], df.iloc[j, 0])
            my_matrix[i][j] = my_rate
    # 转换成结果集数组
    data = np.array(my_matrix)
    # 定义结果集的列名,索引名同列名
    my_columns = []
    for i in range(sen_nums):
        my_columns.append(df.iloc[i, 0])
    # 创建结果集DateFrame
    res_df = pd.DataFrame(data, columns=my_columns, index=my_columns)
    # 输入的文件名
    input_file_name = os.path.basename(in_path)
    # 生成输出的文件名
    output_file_name = input_file_name[0:-len(input_file_name.split('.')[-1]) - 1] + '_result.' + \
                       input_file_name.split('.')[-1]
    # 生成输出的文件路径
    result_path = os.path.join(out_dir, output_file_name)
    # 保存文件
    res_df.to_excel(result_path, header=True)


# 计算相似度
def compute(str1, str2):
    # 初始化数据
    my_row = len(str1) + 1
    my_column = len(str2) + 1
    # 初始化Matrix
    my_matrix = [[0 for i in range(my_column)] for i in range(my_row)]
    # 初始化Matrix的第一行和第一列
    for i in range(my_column):
        my_matrix[0][i] = i
    for i in range(my_row):
        my_matrix[i][0] = i
    # for i in range(my_row):
    #     for j in range(my_column):
    #         print(my_matrix[i][j], end='')
    #     print()
    # 开始计算
    int_cost = 0
    for i in range(1, my_row):
        for j in range(1, my_column):
            if str1[i - 1] == str2[j - 1]:
                int_cost = 0
            else:
                int_cost = 1
            # 关键步骤，计算当前位置值为左边+1、上面+1、左上角+intCost中的最小值
            # 循环遍历到最后_Matrix[_Row - 1, _Column - 1]即为两个字符串的距离
            my_matrix[i][j] = min(my_matrix[i - 1][j] + 1, my_matrix[i][j - 1] + 1, my_matrix[i - 1][j - 1] + int_cost)
    # //相似率 移动次数小于最长的字符串长度的20%算同一题
    int_len = my_row if my_row > my_column else my_column
    rate = str((1 - float(my_matrix[my_row - 1][my_column - 1]) / int_len))
    if len(rate) > 6:
        rate = rate[0:6]
    return rate


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # 声明命令行参数
    parser = argparse.ArgumentParser(description='Computing For CHZ')
    parser.add_argument('-in', '--inDir', required=True, type=str, help='InputDirectory(Required)')
    parser.add_argument('-out', '--outDir', type=str, help='OutputDirectory(Not Required) Default ./result')
    args = parser.parse_args()  # 将变量以标签-值的字典形式存入args字典
    # 初始化输入文件夹和输出文件夹
    InputDirectory = args.inDir
    if args.outDir:
        OutputDirectory = args.outDir
    else:
        OutputDirectory = './result'
    # 是否存在输出文件夹
    if not os.path.exists(OutputDirectory):
        # 创建输出文件夹
        os.mkdir(OutputDirectory)
    if os.path.exists(InputDirectory):
        # 获取输入文件夹中的所有文件
        input_files = [x[2] for x in os.walk(InputDirectory)][0]
        print(input_files)
        for file in input_files:
            print('***************')
            print('Computing '+file)
            if file.split('.')[-1] == 'xls' or file.split('.')[-1] == 'xlsx':
                input_file_path = os.path.join(InputDirectory, file)
                readFile(input_file_path, OutputDirectory)
            print('Compute Finished!')
            print('***************')
    else:
        print('输入的文件夹不存在啦！！！')
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
