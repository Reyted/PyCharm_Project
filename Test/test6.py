import csv
import os


def merge_all_csv_files_in_path(path, output_file):
    # 检查输出文件是否已经存在，如果存在则删除
    if os.path.exists(output_file):
        os.remove(output_file)
    header = None
    input_files = []
    # 遍历指定路径下的所有文件
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith('.csv'):
                input_files.append(os.path.join(root, file))
    for file in input_files:
        with open(file, 'r', encoding='utf-8') as csv_file:
            reader = csv.reader(csv_file)
            if header is None:
                header = next(reader)
                with open(output_file, 'w', newline='', encoding='utf-8') as out_file:
                    writer = csv.writer(out_file)
                    writer.writerow(header)
            else:
                next(reader)  # 跳过已经存在的表头
            for row in reader:
                with open(output_file, 'a', newline='', encoding='utf-8') as out_file:
                    writer = csv.writer(out_file)
                    writer.writerow(row)

file_path='C:/Users/24253/Desktop/工作内容/PRS数据/历史告警20250116154919400/'
# 调用函数，传入路径和输出文件的名称
merge_all_csv_files_in_path(file_path, f'{file_path}历史告警.csv')





