import os
import pandas as pd
import shutil


def read_all_csv(folder_path):
    csv_files = {}
    for root, dirs, files in os.walk(folder_path):
        print(len(files))
        ind1 = 0
        ind2 = 0
        ind3 = 0
        ind4 = 0
        ind5 = 0
        ind6 = 0
        for file in files:
            if file.endswith('.csv'):
                file_path = os.path.join(root, file)
                with open(file_path, encoding='utf-8') as csvfile:
                    line=csvfile.read()
                    if line.find('语音呼叫')>0 and line.find('接入')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'a'+str(ind1)+'.csv'
                        ind1+=1
                        shutil.copy(file_path, new_file_path)
                        print(file_path)
                        continue
                    if line.find('视频呼叫')>0 and line.find('接入')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'b'+str(ind2)+'.csv'
                        ind2+=1
                        shutil.copy(file_path, new_file_path)
                        continue
                    if line.find('语音呼叫')>0 and line.find('保持')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'c'+str(ind3)+'.csv'
                        ind3+=1
                        shutil.copy(file_path, new_file_path)
                        continue
                    if line.find('视频呼叫')>0 and line.find('保持')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'d'+str(ind4)+'.csv'
                        ind4+=1
                        shutil.copy(file_path, new_file_path)
                        continue
                    if line.find('切换')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'e'+str(ind5)+'.csv'
                        ind5+=1
                        shutil.copy(file_path, new_file_path)
                        continue
                    if line.find('通话质量')>0:
                        new_file_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'+'f'+str(ind6)+'.csv'
                        ind6+=1
                        shutil.copy(file_path, new_file_path)
                        continue

if __name__ == "__main__":
    folder_path = 'C:/Users/24253/Desktop/新建文件夹 (4)/'  # 请替换为实际的文件夹路径
    result = read_all_csv(folder_path)
    if result is not None:
        for file_name, df in result.items():
            print(f"文件 {file_name} 的内容信息:")
            df.info()
