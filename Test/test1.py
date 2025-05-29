import tkinter as tk
from tkinter import Tk, filedialog
import os

def list_files_recursively(path):
    """递归列出路径下的所有文件"""
    file_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            file_list.append(file_path)
    return file_list

def select_directory():
    """使用Tkinter选择目录"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="请选择要遍历的文件夹")
    root.destroy()
    return folder_path

if __name__ == '__main__':
    folder_path = select_directory()
    #print(list_files_recursively(folder_path))
    lst=list_files_recursively(folder_path)
    #print(lst)
    for file in lst:
        if 'NR早晚指标监控' in file:
            print(file)
        if 'LTE早晚指标' in file and '泰山' not in file and '新网' not in file:
            print(file)