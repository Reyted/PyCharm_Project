import pandas as pd
import tkinter as tk
from tkinter import Tk, filedialog
import os
import platform
import datetime
import time


four_one_dic={}
five_one_dic={}

save_excel=[]
save_excel_f=[]
lst=['4G泰山上一周','4G欣网上一周','4G泰山这一周','4G欣网这一周','5G泰山上一周','5G欣网上一周','5G泰山这一周','5G欣网这一周']

def fn1(filename1,filename2):
    pf1=pd.read_excel(filename1)
    pf2=pd.read_excel(filename2)
    time1=''
    for li in pf1.values:
        four_one_dic[li[1]+"-"+str(li[5])]=[li[7],li[8],li[9]]
        time1=li[0]
    for li in pf2.values:
        str1=li[1]+"-"+str(li[5])
        try:
            if (four_one_dic[str1][0]-li[7])/four_one_dic[str1][0]>0.7 and four_one_dic[str1][0]>2 and li[7]>2:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0] - li[7]) / four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1]-li[8])/four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2] - li[9]) / four_one_dic[str1][2])
                #del lst[4]
                save_excel.append(lst)
                continue
        except:
            pass

        try:
            if (four_one_dic[str1][1]-li[8])/four_one_dic[str1][1]>0.7 and four_one_dic[str1][1]>5 and li[8]>5:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0] - li[7]) / four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1] - li[8]) / four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2] - li[9]) / four_one_dic[str1][2])
                #del lst[4]
                save_excel.append(lst)
                continue
        except:
            pass

        try:
            if (four_one_dic[str1][2]-li[9])/four_one_dic[str1][2]<-0.5 and li[9]>20 and four_one_dic[str1][2]>20:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0] - li[7]) / four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1] - li[8]) / four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2] - li[9]) / four_one_dic[str1][2])
                #del lst[4]
                save_excel.append(lst)
                continue
        except:
            pass

def fn2(filename1,filename2):
    pf1=pd.read_excel(filename1)
    pf2=pd.read_excel(filename2)
    time1=''
    for li in pf1.values:
        four_one_dic[li[1]+"-"+str(li[2])]=[li[6],li[7],li[8]]
        time1=li[0]
    for li in pf2.values:
        str1=li[1]+"-"+str(li[2])
        try:
            if (four_one_dic[str1][0]-li[6])/four_one_dic[str1][0]>0.7 and four_one_dic[str1][0]>2 and li[6]>2:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0]-li[6])/four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1]-li[7])/four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2]-li[8])/four_one_dic[str1][2])
                save_excel_f.append(lst)
                continue
        except:
            pass

        try:
            if (four_one_dic[str1][1]-li[7])/four_one_dic[str1][1]>0.7 and four_one_dic[str1][1]>2 and li[7]>2:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0] - li[6]) / four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1] - li[7]) / four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2] - li[8]) / four_one_dic[str1][2])
                save_excel_f.append(lst)
                continue
        except:
            pass

        try:
            if (four_one_dic[str1][2]-li[8])/four_one_dic[str1][2]<-0.5 and li[8]>10 and four_one_dic[str1][2]>10:
                lst=list(li)
                lst.insert(1,time1)
                lst.append(four_one_dic[str1][0])
                lst.append(four_one_dic[str1][1])
                lst.append(four_one_dic[str1][2])
                lst.append((four_one_dic[str1][0] - li[6]) / four_one_dic[str1][0])
                lst.append((four_one_dic[str1][1] - li[7]) / four_one_dic[str1][1])
                lst.append((four_one_dic[str1][2] - li[8]) / four_one_dic[str1][2])
                save_excel_f.append(lst)
                continue
        except:
            pass

def select_file():
    selected_files = []
    for i in range(8):
        root = tk.Tk()
        root.withdraw()
        # 设置窗口始终置顶
        root.attributes("-topmost", True)
        # 显示提示窗
        messagebox.showinfo("提示", f"请选择第 {lst[i]} 个文件")

        file_path = filedialog.askopenfilename()
        # 恢复窗口默认属性
        root.attributes("-topmost", False)

        if file_path:
            selected_files.append(file_path)
            print(f"你选择的文件是: {file_path}")
        else:
            print("你未选择任何文件。")

    return selected_files

def write_to_excel(file_path, data, headers):
    df = pd.DataFrame(data, columns=headers)
    try:
        # 若文件存在，写入数据前先清空
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, header=True)
    except FileNotFoundError:
        # 若文件不存在，直接写入数据
        df.to_excel(file_path, index=False, header=True)

    print(f"数据已成功写入 {file_path}")



def get_desktop_path():
    system = platform.system()
    if system == 'Windows':
        import ctypes.wintypes
        CSIDL_DESKTOPDIRECTORY = 0x0000
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
        return buf.value
    elif system == 'Darwin':
        return os.path.join(os.path.expanduser('~'), 'Desktop')
    elif system == 'Linux':
        return os.path.join(os.path.expanduser('~'), 'Desktop')
    else:
        return None

def select_directory():
    """使用Tkinter选择目录"""
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="请选择要遍历的文件夹")
    root.destroy()
    return folder_path


def list_files_recursively(path):
    """递归列出路径下的所有文件"""
    file_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            file_list.append(file_path)
    return file_list


if __name__=='__main__':
    # files=['/Users/reyted/Desktop/未命名文件夹 3/第一周泰山.xlsx', '/Users/reyted/Desktop/未命名文件夹 3/第一周欣网.xlsx',
    #  '/Users/reyted/Desktop/未命名文件夹 3/第二周泰山.xlsx', '/Users/reyted/Desktop/未命名文件夹 3/第二周欣网.xlsx',
    #  '/Users/reyted/Desktop/未命名文件夹 3/第一周泰山-f.xlsx', '/Users/reyted/Desktop/未命名文件夹 3/第一周欣网-f.xlsx',
    #  '/Users/reyted/Desktop/未命名文件夹 3/第二周泰山-f.xlsx', '/Users/reyted/Desktop/未命名文件夹 3/第二周欣网-f.xlsx']

    print("请选择要两周数据所存放路径的文件夹...")
    target_path = select_directory()

    if not target_path:  # 用户取消选择
        print("未选择文件夹，程序退出。")
    elif not os.path.exists(target_path):
        print("路径不存在！")
    else:
        all_files = list_files_recursively(target_path)
        #print("\n找到的文件列表:")
        for i, file in enumerate(all_files, 1):
            if '第一周' in file:
                if '4G新网故障预警' in file:
                    four_file_path1_xw=file
                if '4G泰山故障预警' in file:
                    four_file_path1_ts=file
                if '5G新网故障预警' in file:
                    five_file_path1_xw=file
                if '5G泰山故障预警' in file:
                    five_file_path1_ts=file
            if '第二周' in file:
                if '4G新网故障预警' in file:
                    four_file_path2_xw=file
                if '4G泰山故障预警' in file:
                    four_file_path2_ts=file
                if '5G新网故障预警' in file:
                    five_file_path2_xw=file
                if '5G泰山故障预警' in file:
                    five_file_path2_ts=file


        print('正在处理中。。。')


    #files=select_file()
    #print(files)
    #four_file_path1_ts='/Users/reyted/Desktop/未命名文件夹 3/第一周泰山.xlsx'
    #four_file_path1_ts=files[0]
    #four_file_path2_ts='/Users/reyted/Desktop/未命名文件夹 3/第二周泰山.xlsx'
    #four_file_path2_ts=files[2]

    #four_file_path1_xw = '/Users/reyted/Desktop/未命名文件夹 3/第一周欣网.xlsx'
    #four_file_path1_xw = files[1]
    #four_file_path2_xw = '/Users/reyted/Desktop/未命名文件夹 3/第二周欣网.xlsx'
    #four_file_path2_xw = files[3]

    #five_file_path1_ts = '/Users/reyted/Desktop/未命名文件夹 3/第一周泰山-f.xlsx'
    #five_file_path1_ts = files[4]
    #five_file_path2_ts = '/Users/reyted/Desktop/未命名文件夹 3/第二周泰山-f.xlsx'
    #five_file_path2_ts = files[6]

    #five_file_path1_xw = '/Users/reyted/Desktop/未命名文件夹 3/第一周欣网-f.xlsx'
    #five_file_path1_xw = files[5]
    #five_file_path2_xw = '/Users/reyted/Desktop/未命名文件夹 3/第二周欣网-f.xlsx'
    #five_file_path2_xw = files[7]

    fn1(four_file_path1_ts,four_file_path2_ts)
    fn1(four_file_path1_xw,four_file_path2_xw)

    fn2(five_file_path1_ts,five_file_path2_ts)
    fn2(five_file_path1_xw,five_file_path2_xw)
    #print(save_excel)
    #print(len(save_excel_f))

    #定义表头
    headers = [
        '基准数据日期','对比数据日期', '基站名称', 'eNodeB名称', '小区双工模式', '小区名称',
        '本地小区标识', '完整度', '总流量（GB）-3.0平台',
        '小区内的平均用户数', '上行PUSCH的RSRP低于-130的比例','上一周总流量（GB）','上一周平均用户数','上一周上行PUSCH的RSRP低于-130占比','总流量（GB）变化幅度','平均用户数变化幅度','上行PUSCH的RSRP变化幅度'
    ]

    desktop_path = get_desktop_path()
    current_date = datetime.date.today()

    # 获取月份和号数
    month = current_date.month
    day = current_date.day


    file_path = f"{desktop_path}/4G故障预警清单{month}月{day}号.xlsx"

    write_to_excel(file_path, save_excel, headers)

    # 定义表头
    headers1 = [
        '基准数据日期','对比数据日期', '基站名称', 'NR小区标识', 'gNodeB名称', '小区名称',
        '完整度', '总流量（GB）-1024版本',
        '平均用户数', '上行PUSCH的RSRP低于-130占比','上一周总流量（GB）','上一周平均用户数','上一周上行PUSCH的RSRP低于-130占比','总流量（GB）变化幅度','平均用户数变化幅度','上行PUSCH的RSRP变化幅度'
    ]

    file_path2 = f"{desktop_path}/5G故障预警清单{month}月{day}号.xlsx"

    write_to_excel(file_path2, save_excel_f, headers1)

    print('处理完成！')
    time.sleep(1.5)