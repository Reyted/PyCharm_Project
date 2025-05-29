import pandas as pd


def read_excel_file(file_path, sheet_name=0):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print("文件读取成功！")
        return df
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None


# 使用示例
if __name__ == "__main__":
    # 替换为你的Excel文件路径
    excel_path1 = "C:/Users/24253/Desktop/新建文件夹/苏州华为爱立信5G工参整理17.xlsx"
    excel_path2 = "C:/Users/24253/Desktop/新建文件夹/小区合并.xlsx"

    # 读取Excel文件
    data = read_excel_file(excel_path)

    if data is not None:
        # 显示前5行数据
        print("\n文件内容预览：")
        print(data.head())

        # 获取所有列名
        print("\n列名：")
        print(data.columns.tolist())