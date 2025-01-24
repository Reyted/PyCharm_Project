import rarfile
import os


def unrar_file(rar_path, output_dir):
    """
    此函数用于解压 rar 文件
    :param rar_path: 要解压的 rar 文件的完整路径
    :param output_dir: 解压文件的目标目录
    """
    # 检查输出目录是否存在，如果不存在则创建
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # 打开 rar 文件
    with rarfile.RarFile(rar_path) as rf:
        # 解压所有文件到目标目录
        rf.extractall(output_dir)





if __name__ == "__main__":
    # 这里修改为你的 rar 文件的实际路径
    rar_path = 'C:/Users/24253/Desktop/工作内容/25年每月例行工作/1月工作内容/指标提取/14/5G服务小区覆盖分析--小区详情数据.rar'
    # 这里修改为你想要解压到的目录路径
    output_dir = 'C:/Users/24253/Desktop/工作内容/25年每月例行工作/1月工作内容/指标提取/temp/'
    unrar_file(rar_path, output_dir)