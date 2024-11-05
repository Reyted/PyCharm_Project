import os


def rename_file(old_name, new_name):
    os.rename(old_name, new_name)


def traverse_folder(folder_path):
    ind = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            ind+=1
            file_path = os.path.join(root, file)
            rename_file(file_path,f"C:/Users/24253/Desktop/每月例行工作/123/资产标签/{ind}.jpg")

traverse_folder("C:/Users/24253/Desktop/每月例行工作/123/资产标签")