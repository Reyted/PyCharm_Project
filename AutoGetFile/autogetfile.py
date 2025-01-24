import pyautogui
import time
import os


def open_chrome():
    try:
        # Chrome浏览器的默认安装路径
        chrome_path = 'C:/Program Files/SBrowser/ESBrowser.exe'
        chrome_path = 'C:/Program Files\SBrowser\ESBrowser.exe'

        # 可以添加启动参数，比如要打开的网址
        url = 'https://www.google.com'

        # 使用os.system启动Chrome
        os.system(f'"{chrome_path}"')
        print("Chrome浏览器已启动")

    except Exception as e:
        print(f"发生错误: {str(e)}")




def move_and_click(x, y):
    pyautogui.PAUSE = 0.1
    pyautogui.FAILSAFE = True

    try:
        # 移动并点击
        pyautogui.moveTo(x, y, duration=0.5)
        print(f"鼠标已移动到坐标: ({x}, {y})")
        pyautogui.click()
        print("已执行单击操作")

    except pyautogui.FailSafeException:
        print("触发了防故障保护")
    except Exception as e:
        print(f"发生错误: {str(e)}")


def get_current_position():
    x, y = pyautogui.position()
    print(f"当前鼠标位置: ({x}, {y})")
    return x, y


def main():
    while True:
        print("\n请选择操作：")
        print("1. 移动鼠标到指定坐标并点击")
        print("2. 获取当前鼠标位置")
        print("3. 退出程序")

        choice = input("请输入选项（1-3）: ")

        if choice == '1':
            try:
                x = int(input("请输入X坐标: "))
                y = int(input("请输入Y坐标: "))
                print("3秒后开始移动并点击...")
                time.sleep(3)
                move_and_click(x, y)
            except ValueError:
                print("请输入有效的数字坐标！")

        elif choice == '2':
            get_current_position()

        elif choice == '3':
            print("程序退出")
            break

        else:
            print("无效的选项，请重新选择")


if __name__ == "__main__":
    open_chrome()
    #main()