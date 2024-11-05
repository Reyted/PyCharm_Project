import win32com.client
import os
from PIL import ImageGrab
import time
import win32gui
import win32con


class SimpleWordScreenshotter:
    def __init__(self):
        """
        初始化Word截图工具
        """
        self.word = win32com.client.Dispatch('Word.Application')
        self.word.Visible = False

    def bring_to_front(self, title_part):
        """
        将指定标题的窗口置顶并全屏显示
        """

        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if title_part in window_text:
                    # 将窗口置顶
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    # 全屏显示
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    windows.append(hwnd)
            return True

        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    def capture_area(self, word_path, output_path, left=0, top=0, width=500, height=500):
        """
        截取Word文档指定区域
        """
        try:
            # 获取文档的完整路径
            abs_path = os.path.abspath(word_path)

            # 打开Word文档
            doc = self.word.Documents.Open(abs_path)

            # 激活窗口并设置可见
            self.word.Visible = True
            self.word.ActiveWindow.View.Zoom.Percentage = 100

            # 最大化Word窗口
            self.word.ActiveWindow.WindowState = 1  # 1 = wdWindowStateMaximize

            # 等待窗口加载
            time.sleep(3)

            # 获取文档名称（用于查找窗口）
            doc_name = os.path.basename(word_path)

            # 将Word窗口置顶并全屏
            self.bring_to_front(doc_name)

            # 再次等待以确保窗口完全显示
            time.sleep(1)

            # 获取屏幕截图
            screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))

            # 保存截图
            screenshot.save(output_path)
            print(f"截图已保存到: {output_path}")

            # 关闭文档
            doc.Close()
            return True

        except Exception as e:
            print(f"截图时出错: {str(e)}")
            return False

        finally:
            # 确保关闭Word应用
            try:
                self.word.Quit()
            except:
                pass


def main(ind:int):
    # 创建截图工具实例
    screenshotter = SimpleWordScreenshotter()

    # 设置文件路径
    word_path = f"C:/Users/24253/Desktop/每月例行工作/123/模板/allfile/{ind}.docx"  # 替换为您的Word文档路径
    output_path = f"C:/Users/24253/Desktop/每月例行工作/123/image/资产标签/{ind}.png"  # 截图保存路径

    # 设置截图区域（根据实际需要调整坐标）
    left = 621  # 左边距
    top = 291  # 上边距
    width = 743  # 宽度
    height = 702  # 高度

    # 执行截图
    success = screenshotter.capture_area(
        word_path=word_path,
        output_path=output_path,
        left=left,
        top=top,
        width=width,
        height=height
    )

    if success:
        #print("截图完成")
        pass
    else:
        #print("截图失败")
        pass
