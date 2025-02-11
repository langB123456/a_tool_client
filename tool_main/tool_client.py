import tkinter as tk
import tool_left
import pyautogui
import os


def get_window_size():
    size = pyautogui.size()
    return size


def tool_window():
    # 获取当前路径
    now_path = os.getcwd()
    tool = tk.Tk()
    tool.title('use your tool')
    tool.iconbitmap('../tool_picture/dog_icon.ico')
    tool.wm_iconbitmap('../tool_picture/dog_icon.ico')
    tool.iconphoto(False, tk.PhotoImage(file='../tool_picture/dog_icon.ico'))

    sw = tool.winfo_screenwidth()
    sh = tool.winfo_screenheight()

    w = get_window_size()[0] // 2
    h = get_window_size()[1] // 3 * 2
    x = 200
    y = (sh - h) // 2
    tool.geometry(f'{w}x{h}+{x}+{y}')
    tool.resizable(False, False)

    # 左边功能选择 需要在这里配置切换功能页
    tool_left.tool_left_window(tool)

    tool.mainloop()


# pyinstaller -p D:\ForwardEveryone\.venv\Lib\site-packages -i D:\ForwardEveryone\dog_icon.ico forward_demo.py
if __name__ == '__main__':
    tool_window()