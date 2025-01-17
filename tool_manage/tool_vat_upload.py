import tkinter as tk
from tkinter import ttk


def main_use_flow(event):
    print('hello,world')


def tool_vat_upload_window(tool):
    h = 884
    w = 1000
    tool_use_page = tk.Frame(tool, height=h, width=w, background='white')

    # 右边功能窗口
    # 顶部标题显示
    r1 = tk.Label(tool_use_page, background='blue', justify='left', width=142)
    r1.place(x=0, y=0)

    t1 = tk.Label(r1, text='税金单上传', fg='white', background='blue', justify='left')
    t1.place(x=0, y=0)

    tool_use_page.place(x=220, y=20)
    return tool_use_page
