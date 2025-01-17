import tkinter as tk
from tkinter import ttk
from tool_manage import tool_forward_everyone, tool_default, tool_vat_deal, tool_vat_upload


# 组件触发事件
def on_tree_click(event, tool):
    # 获取被点击的项 ID
    item_id = bar_list.identify_row(event.y)

    if not item_id:
        return

    # 获取该项的文本（即名称）
    item_text = bar_list.item(item_id, 'text')
    print(f"点击的项名称: {item_text}")

    # 动态调用 frame_change 函数并传递工具和点击项名称
    if item_text == '企微转发':
        first_frame = tool_forward_everyone.tool_forward_everyone_window(tool)
        frame_change(tool, first_frame)
    elif item_text == '税金单处理':
        first_frame = tool_vat_deal.tool_vat_deal_window(tool)
        frame_change(tool, first_frame)
    elif item_text == '税金单上传':
        first_frame = tool_vat_upload.tool_vat_upload_window(tool)
        frame_change(tool, first_frame)
    else:
        default_frame = tool_default.tool_default_window(tool)
        frame_change(tool, default_frame)


# frame工具页切换
def frame_change(tool, frame_name):
    # 查找所有子 Frame 并切换显示
    all_frames = [child for child in tool.winfo_children() if isinstance(child, tk.Frame)]
    found = False

    for single_frame in all_frames:
        if single_frame.winfo_exists():
            if single_frame.winfo_class() == 'Frame' and hasattr(single_frame, 'name') and single_frame.name == frame_name:
                single_frame.pack(expand=1)
                found = True
            else:
                single_frame.forget()

    if not found:
        return


def tool_left_window(tool):
    global bar_list  # 将 bar_list 设为全局变量以便可以在其他地方访问

    left_label = tk.Label(tool)
    left_label.place(x=0, y=40)

    # 创建树形列表
    bar_list = ttk.Treeview(left_label, height=43, show='tree')  # height设置tree的上下距离的高度
    bar_list.place(x=2, y=30)

    # 顶部操作栏
    bar_btn = tk.Button(text='分组')
    bar_btn.place(x=10, y=3)

    tool_list = tk.Button(text='添加工具')
    tool_list.place(x=45, y=3)

    into_bar_btn = tk.Button(text='添加分组')
    into_bar_btn.place(x=105, y=3)

    bar_list_manage = {
        '通用功能': ['企微转发'],
        '天图系统': ['税金单处理', '税金单上传'],
        '安速系统': []
    }

    for first_tree in bar_list_manage.__reversed__():
        if first_tree == [] or first_tree == [''] or first_tree is None:
            continue
        bar_tree_first = bar_list.insert('', len(bar_list_manage), first_tree, text=first_tree, values=('1'))
        for second_tree in bar_list_manage[first_tree]:
            bar_tree_second = bar_list.insert(bar_tree_first, len(bar_list_manage[first_tree]), second_tree, text=second_tree, values=('2'))

    bar_list.bind('<Button-1>', lambda event: on_tree_click(event, tool))
    bar_list.pack(expand=1)




