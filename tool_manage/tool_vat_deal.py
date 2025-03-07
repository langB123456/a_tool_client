import deal_tax_invoice.deal_excel as deal_excel
import deal_tax_invoice.finally_deal_vat as finally_deal_vat
import deal_tax_invoice.merge_order_number as merge_order_number
import tkinter as tk


def main_use_flow():
    deal_excel.main_deal_excel()
    finally_deal_vat.main_finally_deal()
    merge_order_number.main_merge_order()


def tool_vat_deal_window(tool):
    tool_use_page = tk.Frame(tool, height=884, width=1000, background='white', highlightcolor='black', relief='ridge')
    # 右边功能窗口
    r1 = tk.Label(tool_use_page, text='处理税金单', background='white', justify='left')
    r1.place(x=0, y=0)

    # 说明内容
    describe = ('说明：该功能用于本地处理天图税金单，执行前须知：\n'
                '1、执行中不要操作键盘和鼠标避免转发内容出错')

    d3 = tk.Label(tool_use_page, text=describe, background='white', justify='left')
    d3.place(x=0, y=25)

    # 执行按钮
    implement_button = tk.Button(tool_use_page,
                                 text='执行',
                                 background='#AFEEEE',
                                 command=main_use_flow)
    implement_button.place(x=2, y=126)

    tool_use_page.place(x=220, y=20)
    return tool_use_page
