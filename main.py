import tkinter as tk
from tkinter import ttk
import threading

#不加这个执行会报错  发生错误: [WinError -2147221008] 尚未调用 CoInitialize。
# 安装依赖：pip install pywin32
import pythoncom

kamiVlaue = None
payCallBackValue = None
listenIntervalValue = None
startListenButton = None
endListenButton = None

#是否需要监听
islisten = None

#--------------------------------------------------

import re
import time
import uiautomation as automation

#这句还不能少，少了会报错comtypes.stream模块不存在
import comtypes.stream as comtypes
import requests

last_matched_info = None

#获取depth深度
def getDepth(control, depth):
    try:
        name = control.Name
        match = re.search(r'收款金额￥([\d.]+)', name)
        if match:
            return depth  # 找到匹配项，返回当前深度
        # 递归处理子控件，并检查返回值
        for child in control.GetChildren():
            found_depth = getDepth(child, depth + 4)
            if found_depth is not None:  # 如果子控件找到了匹配项，则返回该深度
                return found_depth
        return None  # 如果没有找到匹配项，返回None
    except Exception as e:
        print(f"处理控件时发生错误: {str(e)}")
        return None  # 发生错误时也返回None


def explore_control(control, depth, target_depth):
    global last_matched_info
    try:
        name = control.Name

        if name:
            if depth == target_depth:
                # 匹配收款金额信息
                match = re.search(r'收款金额￥([\d.]+)', name)
                if match:
                    global amount
                    amount = match.group(1)
                    last_matched_info = f"收款金额: ￥{amount}, "

                    # 匹配来自、到账时间信息
                    match = re.search(r'来自(.+?)到账时间', name)
                    global sender
                    sender = match.groups(1) if match else ('')
                    if sender:
                        last_matched_info += f"来自: {sender if sender else '未知'}, " if sender else ""

                    match = re.search(r'到账时间(.+?)备注', name)
                    global timestamp
                    timestamp = match.group(1) if match else ('')
                    if timestamp:
                        last_matched_info += f"到账时间: {timestamp if timestamp else '未知'}, " if timestamp else ""

                        # 匹配来自、到账时间信息
                match = re.search(r'共计￥([\d.]+)', name)
                global amountAll
                amountAll = match.group(1) if match else ('')
                if amountAll:
                    last_matched_info += f"收款金额总额: ￥{amountAll}, "
                # if match:
                #     global amountAll
                #     amountAll = match.group(1)
                #     last_matched_info += f"收款金额总额: ￥{amountAll}, "

                return
        # 递归处理子控件
        for child in control.GetChildren():
            explore_control(child, depth + 4, target_depth)
    except Exception as e:
        print(f"发生错误: {str(e)}")


def process_wechat_window(wechat_window, prev_info):
    global last_matched_info
    if wechat_window.Exists(0):
        # 假设 getDepth 函数已经定义好，并且 wechat_window 是一个有效的控件对象
        depth_of_match = getDepth(wechat_window, 0)

        explore_control(wechat_window, 0, depth_of_match)
        if last_matched_info and last_matched_info != prev_info:
            print(last_matched_info)
            print("-----------------------------------------------------------------")
            print("持续监听中...")
            print("-----------------------------------------------------------------")
            prev_info = last_matched_info

            # 向服务器发送请求
            send_http_request(last_matched_info, amount, amountAll, sender, timestamp)

    else:
        print("无法获取到窗口，请保持微信支付窗口显示...")
    return prev_info


def send_http_request(info, amount, amountAll, sender, timestamp):
    # 接收通知的Url
    server_url = payCallBackValue.get()
    try:

        params = {
            'amount': amount if amount is not None else '',
            'amountAll': amountAll if amountAll is not None else '',
            'sender': sender if sender is not None else '',
            'timestamp': timestamp if timestamp is not None else '',
        }
        # 将金额、来自、到账时间POST给服务器
        response = requests.post(server_url, json=params)
        # 通知成功
        # print("通知成功")
    except Exception as e:
        # 通知失败
        print(f"通知服务器失败...: {str(e)}")

def main():
    pythoncom.CoInitialize()

    global last_matched_info
    prev_info = None
    try:
        # 获取微信窗口
        wechat_window = automation.WindowControl(searchDepth=1, ClassName='ChatWnd')
        prev_info = process_wechat_window(wechat_window, prev_info)
    except Exception as e:
        print(f"发生错误: {str(e)}")

    while True:
        global islisten
        if not islisten: #是否监听标记为false时退出监听
            print("退出监听!")
            # 显示"开始监听按钮"按钮
            startListenButton.place(x=210, y=200)
            # 显示"终止监听"按钮
            endListenButton.place_forget()
            endListenButton['text'] = "终止监听" #改回来叫"终止监听"

            break

        try:
            # 持续监听微信窗口
            wechat_window = automation.WindowControl(searchDepth=1, ClassName='ChatWnd')
            prev_info = process_wechat_window(wechat_window, prev_info)
        except Exception as e:
            print(f"发生错误: {str(e)}")

        time.sleep(int(listenIntervalValue.get()))

    pythoncom.CoUninitialize()



# 窗口居中方法
def center_window(window, width, height):
    # 获取屏幕宽度和高度
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # 计算窗口的x和y坐标，使窗口居中
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # 设置窗口的大小和位置
    window.geometry(f'{width}x{height}+{x}+{y}')

#切换tab选项卡的时候触发的方法
def on_tab_change(event):
    tab_id = event.widget.select()
    print("Selected tab:", notebook.index(tab_id))

#初始化Frame作为基本配置选项卡的内容:
def initBaseConfigTab(notebook):
    frame = ttk.Frame(notebook, width=300, height=200)

    # # 创建"卡密"标签组件
    # kamiLabel = tk.Label(frame, text="卡密：", font=("Microsoft YaHei", 10), fg="black")
    # # 使用 place 布局管理器 设置绝对位置 放置标签
    # kamiLabel.place(x=80, y=20)

    # 创建 卡密 输入框并设置默认文本
    # global kamiVlaue
    # kamiVlaue = tk.StringVar()
    # kamiEntry = tk.Entry(frame,textvariable=kamiVlaue)
    # kamiEntry.insert(0, "")  # 设置默认文本
    # # 设置输入框宽度
    # kamiEntry.config(width=30)
    # kamiEntry.place(x=140,y=20)

    # 创建"支付回调"标签组件
    payCallbackLabel = tk.Label(frame, text="支付回调：", font=("Microsoft YaHei", 10), fg="black")
    # 使用 place 布局管理器 设置绝对位置 放置标签
    payCallbackLabel.place(x=52, y=70)

    # 创建 支付回调 输入框
    global payCallBackValue
    payCallBackValue = tk.StringVar()
    payCallBackEntry = tk.Entry(frame,textvariable=payCallBackValue)
    payCallBackEntry.insert(0, "")  # 设置默认文本
    # 设置输入框宽度
    payCallBackEntry.config(width=30)
    payCallBackEntry.place(x=140,y=70)

    # 创建"监听间隔(秒)"标签组件
    listenIntervalLabel = tk.Label(frame, text="监听间隔(秒)：", font=("Microsoft YaHei", 10), fg="black")
    # 使用 place 布局管理器 设置绝对位置 放置标签
    listenIntervalLabel.place(x=30, y=120)

    # 创建 支付回调 输入框并设置默认文本
    global listenIntervalValue
    listenIntervalValue = tk.StringVar()
    listenIntervalEntry = tk.Entry(frame,textvariable=listenIntervalValue)
    listenIntervalEntry.insert(0, "")  # 设置默认文本
    # 设置输入框宽度
    listenIntervalEntry.config(width=5)
    listenIntervalEntry.place(x=140,y=120)

    # 创建 开始监听 按钮
    global startListenButton
    startListenButton = tk.Button(frame, text="开始监听")
    startListenButton.bind("<Button-1>", start_listen_click)
    startListenButton.place(x=210, y=200)

    #创建 终止监听 按钮
    global endListenButton
    endListenButton = tk.Button(frame, text="终止监听")
    endListenButton.bind("<Button-1>", end_listen_click)
    # endListenButton.place(x=210, y=200) #一开始不显示

    return frame

# #初始化Frame作为支付日志选项卡的内容:
# def payLogs(notebook):
#     frame = ttk.Frame(notebook, width=300, height=200)
#
#     return frame



# 开始监听 按钮回调函数
def start_listen_click(event):
    global islisten
    if islisten:
        print("正在运行，请勿重复点击！")
        return

    islisten = True
    #隐藏"开始监听按钮"按钮 (如果是pack布局，就要用pack_forget方法)
    startListenButton.place_forget()

    #显示"终止监听"按钮
    endListenButton.place(x=210, y=200)

    print("卡密: ",kamiVlaue.get())
    print("回调地址: ", payCallBackValue.get())
    print("监听间隔: ", listenIntervalValue.get())



    #开启一个新线程来执行main方法
    thread1 = threading.Thread(target=main)

    thread1.start()


    print("开始监听了!")

#终止监听按钮事件处理函数
def end_listen_click(event):
    global islisten
    if not islisten:
        print("正在关闭，请勿重复点击！")
        return
    islisten = False #退出监听标记
    event.widget['text'] = "正在终止"



if __name__ == '__main__':
    root = tk.Tk()
    root.title("作者微信：cola521x")  # 设置窗口标题

    # 设置窗口大小并居中显示
    window_width = 500
    window_height = 400
    center_window(root, window_width, window_height)

    # 设置图标。图片格式
    # icon = tk.PhotoImage(file="./static/icon.png")
    # root.iconphoto(False, icon)

    # 创建Notebook组件
    notebook = ttk.Notebook(root)

    # 向Notebook添加选项卡
    notebook.add(initBaseConfigTab(notebook), text="微信支付回调插件")
    # notebook.add(payLogs(notebook), text="收款日志")

    # 布局Notebook
    notebook.pack(expand=True, fill=tk.BOTH)

    # 当用户更改选项卡时触发
    notebook.bind("<<NotebookTabChanged>>", on_tab_change)

    # 设置窗口不可调整大小on_tab_change
    root.resizable(False, False)

    root.mainloop()  # 进入事件循环













