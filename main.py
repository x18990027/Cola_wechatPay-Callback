import json  # 新增导入
import tkinter as tk
from tkinter import ttk, messagebox
import threading

#不加这个执行会报错  发生错误: [WinError -2147221008] 尚未调用 CoInitialize。
# 安装依赖：pip install pywin32
import pythoncom

CONFIG_FILE = "config.json"
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
    global last_matched_info, amount, amountAll, sender, timestamp
    try:
        name = control.Name
        if name and depth == target_depth:
            # 优化金额匹配逻辑
            amount_match = re.search(r'收款金额￥([\d.]+)', name)
            if amount_match:
                amount = amount_match.group(1)
                last_matched_info = f"收款金额: ￥{amount}, "

                # 优化发款人匹配
                sender_match = re.search(r'来自\s*([^\s]+)', name)
                sender = sender_match.group(1) if sender_match else ''
                if sender:
                    last_matched_info += f"来自: {sender}, "

                # 优化到账时间匹配
                time_match = re.search(r'到账时间[:：]\s*([^\s]+)', name)
                timestamp = time_match.group(1) if time_match else ''
                if timestamp:
                    last_matched_info += f"到账时间: {timestamp}, "

                # 优化总额匹配
                total_match = re.search(r'共计￥([\d.]+)', name)
                if not total_match:  # 兼容不同文案
                    total_match = re.search(r'收款金额总额.*?￥([\d.]+)', name)
                amountAll = total_match.group(1) if total_match else ''
                if amountAll:
                    last_matched_info += f"收款金额总额: ￥{amountAll}, "
                return
        # 递归处理子控件
        for child in control.GetChildren():
            explore_control(child, depth + 4, target_depth)
    except Exception as e:
        print(f"发生错误: {str(e)}")
def process_wechat_window(wechat_window, prev_info):
    global last_matched_info, amount, amountAll, sender, timestamp
    if wechat_window.Exists(0):
        depth_of_match = getDepth(wechat_window, 0)
        explore_control(wechat_window, 0, depth_of_match)
        if last_matched_info and last_matched_info != prev_info:
            # 添加支付信息日志
            log_message("💰 监听到支付信息：")
            log_message(f"▸ {last_matched_info.replace(', ', '\n▸ ')}")
            log_message("-"*30)
            prev_info = last_matched_info

            # 添加请求开始日志
            log_message("📡 正在请求回调接口...")
            log_message(f"▪ 收款金额：￥{amount}")
            log_message(f"▪ 收款总额：￥{amountAll}")
            log_message("——— 请求回调 ———")
            send_http_request(last_matched_info, amount, amountAll, sender, timestamp)

    else:
        print("无法获取到窗口，请保持微信支付窗口显示...")
    return prev_info


def send_http_request(info, amount, amountAll, sender, timestamp):
    server_url = payCallBackValue.get()
    try:
        params = {
            'amount': amount if amount is not None else '',
            'amountAll': amountAll if amountAll is not None else '',
            'sender': sender if sender is not None else '',
            'timestamp': timestamp if timestamp is not None else '',
        }
        response = requests.post(server_url, json=params)
        response.raise_for_status()
        # 添加成功日志
        log_message(f"✅ 回调成功（状态码 {response.status_code}）")
        log_message(f"📤 发送参数：{params}")
        log_message("-"*30 + "\n")
    except Exception as e:
        # 添加失败日志
        log_message(f"❌ 回调失败：{str(e)}")
        log_message(f"⚠️ 失败参数：{params}")
        log_message("🛑 请检查：1.网络连接 2.服务器状态 3.接口协议")
        log_message("-"*30 + "\n")

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
    frame = ttk.Frame(notebook)

    # 配置项部分
    config_frame = ttk.Frame(frame)
    config_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nw")

    # 支付回调地址（添加输入监听）
    ttk.Label(config_frame, text="支付回调地址：").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    pay_entry = ttk.Entry(config_frame, textvariable=payCallBackValue, width=40)
    pay_entry.grid(row=0, column=1, padx=5, pady=5)
    payCallBackValue.trace_add("write", lambda *_: save_config())  # 输入实时保存

    # 监听间隔（添加输入监听）
    ttk.Label(config_frame, text="监听间隔(秒)：").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    listenIntervalEntry = ttk.Entry(config_frame, textvariable=listenIntervalValue, width=8)
    listenIntervalEntry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    listenIntervalValue.trace_add("write", lambda *_: save_config())  # 输入实时保存

    # 日志面板（配置项下方）
    log_frame = ttk.LabelFrame(frame, text="运行日志", padding=5)
    log_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

    global log_text
    log_text = tk.Text(log_frame,
                     height=10,
                     state='disabled',
                     bg="black",    # 黑色背景
                     fg="white",    # 白色文字
                     insertbackground="white")  # 光标颜色
    log_text.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    scrollbar.pack(side="right", fill="y")
    log_text.configure(yscrollcommand=scrollbar.set)

    # 按钮容器
    btn_frame = ttk.Frame(config_frame)
    btn_frame.grid(row=2, column=0, columnspan=2, pady=10)

    global controlButton

    # 创建统一按钮
    controlButton = tk.Button(btn_frame, text="开始监听",
                            command=toggle_listen,
                            bg="#87CEFA",  # 初始蓝色背景
                            fg="black",    # 黑色文字
                            activebackground="#FF4500")  # 点击时的红色
    controlButton.pack(side="left", padx=5)

    # 配置框架自适应
    frame.grid_rowconfigure(1, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    return frame



#终止监听按钮事件处理函数
def end_listen_click(event):
    global islisten
    islisten = False
    # 移除原有的文本修改操作
    if event and event.widget:  # 添加空值检查
        event.widget['text'] = "正在终止"



def toggle_listen():
    global islisten
    if not islisten:
        # 启动监听
        islisten = True
        controlButton.config(text="终止监听", bg="#FF4500")
        # 直接调用启动逻辑
        log_message("✅ 启动参数")
        log_message(f"• 回调地址: {payCallBackValue.get()}")
        log_message(f"• 监听间隔: {listenIntervalValue.get()}秒")
        log_message("🌟" * 30)
        log_message("✅ 欢迎使用支付支付插件 by cola!")
        log_message("PS：启动时会监听到最新一条的收款记录并发送回调，请忽略！\n")
        log_message("🚀 开始监听微信支付通知...")
        log_message("🌟" * 30)
        thread1 = threading.Thread(target=main)
        thread1.start()
    else:
        # 终止监听
        islisten = False
        controlButton.config(text="开始监听", bg="#87CEFA")
        log_message("🛑 监听已终止")
        controlButton['state'] = 'normal'  # 确保按钮状态恢复


def load_config():
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            payCallBackValue.set(config.get('callback_url', ''))
            listenIntervalValue.set(config.get('interval', '1'))
    except (FileNotFoundError, json.JSONDecodeError):
        payCallBackValue.set('')
        listenIntervalValue.set('2')

def save_config():
    config = {
        'callback_url': payCallBackValue.get(),
        'interval': listenIntervalValue.get()
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def log_message(message):
    log_text.configure(state='normal')
    log_text.insert(tk.END, f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    log_text.configure(state='disabled')
    log_text.see(tk.END)  # 自动滚动到底部

if __name__ == '__main__':
    root = tk.Tk()

    # Initialize StringVars
    payCallBackValue = tk.StringVar()
    listenIntervalValue = tk.StringVar()

    # Then load config
    load_config()

    # 绑定窗口关闭事件
    root.protocol("WM_DELETE_WINDOW", lambda: [save_config(), root.destroy()])
    root.title("作者微信：cola521x")  # 设置窗口标题

    # 设置窗口大小并居中显示
    window_width = 600  # 调整为600宽度
    window_height = 500  # 增加高度
    center_window(root, window_width, window_height)

    # 设置图标。图片格式
    # icon = tk.PhotoImage(file="./static/icon.png")
    # root.iconphoto(False, icon)

    # 创建Notebook组件
    notebook = ttk.Notebook(root)

    # 向Notebook添加选项卡
    notebook.add(initBaseConfigTab(notebook), text="微信支付回调插件 V1.1.0")
    # notebook.add(payLogs(notebook), text="收款日志")

    # 布局Notebook
    notebook.pack(expand=True, fill=tk.BOTH)

    # 当用户更改选项卡时触发
    notebook.bind("<<NotebookTabChanged>>", on_tab_change)

    # 设置窗口不可调整大小on_tab_change
    root.resizable(False, False)

    root.mainloop()  # 进入事件循环













