#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最简单登录窗口（无美化，仅核心功能）
"""

import tkinter as tk
from tkinter import messagebox
import hashlib
import requests
import threading

DEFAULT_HOST = "http://21.1.2.246:18080"
DEFAULT_USERNAME = "admin"


class LoginWindow:
    def __init__(self, on_success_callback):
        self.on_success = on_success_callback
        self.root = tk.Tk()
        self.root.title("")
        self.root.geometry("400x180")
        self.root.resizable(False, False)

        # 居中显示
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 400) // 2
        y = (self.root.winfo_screenheight() - 250) // 2
        self.root.geometry(f"+{x}+{y}")

        # 变量
        self.server_host = tk.StringVar(value=DEFAULT_HOST)
        self.username_var = tk.StringVar(value=DEFAULT_USERNAME)
        self.password_var = tk.StringVar(value="")

        self.create_widgets()

    def create_widgets(self):
        # 主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        # 服务器地址
        tk.Label(main_frame, text="服务器地址:").grid(row=0, column=0, sticky="w", pady=5)
        self.server_entry = tk.Entry(main_frame, textvariable=self.server_host, width=40)
        self.server_entry.grid(row=0, column=1, pady=5, padx=5)

        # 用户名
        tk.Label(main_frame, text="用户名:").grid(row=1, column=0, sticky="w", pady=5)
        self.user_entry = tk.Entry(main_frame, textvariable=self.username_var, width=40)
        self.user_entry.grid(row=1, column=1, pady=5, padx=5)

        # 密码
        tk.Label(main_frame, text="密码:").grid(row=2, column=0, sticky="w", pady=5)
        self.pass_entry = tk.Entry(main_frame, textvariable=self.password_var, show="*", width=40)
        self.pass_entry.grid(row=2, column=1, pady=5, padx=5)
        self.pass_entry.bind("<Return>", lambda e: self.do_login())

        # 登录按钮
        self.login_btn = tk.Button(main_frame, text="登录", command=self.do_login, width=15)
        self.login_btn.grid(row=3, column=0, columnspan=2, pady=15)

        # 状态标签
        self.status_label = tk.Label(main_frame, text="", fg="red")
        self.status_label.grid(row=4, column=0, columnspan=2)

        # 使列可扩展
        main_frame.columnconfigure(1, weight=1)

    @staticmethod
    def md5(text):
        return hashlib.md5(text.encode('utf-8')).hexdigest()

    def do_login(self):
        host = self.server_host.get().strip().rstrip('/')
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()

        if not all([host, username, password]):
            self.status_label.config(text="请填写完整信息")
            return

        # 自动补全协议
        if not host.startswith(('http://', 'https://')):
            host = 'http://' + host

        self.login_btn.config(state=tk.DISABLED, text="登录中...")
        self.status_label.config(text="正在连接...", fg="blue")

        def task():
            try:
                url = f"{host}/api/user/login"
                resp = requests.get(url, params={
                    "username": username,
                    "password": self.md5(password)
                }, timeout=15)

                if resp.status_code == 200:
                    data = resp.json()
                    token = data.get("accessToken") or resp.headers.get("access-token")
                    if token:
                        self.root.after(0, lambda: self.on_success(token, data, host))
                    else:
                        self.root.after(0, lambda: self.fail("未获取到token"))
                else:
                    msg = resp.json().get("msg", f"HTTP {resp.status_code}")
                    self.root.after(0, lambda m=msg: self.fail(m))
            except Exception as e:
                self.root.after(0, lambda e=e: self.fail(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def fail(self, msg):
        self.login_btn.config(state=tk.NORMAL, text="登录")
        self.status_label.config(text=f"错误: {msg}", fg="red")


if __name__ == "__main__":
    def test_callback(token, data, host):
        print(f"登录成功！Token: {token[:20]}...")
        messagebox.showinfo("成功", "登录成功")

    win = LoginWindow(test_callback)
    win.root.mainloop()