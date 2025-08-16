import os
import shutil
import platform
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("win10、11专用：ie浏览器回归软件")
        self.geometry("500x452")

        
        try:
            self.iconbitmap(self.get_resource_path("src/icon.ico"))
        except tk.TclError:
            print("图标 'src/icon.ico' 未找到或格式不正确。")

        self.grid_columnconfigure(0, weight=1)

        # 创建密钥验证界面
        self.create_key_verification()
        
        # 延迟创建主界面，直到密钥验证通过
        self.main_widgets_created = False

        if platform.system() != "Windows":
            messagebox.showerror("系统不兼容", "此软件仅支持 Windows 系统。")
            self.destroy()

    def create_widgets(self):
        container = ttk.Frame(self, padding=20)
        container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        container.grid_columnconfigure(0, weight=1)

        # --- IE 环境监测 ---
        detect_frame = ttk.Labelframe(container, text="第一步：检测IE环境", padding=15)
        detect_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        detect_frame.grid_columnconfigure(0, weight=1)

        self.detect_status_label = ttk.Label(detect_frame, text="请点击按钮开始检测...", bootstyle="info")
        self.detect_status_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        detect_button = ttk.Button(detect_frame, text="监测IE环境", command=self.check_ie_existence, bootstyle="primary")
        detect_button.grid(row=1, column=0, sticky=tk.W)

        # --- 快捷方式复制 ---
        copy_frame = ttk.Labelframe(container, text="第二步：创建IE快捷方式", padding=15)
        copy_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        copy_frame.grid_columnconfigure(0, weight=1)
        
        copy_intro_label = ttk.Label(copy_frame, text="选择一个位置来创建IE浏览器快捷方式。")
        copy_intro_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        recommend_label = ttk.Label(copy_frame, text="(推荐放在桌面)", bootstyle="secondary")
        recommend_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        self.path_entry = ttk.Entry(copy_frame, state="readonly")
        self.path_entry.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        browse_button = ttk.Button(copy_frame, text="选择路径", command=self.browse_path, bootstyle="secondary")
        browse_button.grid(row=2, column=1, sticky=tk.E)

        self.copy_button = ttk.Button(copy_frame, text="创建快捷方式", command=self.copy_shortcut, state="disabled")
        self.copy_button.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))

        # --- 底部信息 ---
        footer_frame = ttk.Frame(container)
        footer_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(20, 0))
        
        os_info = f"系统: {platform.system()} {platform.release()} (版本: {platform.version()})"
        os_label = ttk.Label(footer_frame, text=os_info, bootstyle="secondary")
        os_label.grid(row=0, column=0, sticky=tk.W)
        
        author_label = ttk.Label(footer_frame, text="等风时-制作", bootstyle="secondary")
        author_label.grid(row=0, column=1, sticky=tk.E)
        
        footer_frame.grid_columnconfigure(0, weight=1)


    def check_ie_existence(self):
        ie_path = r"C:\Program Files (x86)\Internet Explorer\iexplore.exe"
        if os.path.exists(ie_path):
            self.detect_status_label.config(text="检测通过：IE环境已安装，可以使用本软件。", bootstyle="success")
            messagebox.showinfo("检测成功", "您的系统已安装IE环境，可以继续创建快捷方式。")
            self.copy_button.config(state="normal")
        else:
            self.detect_status_label.config(text="检测失败：IE环境未安装。", bootstyle="danger")
            result = messagebox.askyesno("检测失败", "IE环境没有安装，是否下载IE11离线安装包？")
            if result:
                webbrowser.open("https://raw-cdn.gitcode.com/open-source-toolkit/9442d/blobs/b522a377801b7059c293ea2b4a9424155f5b5e15/IE11(x32,X64)%E7%A6%BB%E7%BA%BF%E4%B8%80%E9%94%AE%E7%9B%B4%E8%A3%85%E5%AE%89%E8%A3%85%E5%8C%85%EF%BC%88%E5%8C%85%E5%90%AB%E5%85%A8%E9%83%A8%E8%A1%A5%E4%B8%81%EF%BC%89.zip")
            self.copy_button.config(state="disabled")

    def browse_path(self):
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        directory = filedialog.askdirectory(initialdir=desktop_path, title="请选择保存快捷方式的文件夹")
        if directory:
            self.path_entry.config(state="normal")
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, directory)
            self.path_entry.config(state="readonly")

    def create_key_verification(self):
        container = ttk.Frame(self, padding=20)
        container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        container.grid_columnconfigure(0, weight=1)

        key_frame = ttk.Labelframe(container, text="密钥验证", padding=15)
        key_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        key_frame.grid_columnconfigure(0, weight=1)

        ttk.Label(key_frame, text="请输入使用密钥:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        # 密码输入框和显示复选框容器
        password_container = ttk.Frame(key_frame)
        password_container.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        password_container.grid_columnconfigure(0, weight=1)

        self.key_entry = ttk.Entry(password_container, show="*")
        self.key_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        self.show_password_var = tk.BooleanVar()
        show_password_check = ttk.Checkbutton(password_container, text="显示密码", variable=self.show_password_var, command=self.toggle_password_visibility)
        show_password_check.grid(row=0, column=1, sticky=tk.E)

        verify_button = ttk.Button(key_frame, text="验证", command=self.verify_key, bootstyle="primary")
        verify_button.grid(row=2, column=0, sticky=tk.E)

    def verify_key(self):
        key = self.key_entry.get()
        if key == "0d000721":
            # 清除密钥验证界面
            for widget in self.winfo_children():
                widget.destroy()
            # 创建主界面
            self.create_widgets()
            self.main_widgets_created = True
        else:
            messagebox.showerror("错误", "密钥不正确，请咨询卖家")
            self.key_entry.delete(0, tk.END)

    def toggle_password_visibility(self):
        """切换密码显示/隐藏"""
        if self.show_password_var.get():
            self.key_entry.config(show="")
        else:
            self.key_entry.config(show="*")

    def get_resource_path(self, relative_path):
        """获取资源的绝对路径，兼容打包后的exe"""
        try:
            # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def copy_shortcut(self):
        source_path = self.get_resource_path("src/IE浏览器专用.lnk")
        destination_folder = self.path_entry.get()

        if not destination_folder:
            messagebox.showerror("错误", "请先选择一个目标路径。")
            return

        if not os.path.exists(source_path):
            messagebox.showerror("错误", f"源文件 '{source_path}' 不存在。")
            return

        destination_path = os.path.join(destination_folder, "IE浏览器专用.lnk")

        try:
            shutil.copy(source_path, destination_path)
            messagebox.showinfo("成功", f"快捷方式已成功创建到:\n{destination_path}")
        except Exception as e:
            messagebox.showerror("复制失败", f"创建快捷方式时发生错误: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
