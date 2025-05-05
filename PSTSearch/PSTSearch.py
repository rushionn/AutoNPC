import os
import fnmatch
import tkinter as tk
from tkinter import ttk
import threading
import webbrowser

class PSTFileSearcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PST Search")
        self.root.geometry("500x500")  # 調整大小以適應新功能
        self.root.configure(background='#172B4B')  # 背景顏色

        # 字體設置
        self.font_size = 14
        self.font = ('微軟正黑體', self.font_size)

        self.root.option_add("*Font", self.font)
        self.root.option_add("*Foreground", "#F5D07D")  # 全局前景色設置

        self.paused = False  # 初始化暫停標誌
        self.search_thread = None  # 初始化搜索線程

        self.create_widgets()

    def create_widgets(self):
        # 創建按鈕框架
        button_frame = tk.Frame(self.root, background='#172B4B')
        button_frame.pack(pady=5, anchor='w')

        # 按鈕設置
        self.btn_search = self.create_button(button_frame, "搜索 PST 文件", self.start_search)
        self.btn_pause = self.create_button(button_frame, "暫停", self.pause_search)
        self.btn_about = self.create_button(button_frame, "關於", self.show_about)

        # 日誌顯示框
        self.action_display = tk.Text(self.root, height=10, width=50, bg='black', fg='#F5D07D', font=self.font, bd=0)
        self.action_display.pack(pady=5, padx=10)

        # 進度條
        self.progress_bar = ttk.Progressbar(self.root, length=480, mode='determinate')
        self.progress_bar.pack(pady=5)

        # 檔案顯示框
        self.file_display = tk.Text(self.root, height=10, width=50, bg='#172B4B', fg='#F5D07D', font=self.font, bd=0)
        self.file_display.pack(pady=5, padx=10)

    def create_button(self, parent, text, command):
        """創建按鈕並設置顏色"""
        btn = ttk.Button(parent, text=text, command=command)
        btn.pack(side=tk.LEFT, padx=5)
        btn.configure(style="Custom.TButton")

        style = ttk.Style()
        style.configure("Custom.TButton", background='#172B4B', foreground='#000000')

        return btn

    def start_search(self):
        """啟動搜索線程"""
        self.paused = False  # 重置暫停狀態
        self.progress_bar['value'] = 0  # 重置進度條
        self.action_display.delete(1.0, tk.END)  # 清空日誌顯示
        self.file_display.delete(1.0, tk.END)    # 清空檔案顯示

        # 創建新線程以搜索 PST 文件
        self.search_thread = threading.Thread(target=self.search_pst_files)
        self.search_thread.start()

    def pause_search(self):
        """切換暫停和繼續"""
        self.paused = not self.paused
        self.btn_pause.config(text="繼續" if self.paused else "暫停")  # 切換按鈕文本

    def get_file_size(self, file_path):
        """獲取文件大小並格式化顯示"""
        size = os.path.getsize(file_path)
        if size >= 1024**3:  # GB
            return f"{round(size / (1024**3), 2)}GB"
        elif size >= 1024**2:  # MB
            return f"{round(size / (1024**2), 2)}mb"
        elif size >= 1024:  # KB
            return f"{round(size / 1024, 2)}kb"
        else:  # Bytes
            return f"{size}bytes"

    def search_pst_files(self):
        """搜索 .pst 文件並顯示結果"""
        home_dir = os.path.expanduser("~")
        search_dirs = [
            os.path.join(home_dir, 'Documents'),  # 文檔文件夾
            os.path.join(home_dir, 'Desktop'),     # 桌面
            os.path.join(home_dir, 'Downloads'),    # 下載
        ]

        pst_files = []  # 存儲找到的.pst文件路徑
        opened_folders = set()  # 用於記錄已開啟的資料夾

        total_files = 0  # 總文件數
        processed_files = 0  # 已處理文件數

        # 計算總文件數
        for search_dir in search_dirs:
            if os.path.exists(search_dir):
                for root, dirs, files in os.walk(search_dir):
                    total_files += len(fnmatch.filter(files, '*.pst'))

        # 開始搜索文件
        for search_dir in search_dirs:
            if os.path.exists(search_dir):
                for root, dirs, files in os.walk(search_dir):
                    for filename in fnmatch.filter(files, '*.pst'):
                        while self.paused:  # 如果被暫停，則等待
                            pass

                        pst_file_path = os.path.join(root, filename)
                        folder_path = os.path.dirname(pst_file_path)

                        # 添加唯一的文件路徑到列表
                        pst_files.append(pst_file_path)
                        processed_files += 1  # 更新已處理文件數
                        self.progress_bar['value'] = (processed_files / total_files) * 100  # 更新進度條

                        self.log_action(pst_file_path)  # 記錄完整路徑
                        
                        # 顯示檔案名稱及大小
                        file_size = self.get_file_size(pst_file_path)
                        self.display_file_info(filename, file_size)
                        
                        # 自動開啟資料夾
                        if folder_path not in opened_folders:
                            webbrowser.open(folder_path)  # 開啟資料夾
                            opened_folders.add(folder_path)  # 記錄已開啟的資料夾

        # 如果沒有找到文件，顯示消息
        if not pst_files:
            self.log_action("沒有找到任何.pst文件。")

    def log_action(self, message):
        """記錄操作，並在日誌框中顯示"""
        self.action_display.insert(tk.END, message + "\n")
        self.action_display.see(tk.END)

    def display_file_info(self, filename, file_size):
        """顯示檔案名稱和大小"""
        self.file_display.insert(tk.END, f"{filename} － {file_size}\n")
        self.file_display.see(tk.END)

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("關於")
        about_window.geometry("300x200")
        about_window.configure(background='#172B4B')


        author_label = tk.Label(about_window, text="v1.0", bg="#172B4B", fg="#F5D07D", font=self.font)
        author_label.pack(pady=10)
        author_label = tk.Label(about_window, text="Creat by: Lucien", bg="#172B4B", fg="#F5D07D", font=self.font)
        author_label.pack(pady=10)

        close_button = ttk.Button(about_window, text="關閉", command=about_window.destroy)
        close_button.pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = PSTFileSearcherApp(root)
    root.mainloop()