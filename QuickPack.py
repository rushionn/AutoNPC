import os
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import locale
import webbrowser  # 引入webbrowser以便在新的窗口中打開超連結

class ToolTip:
    """懸浮提示框類"""
    def __init__(self, widget, text, text_ja=None, font_size=10, text_color="black"):
        self.widget = widget
        self.text = text
        self.text_ja = text_ja
        self.font_size = font_size
        self.text_color = text_color
        self.tip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tip_window, text=self.get_current_tooltip_text(), bg="lightyellow", fg=self.text_color, font=('微軟正黑體', self.font_size))
        label.pack()

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def get_current_tooltip_text(self):
        """根據當前語言返回提示文本"""
        return self.text_ja if app.current_language == "ja" else self.text

class FileCompressorApp:
    def __init__(self, root):
        self.root = root

        # 提示框控制變數
        self.tooltip_font_size = 10
        self.tooltip_text_color = "black"

        # 多語言支持映射
        self.language_mapping = {
            "app_title": {
                "zh": "QuickPack v1.0", 
                "ja": "QuickPack v1.0"
            },
            "check": {
                "zh": "檢查", 
                "ja": "チェック"
            },
            "pack": {
                "zh": "打包", 
                "ja": "パック"
            },
            "pause": {
                "zh": "暫停", 
                "ja": "一時停止"
            },
            "delete": {
                "zh": "刪除", 
                "ja": "削除"
            },
            "split": {
                "zh": "分割", 
                "ja": "分割"
            },
            "checking_result": {
                "zh": "檢測結果：", 
                "ja": "チェック結果："
            },
            "compression_complete": {
                "zh": "所選資料夾已經壓縮完成。", 
                "ja": "選択したフォルダが圧縮されました。"
            },
            "delete_msg": {
                "zh": "沒有找到檔案，未執行動作", 
                "ja": "ファイルが見つかりませんでした。アクションを実行しませんでした。"
            },
            "delete_folder_msg": {
                "zh": "刪除已建立的壓縮檔: ", 
                "ja": "圧縮ファイルを削除する: "
            },
            "folder_deleted": {
                "zh": "已經刪除建立的壓縮檔: ", 
                "ja": "作成された圧縮ファイルは削除されました: "
            },
            "folder_not_found": {
                "zh": "檔案不存在: ", 
                "ja": "ファイルが存在しません: "
            },
            "toggle_language": {
                "zh": "日本語へ切換", 
                "ja": "切換為中文"
            },
            "about": {
                "zh": "詳細", 
                "ja": "詳細"
            },
            "author": {
                "zh": "Created by Lucien", 
                "ja": "Created by Lucien"
            },
            "description": {
                "zh": "v.1.0", 
                "ja": "v.1.0"
            },
            "website": {
                "zh": "IS線上申請 (請點擊)", 
                "ja": "ISオンライン申請（クリック）"
            },
            "contact": {
                "zh": "透過信箱聯絡我們(請點擊)", 
                "ja": "メールでお問い合わせください（クリック）"
            }
        }

        self.tooltip_mapping = {
            "check": {
                "zh": "檢查所選資料夾的大小",
                "ja": "選択したフォルダのサイズを確認します"
            },
            "pack": {
                "zh": "開始壓縮所選資料夾",
                "ja": "選択したフォルダを圧縮します"
            },
            "pause": {
                "zh": "暫停或繼續壓縮",
                "ja": "圧縮プロセスを一時停止または再開します"
            },
            "delete": {
                "zh": "刪除生成的壓縮檔",
                "ja": "生成された圧縮ファイルを削除します"
            },
            "split": {
                "zh": "壓縮檔超過40GB時自動分割",
                "ja": "圧縮ファイルが40GBを超えた場合、自動で分割します"
            }
        }

        self.checkbox_language_mapping = {
            "Desktop": "Desktop",
            "Downloads": "Downloads",
            "Documents": "Documents",
            "Pictures": "Pictures",
            "Favorites": "Favorites",
        }

        self.current_language = "zh"  # 預設語言
        self.detect_language()

        self.root.title(self.get_translation("app_title"))
        self.root.geometry("500x400")
        self.root.configure(background='#172B4B')  # 設置背景顏色

        # 字體設置
        font_size = 14  
        font_label_size = 10  
        font_large = ('微軟正黑體', font_size)  
        font_label = ('微軟正黑體', font_label_size)  

        self.root.option_add("*TButton*Font", font_large)
        self.root.option_add("*Label*Font", font_large)
        self.root.option_add("*Foreground", "#F5D07D")  # 設置全局前景色

        self.create_widgets()

        self.is_running = False
        self.is_paused = False
        self.worker_thread = None

    def create_widgets(self):
        # 按鈕區域
        button_frame = tk.Frame(self.root, background='#172B4B')
        button_frame.pack(pady=5, anchor='w')  

        button_width = 10  

        # 顏色控制
        button_bg_color = '#172B4B'
        button_text_color = '#000000'

        self.btn_check = self.create_button(button_frame, self.get_translation("check"), self.check_size, width=button_width, bg_color=button_bg_color, text_color=button_text_color)
        self.btn_pack = self.create_button(button_frame, self.get_translation("pack"), self.start_pack, width=button_width, bg_color=button_bg_color, text_color=button_text_color)
        self.btn_pause = self.create_button(button_frame, self.get_translation("pause"), self.pause_pack, width=button_width, bg_color=button_bg_color, text_color=button_text_color)
        self.btn_delete = self.create_button(button_frame, self.get_translation("delete"), self.delete_user_folder, width=button_width, bg_color=button_bg_color, text_color=button_text_color)

        ToolTip(self.btn_check, self.tooltip_mapping["check"]["zh"], self.tooltip_mapping["check"]["ja"], self.tooltip_font_size, self.tooltip_text_color)
        ToolTip(self.btn_pack, self.tooltip_mapping["pack"]["zh"], self.tooltip_mapping["pack"]["ja"], self.tooltip_font_size, self.tooltip_text_color)
        ToolTip(self.btn_pause, self.tooltip_mapping["pause"]["zh"], self.tooltip_mapping["pause"]["ja"], self.tooltip_font_size, self.tooltip_text_color)
        ToolTip(self.btn_delete, self.tooltip_mapping["delete"]["zh"], self.tooltip_mapping["delete"]["ja"], self.tooltip_font_size, self.tooltip_text_color)

        self.checkbox_frame = tk.Frame(self.root, background='#172B4B')  
        self.checkbox_frame.pack(pady=5, anchor='w')

        # 修改勾選框的樣式
        style = ttk.Style()
        style.configure("Custom.TCheckbutton", background='#172B4B', foreground="#F5D07D", indicatorcolor='#172B4B', focuscolor='#172B4B')

        self.btn_toggle_language = ttk.Button(button_frame, text=self.get_translation("toggle_language"), command=self.toggle_language)
        self.btn_toggle_language.pack(side=tk.LEFT, padx=5)

        self.split_var = tk.BooleanVar(value=True)  
        self.chk_split = ttk.Checkbutton(self.checkbox_frame, variable=self.split_var, text=self.get_translation("split"), style="Custom.TCheckbutton")
        self.chk_split.pack(side=tk.LEFT, padx=5)

        ToolTip(self.chk_split, self.tooltip_mapping["split"]["zh"], self.tooltip_mapping["split"]["ja"], self.tooltip_font_size, self.tooltip_text_color)

        self.checkboxes = {}
        self.create_language_checkboxes()

        # 信息框-顯示選擇資料夾路徑
        self.info_frame = tk.Frame(self.root, background='#172B4B')
        self.info_frame.pack(pady=10, anchor='w')  

        self.info_display_width = 30  # 設定顯示的寬度
        self.info_display_color = '#F5D07D'  # 設定顯示的文字顏色

        # 預設存檔路徑
        self.default_path = f"C:\\Box\\Personal_{Path.home().name}\\PIP_{Path.home().name}"  # 使用當前用戶名稱替換

        # 在信息框中顯示預設路徑，靠左對齊
        self.info_display = tk.Label(self.info_frame, text=self.default_path, bg='black', fg=self.info_display_color, font=('微軟正黑體', 14), anchor='w', width=self.info_display_width)
        self.info_display.pack(side=tk.LEFT, padx=10)

        self.btn_choose_location = ttk.Button(self.info_frame, text="選擇存檔位置", command=self.choose_save_location)  
        self.btn_choose_location.pack(side=tk.LEFT, padx=5)

        # 日誌顯示框
        self.action_display = tk.Text(self.root, height=10, width=50, bg='black', fg='#F5D07D', font=('微軟正黑體', 14), bd=0)
        self.action_display.pack(pady=5, padx=10)

        # 進度條框架
        self.progress_frame = tk.Frame(self.root, background='#172B4B')  
        self.progress_frame.pack(pady=5, anchor='w')  

        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar", troughcolor="#172B4B", background="#F5D07D")  

        self.progress_bar = ttk.Progressbar(self.progress_frame, length=260, mode='determinate', style="Custom.Horizontal.TProgressbar")  
        self.progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X)

        self.progress_label = tk.Label(self.progress_frame, text="", bg="#172B4B", fg='#F5D07D', font=('微軟正黑體', 10))  
        self.progress_label.pack(side=tk.RIGHT, padx=3)

        # 添加關於按鈕
        self.btn_about = ttk.Button(button_frame, text=self.get_translation("about"), command=self.show_about)
        self.btn_about.pack(side=tk.RIGHT, padx=5)

    def create_button(self, parent, text, command, width=10, bg_color=None, text_color=None):
        """創建按鈕並控制顏色"""
        btn = ttk.Button(parent, text=text, command=command, width=width)
        btn.pack(side=tk.LEFT, padx=5)
        
        # 設置按鈕背景顏色和文字顏色
        if bg_color or text_color:
            style = ttk.Style()
            style.configure("Custom.TButton", background=bg_color, foreground=text_color)
            btn.configure(style="Custom.TButton")
        
        return btn

    def detect_language(self):
        lang, _ = locale.getdefaultlocale()
        # 將預設語言設置為繁體中文
        self.current_language = "zh"

    def get_translation(self, key):
        """獲取翻譯對應值，預設為繁體中文。"""
        return self.language_mapping[key][self.current_language]

    def toggle_language(self):
        """切換語言的功能，現在支援繁體中文和日文。"""
        if self.current_language == "zh":
            self.current_language = "ja"  # 切換到日文
        else:
            self.current_language = "zh"  # 切換回繁體中文
        
        # 更新語言切換按鈕的文本
        self.btn_toggle_language.config(text=self.get_translation("toggle_language"))

        # 更新 UI 中的所有文本
        self.update_ui_text()

    def update_ui_text(self):
        """更新 UI 中所有文本以反映當前所選語言。"""
        self.btn_check.config(text=self.get_translation("check"))
        self.btn_pack.config(text=self.get_translation("pack"))
        self.btn_pause.config(text=self.get_translation("pause"))
        self.btn_delete.config(text=self.get_translation("delete"))
        self.chk_split.config(text=self.get_translation("split"))  # 更新 "分割" 的勾選框文本

        # 刪除舊的勾選框
        for chk in list(self.checkboxes.values()):
            chk['text'] = ""
            chk.destroy()
        self.checkboxes.clear()

        # 創建新的勾選框
        self.create_language_checkboxes()

        self.action_display.delete(1.0, tk.END)  # 清空日誌顯示
        self.action_display.insert(tk.END, self.get_translation("checking_result"))  # 插入檢查結果文本

    def create_language_checkboxes(self):
        """創建勾選框以反映當前語言選項 (現在為英文)"""
        for folder_key in self.checkbox_language_mapping.keys():
            var = tk.BooleanVar(value=True)
            self.checkboxes[folder_key] = var
            text = self.checkbox_language_mapping[folder_key]  # 直接使用繁體中文名稱
            chk = ttk.Checkbutton(self.checkbox_frame, text=text, variable=var, style="Custom.TCheckbutton")  # 使用自定義樣式
            chk.pack(side=tk.LEFT, padx=5)

    def log_action(self, message):
        self.action_display.insert(tk.END, message + "\n")
        self.action_display.see(tk.END)
        self.root.update()

    def format_size(self, size):
        """格式化大小，返回可讀格式。"""
        if size >= 1024 ** 3:
            return f"{size / (1024 ** 3):.2f} GB"
        elif size >= 1024 ** 2:
            return f"{size / (1024 ** 2):.2f} MB"
        elif size >= 1024:
            return f"{size / 1024:.2f} KB"
        else:
            return f"{size} Bytes"

    def check_size(self):
        user_folder = Path.home()
        total_size = 0

        self.log_action(self.get_translation("checking_result"))

        for folder_name, var in self.checkboxes.items():
            if var.get():
                folder_path = user_folder / folder_name
                self.log_action(f"檢查 {folder_path}")
                if folder_path.exists() and folder_path.is_dir():
                    size = self.get_folder_size(folder_path)
                    total_size += size
                    formatted_size = self.format_size(size)
                    bold_text = f"－{folder_name}：{formatted_size}"
                    self.log_action(bold_text)
                else:
                    self.log_action(f"{self.get_translation('folder_not_found')}{folder_path}")

        self.log_action(f"當前需要打包的總容量：{self.format_size(total_size)}")
        if total_size > 0:
            self.progress_bar['value'] = 100 * (total_size / (1024 ** 3))
        self.update_progress_label(0, total_size, total_size)  # 更新初始進度標籤

    def get_folder_size(self, folder):
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(folder):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                total_size += os.path.getsize(fp)
        return total_size

    def start_pack(self):
        if self.worker_thread is None or not self.worker_thread.is_alive():
            self.is_running = True
            self.is_paused = False
            self.progress_bar['value'] = 0
            self.worker_thread = threading.Thread(target=self.compress_files, args=(self.split_var.get(),))
            self.worker_thread.start()

    def compress_files(self, split=False):
        # 使用 "C:\Box" 作為目標文件夾
        target_folder = Path(f"C:\\Box\\Personal_{Path.home().name}\\PIP_{Path.home().name}")  # 這裡使用用戶名生成路徑

        # 確保目標資料夾存在
        self.log_action(f"{self.get_translation('delete_folder_msg')}{target_folder}")
        target_folder.mkdir(parents=True, exist_ok=True)

        folders_to_compress = [folder_name for folder_name, var in self.checkboxes.items() if var.get()]

        total_initial_size = sum(self.get_folder_size(Path.home() / folder_name) for folder_name in folders_to_compress)
        processed_size = 0  
        for folder_name in folders_to_compress:
            folder_to_compress = Path.home() / folder_name
            if folder_to_compress.exists() and folder_to_compress.is_dir():
                size = self.get_folder_size(folder_to_compress)
                num_parts = (size // (40 * 1024 ** 3)) + 1 if split and size > (40 * 1024 ** 3) else 1

                for i in range(num_parts):
                    if num_parts > 1:
                        zip_file_path = target_folder / f"{Path.home().name}_{folder_name}_{i + 1:02}.7z"
                    else:
                        zip_file_path = target_folder / f"{Path.home().name}_{folder_name}.7z"
                    
                    self.log_action(f"正在壓縮: {folder_to_compress.name} 到 {zip_file_path}")
                    self.progress_bar['value'] = 100 * (processed_size / total_initial_size)  
                    self.update_progress_label(processed_size, total_initial_size, size)  
                    self.root.update()

                    command = [r"C:\Program Files\7-Zip\7z.exe", 'a', str(zip_file_path), str(folder_to_compress) + '\\*']
                    process = subprocess.Popen(command)
                    process.wait()

                    processed_size += size  

        self.log_action(f"所有壓縮文件已生成。")
        self.progress_bar['value'] = 100
        self.update_progress_label(processed_size, total_initial_size, processed_size)  

        if self.is_running:
            messagebox.showinfo("完成", self.get_translation("compression_complete"))
        self.is_running = False

    def update_progress_label(self, current_size, total_size, folder_size):
        current_size_formatted = self.format_size(current_size)
        total_size_formatted = self.format_size(total_size)
        progress_percentage = int((current_size / total_size) * 100) if total_size > 0 else 0

        label_text = f"{current_size_formatted} / {total_size_formatted} － {progress_percentage}%"
        self.progress_label.config(text=label_text)

    def pause_pack(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.log_action("打包已暫停。")
        else:
            self.log_action("打包已繼續。")

    def delete_user_folder(self):
        user_folder = Path.home()
        target_folder = Path(f"C:\\Box\\Personal_{user_folder.name}\\PIP_{user_folder.name}")  # 確定將要刪除的目標資料夾

        if target_folder.exists() and target_folder.is_dir():
            self.log_action(f"{self.get_translation('delete_folder_msg')}{target_folder}")
            for item in target_folder.glob("*"):
                if item.is_file():
                    os.remove(item)
                    self.log_action(f"已刪除文件: {item.name}")
                elif item.is_dir():
                    os.rmdir(item)

            os.rmdir(target_folder)
            self.log_action(f"已刪除使用者資料夾: {target_folder.name}")
        else:
            messagebox.showinfo("刪除", self.get_translation("delete_msg"))

    def choose_save_location(self):
        from tkinter import filedialog
        folder_selected = filedialog.askdirectory(initialdir=self.default_path, title="選擇存檔位置")  # 使用預設路徑
        if folder_selected:
            self.info_display.config(text=folder_selected)

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title(self.get_translation("about"))
        about_window.geometry("400x300")
        about_window.configure(background='#172B4B')

        description_label = tk.Label(about_window, text=self.get_translation("description"), bg="#172B4B", fg="#F5D07D", font=('微軟正黑體', 12), justify='left')
        description_label.pack(pady=10)

        website_label = tk.Label(about_window, text=self.get_translation("website"), fg="#F5D07D", cursor="hand2", bg="#172B4B")  # 更改為紅色
        website_label.pack(pady=5)
        website_label.bind("<Button-1>", lambda e: webbrowser.open_new("https://tetservicedesk.asia.tel.com:8080/"))

        contact_label = tk.Label(about_window, text=self.get_translation("contact"), fg="#F5D07D", cursor="hand2", bg="#172B4B")  # 更改為紅色
        contact_label.pack(pady=5)
        contact_label.bind("<Button-1>", lambda e: webbrowser.open_new("mailto:TETISTeamMis@tel.com"))

        author_label = tk.Label(about_window, text=self.get_translation("author"), bg="#172B4B", fg="#F5D07D", font=('微軟正黑體', 12))
        author_label.pack(pady=10)

        close_button = ttk.Button(about_window, text="關閉", command=about_window.destroy)
        close_button.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()  # 創建主窗口
    app = FileCompressorApp(root)  # 創建應用程序實例
    root.mainloop()  # 啟動事件循環