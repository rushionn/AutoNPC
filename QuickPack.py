import os
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import locale
import webbrowser  # 引入webbrowser以便在新的窗口中打開超連結

class ToolTip:
    """懸浮提示框類"""
    def __init__(self, widget, text, text_ja=None, font_size=10, text_color="black"):
        self.widget = widget  # 綁定的元件
        self.text = text  # 提示框顯示的文字
        self.text_ja = text_ja  # 日文提示文字（多語言支持）
        self.font_size = font_size  # 提示框文字大小
        self.text_color = text_color  # 提示框文字顏色
        self.tip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.show_tip)  # 滑鼠進入時顯示提示框
        self.widget.bind("<Leave>", self.hide_tip)  # 滑鼠離開時隱藏提示框

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20  # 提示框的 X 座標
        y = self.widget.winfo_rooty() + 20  # 提示框的 Y 座標
        self.tip_window = tk.Toplevel(self.widget)  # 建立提示框
        self.tip_window.wm_overrideredirect(True)  # 移除標題列
        self.tip_window.wm_geometry(f"+{x}+{y}")  # 設定提示框位置
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
        self.tooltip_font_size = 10  # 提示框字體大小
        self.tooltip_text_color = "black"  # 提示框文字顏色

        # 多語言支持映射
        self.language_mapping = {
            "app_title": {"zh": "QuickPack v1.0", "ja": "QuickPack v1.0"},
            "check": {"zh": "檢查", "ja": "チェック"},
            "compress": {"zh": "壓縮", "ja": "圧縮"},
            "copy": {"zh": "複製", "ja": "コピー"},
            "split": {"zh": "分割", "ja": "分割"},
            "checking_result": {"zh": "檢測結果：", "ja": "チェック結果："},
            "compression_complete": {"zh": "所選資料夾已經壓縮完成。", "ja": "選択したフォルダが圧縮されました。"},
            "folder_not_found": {"zh": "檔案不存在: ", "ja": "ファイルが存在しません: "},
            "toggle_language": {"zh": "日本語へ切換", "ja": "切換為中文"},
            "about": {"zh": "詳細", "ja": "詳細"},
            "author": {"zh": "Created by Lucien", "ja": "Created by Lucien"},
            "description": {"zh": "v.1.0", "ja": "v.1.0"},
            "website": {"zh": "IS線上申請 (請點擊)", "ja": "ISオンライン申請（クリックしてください）"},
            "contact": {"zh": "透過信箱聯絡我們(請點擊)", "ja": "メールでお問い合わせください（クリックしてください）"}
        }

        # 勾選框對應的資料夾名稱
        self.checkbox_language_mapping = {
            "Desktop": "Desktop",
            "Downloads": "Downloads",
            "Documents": "Documents",
            "Pictures": "Pictures",
            "Favorites": "Favorites",
        }

        self.current_language = "zh"  # 預設語言
        self.detect_language()  # 自動檢測系統語言

        self.root.title(self.get_translation("app_title"))  # 設定視窗標題
        self.root.geometry("800x500")  # 設定視窗大小
        self.root.configure(background='#172B4B')  # 設定背景顏色

        # 字體設置
        font_size = 14  # 預設字體大小
        font_label_size = 10  # 標籤字體大小
        font_large = ('微軟正黑體', font_size)  # 預設字體
        font_label = ('微軟正黑體', font_label_size)  # 標籤字體

        self.root.option_add("*TButton*Font", font_large)  # 設定按鈕字體
        self.root.option_add("*Label*Font", font_large)  # 設定標籤字體

        # 初始化變數
        self.is_running = False  # 壓縮執行狀態
        self.worker_thread = None  # 壓縮執行緒

        # 初始化 UI
        self.create_widgets()

    def create_widgets(self):
        """建立 UI 元件"""
        # 按鈕區域
        button_frame = tk.Frame(self.root, bg="#172B4B")  # 按鈕框架
        button_frame.pack(pady=5, anchor='w')  # 設定框架位置

        button_width = 10  # 按鈕寬度

        # 檢查按鈕
        self.btn_check = tk.Button(button_frame, text=self.get_translation("check"), command=self.check_size, width=button_width, bg="#4CAF50", fg="white")  # 綠色按鈕
        self.btn_check.pack(side=tk.LEFT, padx=5)

        # 壓縮按鈕
        self.btn_compress = tk.Button(button_frame, text=self.get_translation("compress"), command=self.start_compress, width=button_width, bg="#2196F3", fg="white")  # 藍色按鈕
        self.btn_compress.pack(side=tk.LEFT, padx=5)

        # 複製按鈕
        self.btn_copy = tk.Button(button_frame, text=self.get_translation("copy"), command=self.copy_files, width=button_width, bg="#FFC107", fg="black")  # 黃色按鈕
        self.btn_copy.pack(side=tk.LEFT, padx=5)

        # 語言切換按鈕
        self.btn_toggle_language = tk.Button(button_frame, text=self.get_translation("toggle_language"), command=self.toggle_language, bg="#FF5722", fg="white")  # 橙色按鈕
        self.btn_toggle_language.pack(side=tk.LEFT, padx=5)

        # 勾選框區域
        self.checkbox_frame = tk.Frame(self.root, bg="#172B4B")  # 勾選框框架
        self.checkbox_frame.pack(pady=5, anchor='w')

        # 分割選項
        self.split_var = tk.BooleanVar(value=True)  # 分割選項變數
        self.chk_split = tk.Checkbutton(self.checkbox_frame, variable=self.split_var, text=self.get_translation("split"), bg="#172B4B", fg="white", selectcolor="#4CAF50")  # 綠色選中
        self.chk_split.pack(side=tk.LEFT, padx=5)

        # 初始化勾選框
        self.checkboxes = {}
        self.create_language_checkboxes()

        # 信息框-顯示選擇資料夾路徑
        self.info_frame = tk.Frame(self.root, bg="#172B4B", highlightbackground="white", highlightthickness=1)  # 增加邊框
        self.info_frame.pack(pady=10, anchor='w')

        self.info_display_width = 50  # 信息框寬度

        # 預設存檔路徑
        self.default_path = f"Z:\\{Path.home().name}"  # 使用新的預設路徑

        # 顯示預設路徑
        self.info_display = tk.Label(self.info_frame, text=self.default_path, bg="#172B4B", fg="white", font=('微軟正黑體', 14), anchor='w', width=self.info_display_width)
        self.info_display.pack(side=tk.LEFT, padx=10)

        # 選擇存檔位置按鈕
        self.btn_choose_location = tk.Button(self.info_frame, text="選擇存檔位置", command=self.choose_save_location, bg="#4CAF50", fg="white")  # 綠色按鈕
        self.btn_choose_location.pack(side=tk.LEFT, padx=5)

        # 日誌顯示框
        self.action_display = tk.Text(self.root, height=15, width=90, bg="#263238", fg="white")  # 調整寬度接近填滿
        self.action_display.pack(pady=5, padx=10)

        # 進度條框架
        self.progress_frame = tk.Frame(self.root, bg="#172B4B")  # 進度條框架
        self.progress_frame.pack(pady=5, anchor='w')

        # 進度條
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=700, mode='determinate')  # 調整進度條寬度
        self.progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X)

        # 進度條標籤
        self.progress_label = tk.Label(self.progress_frame, text="", bg="#172B4B", fg="white", font=('微軟正黑體', 10))
        self.progress_label.pack(side=tk.RIGHT, padx=3)

        # 關於按鈕
        self.btn_about = tk.Button(button_frame, text=self.get_translation("about"), command=self.show_about, bg="#FF5722", fg="white")  # 橙色按鈕
        self.btn_about.pack(side=tk.RIGHT, padx=5)

# ==上半部==
    def create_language_checkboxes(self):
        """創建語言選項的勾選框"""
        for folder_key, folder_name in self.checkbox_language_mapping.items():
            var = tk.BooleanVar(value=True)  # 綁定布林變數
            self.checkboxes[folder_key] = var
            # 勾選框的字體大小放大 2 個單位
            chk = tk.Checkbutton(self.checkbox_frame, text=folder_name, variable=var, bg="#172B4B", fg="white", selectcolor="#4CAF50", font=('微軟正黑體', 12))
            chk.pack(side=tk.LEFT, padx=5)

    def choose_save_location(self):
        """選擇存檔位置"""
        selected_path = filedialog.askdirectory(title="選擇存檔位置")  # 彈出選擇資料夾視窗
        if selected_path:
            self.default_path = selected_path
            self.info_display.config(text=self.default_path)  # 更新顯示的存檔路徑
            self.log_action(f"存檔位置已更新為: {self.default_path}")

    def log_action(self, message):
        """記錄操作日誌"""
        self.action_display.insert(tk.END, f"{message}\n")  # 插入日誌訊息
        self.action_display.see(tk.END)  # 自動滾動到最新訊息

    def check_size(self):
        """檢查所選資料夾的大小"""
        total_size = 0
        for folder_name, var in self.checkboxes.items():
            if var.get():  # 如果勾選框被選中
                folder_path = Path.home() / folder_name
                if folder_path.exists() and folder_path.is_dir():
                    folder_size = sum(f.stat().st_size for f in folder_path.glob('**/*') if f.is_file())
                    total_size += folder_size
                    self.log_action(f"{folder_name}: {self.format_size(folder_size)}")
                else:
                    self.log_action(f"{self.get_translation('folder_not_found')}{folder_path}")
        self.log_action(f"總大小: {self.format_size(total_size)}")

    def format_size(self, size):
        """格式化檔案大小"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024:
                return f"{size:.2f} {unit}"
            size /= 1024

    def start_compress(self):
        """開始壓縮所選資料夾"""
        if self.is_running:
            messagebox.showwarning("警告", "壓縮正在進行中，請稍候完成後再操作。")
            return

        self.is_running = True
        self.worker_thread = threading.Thread(target=self.compress_folders)  # 建立壓縮執行緒
        self.worker_thread.start()

    def compress_folders(self):
        """壓縮所選資料夾"""
        total_size = 0
        processed_size = 0

        # 計算總大小
        for folder_name, var in self.checkboxes.items():
            if var.get():
                folder_path = Path.home() / folder_name
                if folder_path.exists() and folder_path.is_dir():
                    total_size += sum(f.stat().st_size for f in folder_path.glob('**/*') if f.is_file())

        # 壓縮過程
        for folder_name, var in self.checkboxes.items():
            if var.get():
                folder_path = Path.home() / folder_name
                if folder_path.exists() and folder_path.is_dir():
                    zip_file = Path(self.default_path) / f"{folder_name}.zip"
                    self.log_action(f"正在壓縮: {folder_path} -> {zip_file}")
                    with shutil.ZipFile(zip_file, 'w') as zipf:
                        for file in folder_path.glob('**/*'):
                            if file.is_file():
                                zipf.write(file, file.relative_to(folder_path))
                                processed_size += file.stat().st_size
                                self.update_progress_label(processed_size, total_size)
                    self.log_action(f"壓縮完成: {zip_file}")
                else:
                    self.log_action(f"{self.get_translation('folder_not_found')}{folder_path}")

        self.log_action(f"所有壓縮文件已生成。")
        self.progress_bar['value'] = 100
        self.update_progress_label(total_size, total_size)

        if self.is_running:
            messagebox.showinfo("完成", self.get_translation("compression_complete"))
        self.is_running = False

    def update_progress_label(self, current_size, total_size):
        """更新進度條標籤"""
        progress_percentage = int((current_size / total_size) * 100) if total_size > 0 else 0
        self.progress_bar['value'] = progress_percentage
        self.progress_label.config(text=f"{progress_percentage}%")

    def copy_files(self):
        """複製勾選的資料夾內容到使用者指定的目標資料夾"""
        target_folder = filedialog.askdirectory(title="選擇目標資料夾")  # 彈出選擇資料夾視窗
        if not target_folder:
            self.log_action("未選擇目標資料夾，複製操作已取消。")
            return

        target_folder = Path(target_folder)
        self.log_action(f"目標資料夾: {target_folder}")

        for folder_name, var in self.checkboxes.items():
            if var.get():
                source_folder = Path.home() / folder_name
                if source_folder.exists() and source_folder.is_dir():
                    destination = target_folder / folder_name
                    self.log_action(f"正在複製 {source_folder} 到 {destination}")
                    try:
                        shutil.copytree(source_folder, destination, dirs_exist_ok=True)
                        self.log_action(f"成功複製 {source_folder} 到 {destination}")
                    except Exception as e:
                        self.log_action(f"複製失敗: {source_folder} -> {destination}，錯誤: {e}")
                else:
                    self.log_action(f"{self.get_translation('folder_not_found')}{source_folder}")

        messagebox.showinfo("完成", "所有選定的資料夾已成功複製。")

    def show_about(self):
        """顯示關於視窗"""
        about_window = tk.Toplevel(self.root)  # 建立新視窗
        about_window.title(self.get_translation("about"))  # 設定視窗標題
        about_window.geometry("300x200")  # 設定視窗大小
        about_window.resizable(False, False)  # 禁止調整視窗大小

        # 顯示作者資訊
        tk.Label(about_window, text=self.get_translation("author"), bg="#172B4B", fg="white", font=('微軟正黑體', 12)).pack(pady=10)
        tk.Label(about_window, text=self.get_translation("description"), bg="#172B4B", fg="white", font=('微軟正黑體', 10)).pack(pady=5)
        tk.Label(about_window, text=self.get_translation("website"), bg="#172B4B", fg="white", font=('微軟正黑體', 10), cursor="hand2").pack(pady=5)
        tk.Label(about_window, text=self.get_translation("contact"), bg="#172B4B", fg="white", font=('微軟正黑體', 10), cursor="hand2").pack(pady=5)

        # 關閉按鈕
        tk.Button(about_window, text="關閉", command=about_window.destroy, bg="#FF5722", fg="white").pack(pady=10)

    def detect_language(self):
        """檢測系統語言"""
        lang, _ = locale.getdefaultlocale()
        self.current_language = "zh" if lang.startswith("zh") else "ja"

    def get_translation(self, key):
        """獲取翻譯對應值"""
        return self.language_mapping[key][self.current_language]

    def toggle_language(self):
        """切換語言"""
        self.current_language = "ja" if self.current_language == "zh" else "zh"
        self.update_ui_text()

    def update_ui_text(self):
        """更新 UI 中的文本"""
        self.btn_check.config(text=self.get_translation("check"))
        self.btn_compress.config(text=self.get_translation("compress"))
        self.btn_copy.config(text=self.get_translation("copy"))
        self.chk_split.config(text=self.get_translation("split"))
        self.btn_toggle_language.config(text=self.get_translation("toggle_language"))
        self.btn_about.config(text=self.get_translation("about"))

# 主程式入口
if __name__ == "__main__":
    root = tk.Tk()
    app = FileCompressorApp(root)
    root.mainloop()