import os
import shutil
from pathlib import Path
import subprocess  # 用於執行 7z 命令
import zipfile  # 修正 ZipFile 的導入
import ttkbootstrap as ttk  # 使用 ttkbootstrap 替代 tkinter
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext
import threading
import locale
import webbrowser  # 引入 webbrowser 以便在新的窗口中打開超連結

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
        self.tip_window = ttk.Toplevel(self.widget)  # 建立提示框
        self.tip_window.wm_overrideredirect(True)  # 移除標題列
        self.tip_window.wm_geometry(f"+{x}+{y}")  # 設定提示框位置
        label = ttk.Label(self.tip_window, text=self.get_current_tooltip_text(), bootstyle="light", font=('微軟正黑體', self.font_size))
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

        # 設定程式標題
        self.root.title("QuickPack1.0")  # 修改程式顯示名稱

        # 預設存檔路徑
        self.default_path = Path("Z:\\")  # 指定到 Z 槽

        # 初始化語言
        self.current_language = "zh"  # 預設語言
        self.detect_language()  # 自動檢測系統語言

        # 初始化變數
        self.is_running = False  # 壓縮執行狀態
        self.worker_thread = None  # 壓縮執行緒
        self.special_checkboxes = {}  # 初始化特殊勾選框變數

        # 初始化普通勾選框的語言對映
        self.checkbox_language_mapping = {
            "Desktop": "Desktop",
            "Documents": "Documents",
            "Downloads": "Downloads",
            "Pictures": "Pictures",
            "Favorites": "Favorites"  
        }

        # 初始化 special_folders
        self.special_folders = {
            "quick_access": {
                "path": os.getenv("APPDATA") + r"\Microsoft\Windows\Recent\AutomaticDestinations",
                "files": None
            },
            "network_shortcuts": {
                "path": os.getenv("APPDATA") + r"\Microsoft\Windows\Network Shortcuts",
                "files": None
            },
            "bookmarks": {
                "path": os.getenv("LOCALAPPDATA") + r"\Google\Chrome\User Data\Default",
                "files": ["Bookmarks", "Bookmarks.bak"]
            },
            "signatures": {"path": os.getenv("APPDATA") + r"\Microsoft\Signatures", "files": None}
        }

        # 初始化 language_mapping
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
            "contact": {"zh": "透過信箱聯絡我們(請點擊)", "ja": "メールでお問い合わせください（クリックしてください）"},
            "bookmarks": {"zh": "書籤", "ja": "ブックマーク"},
            "signatures": {"zh": "簽名檔", "ja": "署名ファイル"},
            "quick_access": {"zh": "快速存取", "ja": "クイックアクセス"},
            "network_shortcuts": {"zh": "網路磁碟", "ja": "ネットワークショートカット"}
        }

        # 初始化 UI
        self.create_widgets()

    def create_widgets(self):
        """建立 UI 元件"""
        # 定義按鈕樣式
        style = ttk.Style()
        style.configure("Custom.TButton", font=('微軟正黑體', 11, 'bold'))  # 設定字體大小和粗體

        # 按鈕區域
        button_frame = ttk.Frame(self.root)  # 按鈕框架
        button_frame.pack(pady=5, anchor='w')  # 設定框架位置

        button_width = 10  # 按鈕寬度

        # 檢查按鈕
        self.btn_check = ttk.Button(button_frame, text=self.get_translation("check"), command=self.check_size, width=button_width, bootstyle="success", style="Custom.TButton")  
        self.btn_check.pack(side=LEFT, padx=5)

        # 複製按鈕（將位置調整到壓縮按鈕之前）
        self.btn_copy = ttk.Button(button_frame, text=self.get_translation("copy"), command=self.copy_files, width=button_width, bootstyle="warning", style="Custom.TButton")  
        self.btn_copy.pack(side=LEFT, padx=5)

        # 壓縮按鈕（大小與檢查按鈕相同，位置調整到複製按鈕之後）
        self.btn_compress = ttk.Button(button_frame, text=self.get_translation("compress"), command=self.start_compress, width=button_width, bootstyle="primary", style="Custom.TButton")  
        self.btn_compress.pack(side=LEFT, padx=5)

        # 語言切換按鈕
        self.btn_toggle_language = ttk.Button(button_frame, text=self.get_translation("toggle_language"), command=self.toggle_language, bootstyle="danger", style="Custom.TButton")  # 紅色按鈕
        self.btn_toggle_language.pack(side=LEFT, padx=5)

        # 勾選框區域
        self.checkbox_frame = ttk.Frame(self.root)  # 勾選框框架
        self.checkbox_frame.pack(pady=5, anchor='w')

        # 分割選項
        self.split_var = ttk.BooleanVar(value=True)  # 分割選項變數
        self.chk_split = ttk.Checkbutton(self.checkbox_frame, variable=self.split_var, text=self.get_translation("split"), bootstyle="success")  # 綠色選中
        self.chk_split.pack(side=LEFT, padx=5)

        # 初始化勾選框
        self.checkboxes = {}
        self.create_language_checkboxes()

        # 創建特殊勾選框
        self.create_special_checkboxes()

        # 信息框-顯示選擇資料夾路徑
        self.info_frame = ttk.Frame(self.root)  # 信息框框架
        self.info_frame.pack(pady=10, anchor='w')

        self.info_display_width = 40  # 信息框寬度

         # 預設存檔路徑
        self.default_path = Path("Z:\\")  # 使用新的預設路徑

        # 顯示預設路徑
        self.info_display = ttk.Label(self.info_frame, text=self.default_path, bootstyle="light", font=('微軟正黑體', 14), anchor='w', width=self.info_display_width)
        self.info_display.pack(side=LEFT, padx=10)

        # 選擇存檔位置按鈕
        self.btn_choose_location = ttk.Button(self.info_frame, text="選擇存檔位置", command=self.choose_save_location, bootstyle="success")  # 綠色按鈕
        self.btn_choose_location.pack(side=LEFT, padx=5)

        # 日誌顯示框
        self.action_display = scrolledtext.ScrolledText(
            self.root, 
            height=15, 
            width=60,  # 原寬度 90
            bg="#263238", 
            fg="white", 
            font=('微軟正黑體', 12)  # 原文字大小 10，放大 2 個單位
        )
        self.action_display.pack(pady=5, padx=10)

        # 進度條框架
        self.progress_frame = ttk.Frame(self.root)  # 進度條框架
        self.progress_frame.pack(pady=5, anchor='w')

        # 進度條
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=700, mode='determinate', style="Custom.Horizontal.TProgressbar")  # 調整進度條寬度
        self.progress_bar.pack(side=LEFT, padx=10, fill=X)

        # 自定義進度條樣式，增加高度
        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar", thickness=15)  # 增加高度一個單位

        # 進度條標籤
        self.progress_label = ttk.Label(self.progress_frame, text="", bootstyle="light", font=('微軟正黑體', 10))
        self.progress_label.pack(side=RIGHT, padx=3)

        # 關於按鈕
        self.btn_about = ttk.Button(button_frame, text=self.get_translation("about"), command=self.show_about, bootstyle="danger")  # 紅色按鈕
        self.btn_about.pack(side=RIGHT, padx=5)


    def create_language_checkboxes(self):
        """創建語言選項的勾選框"""
        # 定義樣式
        style = ttk.Style()
        style.configure("Custom.TCheckbutton", font=('微軟正黑體', 12))  # 自定義字體大小

        for folder_key, folder_name in self.checkbox_language_mapping.items():
            var = ttk.BooleanVar(value=True)  # 綁定布林變數
            self.checkboxes[folder_key] = var
            # 使用自定義樣式 "Custom.TCheckbutton"
            chk = ttk.Checkbutton(self.checkbox_frame, text=folder_name, variable=var, bootstyle="success", style="Custom.TCheckbutton")
            chk.pack(side=LEFT, padx=5)

    def create_special_checkboxes(self):
        """創建特殊資料夾的勾選框"""
        special_frame = ttk.Frame(self.root)  # 特殊勾選框框架
        special_frame.pack(pady=5, anchor='w')  # 放置在主視窗中，靠左對齊

        # 在 Z 槽下建立一個名為「當前使用者」的資料夾
        user_folder = Path(self.default_path) / Path.home().name
        if user_folder.exists() and not user_folder.is_dir():
            user_folder.unlink()  # 如果是檔案則刪除
        user_folder.mkdir(parents=True, exist_ok=True)  # 確保資料夾存在

        for key, info in self.special_folders.items():
            var = ttk.BooleanVar(value=True)  # 預設為勾選
            self.special_checkboxes[key] = var
            chk = ttk.Checkbutton(special_frame, text=self.get_translation(key), variable=var, bootstyle="info")
            chk.pack(side=LEFT, padx=5)  # 水平排列，間距為 5

            # 更新特殊資料夾的目標路徑到「當前使用者」資料夾內
            info["destination"] = user_folder / key

    def choose_save_location(self):
        """選擇存檔位置"""
        selected_path = filedialog.askdirectory(title="選擇存檔位置")  # 彈出選擇資料夾視窗
        if selected_path:
            self.default_path = selected_path
            self.info_display.config(text=self.default_path)  # 更新顯示的存檔路徑
            self.log_action(f"存檔位置已更新為: {self.default_path}")

    def log_action(self, message):
        """記錄操作日誌"""
        self.action_display.insert("end", f"{message}\n")  # 插入日誌訊息
        self.action_display.see("end")  # 自動滾動到最新訊息

    def check_size(self):
        """檢查所選資料夾的大小和 C 槽剩餘空間"""
        # 檢查 C 槽剩餘空間
        c_drive = Path("C:\\")
        free_space = shutil.disk_usage(c_drive).free  # 獲取剩餘空間（以位元組為單位）
        formatted_free_space = self.format_size(free_space)

        # 判斷剩餘空間並顯示提示
        if free_space < 10 * 1024 * 1024 * 1024:  # 小於 10 GB
            self.action_display.insert("end", f"\n硬碟空間剩餘：{formatted_free_space} 請停止動作，請洽IS人員協助\n", "warning")
            self.action_display.tag_config("warning", foreground="red", font=('微軟正黑體', 12, 'bold'))
            
            # 禁用壓縮按鈕
            self.btn_compress.config(state=DISABLED, bootstyle="secondary")  # 灰色按鈕
            return
        else:
            self.action_display.insert("end", f"\n硬碟空間剩餘：{formatted_free_space}\n", "normal")
            self.action_display.tag_config("normal", font=('微軟正黑體', 12, 'bold'))
            
            # 啟用壓縮按鈕
            self.btn_compress.config(state=NORMAL, bootstyle="primary")  # 恢復藍色按鈕

        # 檢查所選資料夾的大小
        total_size = 0
        for folder_name, var in self.checkboxes.items():
            if var.get():  # 如果勾選框被選中
                folder_path = Path("C:/Users") / os.getlogin() / folder_name  # 明確指定路徑
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
        generate_bat = False  # 標記是否需要生成批次檔

        # 統計要壓縮的資料夾數量
        compress_targets = [folder_name for folder_name, var in self.checkboxes.items() if var.get()]
        total_targets = len(compress_targets)
        if total_targets == 0:
            self.log_action("未選擇任何資料夾進行壓縮。")
            return

        self.progress_bar["maximum"] = 10
        self.progress_bar["value"] = 0

        for idx, folder_name in enumerate(compress_targets, 1):
            generate_bat = True
            folder_path = Path.home() / folder_name

            if not folder_path.exists():
                self.log_action(f"來源資料夾不存在: {folder_path}")
                continue

            zip_file = Path(self.default_path) / f"{folder_name}.zip"
            self.log_action(f"正在壓縮: {folder_path} -> {zip_file}")
            try:
                with zipfile.ZipFile(zip_file, 'w') as zipf:
                    for root, _, files in os.walk(folder_path):
                        for file in files:
                            file_path = Path(root) / file
                            zipf.write(file_path, file_path.relative_to(folder_path))
                self.log_action(f"壓縮完成: {zip_file}")
            except Exception as e:
                self.log_action(f"壓縮失敗: {folder_path} -> {zip_file}，錯誤: {e}")

            # 進度條推進
            percent = int(idx / total_targets * 10)
            self.progress_bar["value"] = percent
            self.progress_label.config(text=f"進度: {percent * 10}%")
            self.root.update_idletasks()

        if generate_bat:
            self.generate_restore_bat()  # 生成批次檔

        self.progress_bar["value"] = 10
        self.progress_label.config(text="進度: 100%")
        self.log_action("所有壓縮文件已生成。")
        self.is_running = False

        # 自動開啟目標資料夾
        os.startfile(self.default_path)

    def copy_files(self):
        """複製勾選的資料夾內容到使用者指定的目標資料夾"""
        import time

        generate_bat = False
        total_files = 0
        copied_files = 0
        total_size = 0

        # 計算普通資料夾檔案數量與容量
        for folder_name, var in self.checkboxes.items():
            if var.get():
                folder_path = Path.home() / folder_name
                if folder_path.exists():
                    files = [f for f in folder_path.glob('**/*') if f.is_file()]
                    folder_size = sum(f.stat().st_size for f in files)
                    total_files += len(files)
                    total_size += folder_size

        # 計算特殊資料夾檔案數量與容量，並顯示於日誌
        for key, var in self.special_checkboxes.items():
            if var.get():
                generate_bat = True
                folder_info = self.special_folders[key]
                folder_path = Path(folder_info["path"])
                if folder_path.exists():
                    files = [f for f in folder_path.glob('**/*') if f.is_file()]
                    folder_size = sum(f.stat().st_size for f in files)
                    self.log_action(f"{self.get_translation(key)}：檔案數量 {len(files)}，容量 {self.format_size(folder_size)}")
                    total_files += len(files)
                    total_size += folder_size
                else:
                    self.log_action(f"{self.get_translation('folder_not_found')}{folder_path}")

        self.log_action(f"特殊資料夾總容量：{self.format_size(total_size)}，總檔案數量：{total_files}")
        self.log_action("5秒後開始複製...")
        self.root.update_idletasks()
        time.sleep(5)

        # 暫停進度條功能
        # self.progress_bar["maximum"] = 10
        # self.progress_bar["value"] = 0
        # progress_unit = max(1, total_files // 10)
        # progress_count = 0

        # 複製普通資料夾
        for folder_name, var in self.checkboxes.items():
            if var.get():
                folder_path = Path.home() / folder_name
                destination = Path(self.default_path) / os.getlogin() / folder_name
                destination.mkdir(parents=True, exist_ok=True)
                if folder_path.exists():
                    for file in folder_path.glob('**/*'):
                        if file.is_file():
                            try:
                                shutil.copy(file, destination / file.name)
                                copied_files += 1
                            except PermissionError:
                                self.log_action(f"無法訪問檔案（權限被拒絕）: {file}")
                            except Exception as e:
                                self.log_action(f"複製檔案失敗: {file}，錯誤: {e}")

        # 複製特殊資料夾
        for key, var in self.special_checkboxes.items():
            if var.get():
                folder_info = self.special_folders[key]
                folder_path = Path(folder_info["path"])
                destination = Path(self.default_path) / os.getlogin() / key
                destination.mkdir(parents=True, exist_ok=True)
                if folder_path.exists():
                    for file in folder_path.glob('**/*'):
                        if file.is_file():
                            try:
                                shutil.copy(file, destination / file.name)
                                copied_files += 1
                            except PermissionError:
                                self.log_action(f"無法訪問檔案（權限被拒絕）: {file}")
                            except FileNotFoundError:
                                self.log_action(f"檔案不存在: {file}")
                            except Exception as e:
                                self.log_action(f"複製檔案失敗: {file}，錯誤: {e}")
                else:
                    self.log_action(f"來源資料夾不存在: {folder_path}")

        # self.progress_bar["value"] = 10
        # self.progress_label.config(text="進度: 100%")
        self.log_action("所有選定的資料夾已成功複製。")

        if generate_bat:
            self.generate_restore_bat()

    def generate_restore_bat(self):
        """生成自動復原.bat 文件"""
        # 獲取使用者帳號名稱
        user_folder_name = os.getlogin()
        user_folder_path = Path(self.default_path) / user_folder_name
        user_folder_path.mkdir(parents=True, exist_ok=True)  # 確保使用者資料夾存在

        # 將批次檔放入使用者資料夾內
        bat_file_path = user_folder_path / "自動復原.bat"
        try:
            # 使用 UTF-8 with BOM 編碼寫入批次檔
            with open(bat_file_path, "w", encoding="utf-8-sig") as bat_file:
                bat_file.write(
                    """@echo off

REM 1. 複製 Bookmarks 和 Bookmarks.bak
xcopy /Y /Q "%~dp0Bookmarks\\Bookmarks" "%LOCALAPPDATA%\\Google\\Chrome\\User Data\\Default\\"
xcopy /Y /Q "%~dp0Bookmarks\\Bookmarks.bak" "%LOCALAPPDATA%\\Google\\Chrome\\User Data\\Default\\"

REM 2. 複製 Signatures 資料夾內所有檔案
xcopy /Y /Q "%~dp0Signatures\\*" "%AppData%\\Microsoft\\Signatures\\"

REM 3. 複製 quick_access 資料夾內所有檔案
xcopy /Y /Q "%~dp0quick_access\\*" "%AppData%\\Microsoft\\Windows\\Recent\\AutomaticDestinations\\"

REM 4. 複製 Network Shortcuts 資料夾內所有檔案
xcopy /Y /Q "%~dp0network_shortcuts\\*" "%AppData%\\Microsoft\\Windows\\Network Shortcuts\\"

REM 等待5秒後自動關閉
timeout /t 5 >nul
exit"""
                )
            self.log_action(f"已生成批次檔: {bat_file_path}")
        except Exception as e:
            self.log_action(f"生成批次檔失敗: {e}")

    def show_about(self):
        """顯示關於視窗"""
        about_window = ttk.Toplevel(self.root)  # 建立新視窗
        about_window.title(self.get_translation("about"))  # 設定視窗標題
        about_window.geometry("300x200")  # 設定視窗大小
        about_window.resizable(False, False)  # 禁止調整視窗大小

        # 顯示作者資訊
        ttk.Label(about_window, text=self.get_translation("author"), bootstyle="light", font=('微軟正黑體', 12)).pack(pady=10)
        ttk.Label(about_window, text=self.get_translation("description"), bootstyle="light", font=('微軟正黑體', 10)).pack(pady=5)
        ttk.Label(about_window, text=self.get_translation("website"), bootstyle="light", font=('微軟正黑體', 10), cursor="hand2").pack(pady=5)
        ttk.Label(about_window, text=self.get_translation("contact"), bootstyle="light", font=('微軟正黑體', 10), cursor="hand2").pack(pady=5)

        # 關閉按鈕
        ttk.Button(about_window, text="關閉", command=about_window.destroy, bootstyle="danger").pack(pady=10)

    def detect_language(self):
        """檢測系統語言"""
        lang, _ = locale.getdefaultlocale()
        self.current_language = "zh" if lang and lang.startswith("zh") else "ja"

    def get_translation(self, key):
        """獲取翻譯對應值"""
        return self.language_mapping.get(key, {}).get(self.current_language, key)

    def toggle_language(self):
        """切換語言"""
        self.current_language = "ja" if self.current_language == "zh" else "zh"
        self.update_ui_text()

    def update_ui_text(self):
        """更新 UI 中的文本"""
        # 更新按鈕文本
        self.btn_check.config(text=self.get_translation("check"))
        self.btn_compress.config(text=self.get_translation("compress"))
        self.btn_copy.config(text=self.get_translation("copy"))
        self.btn_toggle_language.config(text=self.get_translation("toggle_language"))
        self.btn_about.config(text=self.get_translation("about"))

        # 更新分割勾選框文本
        self.chk_split.config(text=self.get_translation("split"))

        # 更新特殊勾選框文本
        for key, var in self.special_checkboxes.items():
            checkbox_text = self.get_translation(key)
            self.special_checkboxes[key].widget.config(text=checkbox_text)

        # 更新普通勾選框文本
        for folder_key, var in self.checkboxes.items():
            checkbox_text = self.checkbox_language_mapping.get(folder_key, folder_key)
            self.checkboxes[folder_key].widget.config(text=checkbox_text)

# 主程式入口
if __name__ == "__main__":
    root = ttk.Window(themename="superhero")  # 使用 ttkbootstrap 的主題
    app = FileCompressorApp(root)
    root.mainloop()
