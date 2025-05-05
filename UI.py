import tkinter as tk
from tkinter import ttk
from tkinter import messagebox


class GenericApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generic Application")
        self.root.geometry("500x400")
        self.root.configure(background='#172B4B')  # 背景顏色

        # 字體設置
        self.font_size = 14
        self.font = ('微軟正黑體', self.font_size)

        self.root.option_add("*Font", self.font)
        self.root.option_add("*Foreground", "#F5D07D")  # 全局前景色設置

        self.create_widgets()

    def create_widgets(self):
        # 創建按鈕框架
        button_frame = tk.Frame(self.root, background='#172B4B')
        button_frame.pack(pady=5, anchor='w')

        # 按鈕設置
        self.btn_example = self.create_button(button_frame, "範例按鈕", self.example_command)
        self.btn_about = self.create_button(button_frame, "關於", self.show_about)

        # 日志顯示框
        self.action_display = tk.Text(self.root, height=10, width=50, bg='black', fg='#F5D07D', font=self.font, bd=0)
        self.action_display.pack(pady=5, padx=10)

    def create_button(self, parent, text, command):
        """創建按鈕並設置顏色"""
        btn = ttk.Button(parent, text=text, command=command)
        btn.pack(side=tk.LEFT, padx=5)
        btn.configure(style="Custom.TButton")

        style = ttk.Style()
        style.configure("Custom.TButton", background='#172B4B', foreground='#000000')

        return btn

    def example_command(self):
        self.log_action("範例按鈕被點擊！")

    def log_action(self, message):
        self.action_display.insert(tk.END, message + "\n")
        self.action_display.see(tk.END)

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("關於")
        about_window.geometry("300x200")
        about_window.configure(background='#172B4B')

        author_label = tk.Label(about_window, text="Creat by: Your Name", bg="#172B4B", fg="#F5D07D", font=self.font)
        author_label.pack(pady=10)

        close_button = ttk.Button(about_window, text="關閉", command=about_window.destroy)
        close_button.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = GenericApp(root)
    root.mainloop()