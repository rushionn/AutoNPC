import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, END
from tkinter import ttk
import threading, time, json, os, datetime
import keyboard, mouse
import subprocess
import pynput.mouse
import pyautogui
import ctypes
from screeninfo import get_monitors

# Windows API 結構定義（只定義一次）
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("mi", MOUSEINPUT)]

SCRIPTS_DIR = "scripts"
LAST_SCRIPT_FILE = "last_script.txt"
MOUSE_MOVE_THRESHOLD = 50  # px，建議設 20~50，依你需求調整
MOUSE_MOVE_MIN_INTERVAL = 0.2  # 秒
MOUSE_SAMPLE_INTERVAL = 0.01  # 10ms

def format_time(ts):
    return datetime.datetime.fromtimestamp(ts).strftime("%H:%M:%S")

def get_monitor_info():
    monitors = get_monitors()
    return [{
        'x': m.x,
        'y': m.y,
        'width': m.width,
        'height': m.height
    } for m in monitors]

def find_monitor(x, y, monitors):
    for idx, m in enumerate(monitors):
        if m['x'] <= x < m['x'] + m['width'] and m['y'] <= y < m['y'] + m['height']:
            return idx, m
    return 0, monitors[0]  # fallback

class RecorderApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("操作錄製與回放工具")
        self.geometry("900x500")
        self.resizable(False, False)
        self.recording = False
        self.playing = False
        self.paused = False
        self.events = []
        self.speed = 1.0
        self._record_start_time = None
        self._play_start_time = None

        if not os.path.exists(SCRIPTS_DIR):
            os.makedirs(SCRIPTS_DIR)

        # ====== 上方操作區 ======
        frm_top = tb.Frame(self, padding=(10, 10, 10, 5))
        frm_top.pack(fill="x")

        tb.Button(frm_top, text="開始錄製 (F10)", command=self.start_record, bootstyle=PRIMARY, width=14).grid(row=0, column=0, padx=4)
        tb.Button(frm_top, text="暫停/繼續 (F11)", command=self.toggle_pause, bootstyle=INFO, width=14).grid(row=0, column=1, padx=4)
        tb.Button(frm_top, text="停止錄製 (F9)", command=self.stop_record, bootstyle=WARNING, width=14).grid(row=0, column=2, padx=4)
        tb.Button(frm_top, text="回放", command=self.play_record, bootstyle=SUCCESS, width=10).grid(row=0, column=3, padx=4)

        # 新增一排
        frm_bottom = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_bottom.pack(fill="x")
        tb.Label(frm_bottom, text="回放速度:").grid(row=0, column=0, padx=(0,2))
        self.speed_var = tb.StringVar(value="1.0")
        tb.Entry(frm_bottom, textvariable=self.speed_var, width=6).grid(row=0, column=1, padx=2)
        tb.Button(frm_bottom, text="選擇來源", command=self.load_script, bootstyle=SECONDARY, width=10).grid(row=0, column=2, padx=4)
        tb.Button(frm_bottom, text="預設路徑", command=self.open_scripts_dir, bootstyle=SECONDARY, width=10).grid(row=0, column=3, padx=4)

        # 新增：重複次數設定
        frm_repeat = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_repeat.pack(fill="x")
        tb.Label(frm_repeat, text="重複次數:").grid(row=0, column=0, padx=(0,2))
        self.repeat_var = tb.StringVar(value="1")
        tb.Entry(frm_repeat, textvariable=self.repeat_var, width=6).grid(row=0, column=1, padx=2)
        tb.Label(frm_repeat, text="次").grid(row=0, column=2, padx=(0,2))

        # ====== 腳本選單區 ======
        frm_script = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_script.pack(fill="x")
        tb.Label(frm_script, text="腳本選單:").grid(row=0, column=0, sticky="w")
        self.script_var = tk.StringVar()
        self.script_combo = ttk.Combobox(frm_script, textvariable=self.script_var, width=30, state="readonly")
        self.script_combo.grid(row=0, column=1, sticky="w", padx=4)

        # 新增：腳本名稱輸入框與更名按鈕
        self.rename_var = tk.StringVar()
        self.rename_entry = tb.Entry(frm_script, textvariable=self.rename_var, width=20)
        self.rename_entry.grid(row=0, column=2, padx=4)
        tb.Button(frm_script, text="修改腳本名稱", command=self.rename_script, bootstyle=WARNING, width=12).grid(row=0, column=3, padx=4)

        self.script_combo.bind("<<ComboboxSelected>>", self.on_script_selected)
        self.refresh_script_list()

        # ====== 日誌顯示區 ======
        frm_log = tb.Frame(self, padding=(10, 0, 10, 10))
        frm_log.pack(fill="both", expand=True)
        log_title_frame = tb.Frame(frm_log)
        log_title_frame.pack(fill="x")

        # === 滑鼠座標區 ===
        self.mouse_pos_label = tb.Label(log_title_frame, text="( X:0, Y:0 )", font=("Consolas", 12, "bold"), bootstyle=INFO)
        self.mouse_pos_label.pack(side="right", padx=8)

        # === 回放倒數時間區 ===
        self.countdown_label = tb.Label(log_title_frame, text="00:00.0", font=("Consolas", 12, "bold"), foreground="#B22222")  # 深紅色
        self.countdown_label.pack(side="right", padx=8)

        # === 腳本錄製時間區 ===
        self.time_label = tb.Label(log_title_frame, text="00:00.0", font=("Consolas", 12, "bold"), bootstyle=INFO)
        self.time_label.pack(side="right", padx=8)

        tb.Label(log_title_frame, text="動態:", font=("Microsoft JhengHei", 11, "bold")).pack(side="left", anchor="e")
        self.log_text = tb.Text(frm_log, height=24, width=110, state="disabled", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True, pady=(4,0))
        log_scroll = tb.Scrollbar(frm_log, command=self.log_text.yview)
        log_scroll.pack(side="left", fill="y")
        self.log_text.config(yscrollcommand=log_scroll.set)

        # 設定快捷鍵
        keyboard.add_hotkey('F10', self.start_record)
        keyboard.add_hotkey('F9', self.stop_record)
        keyboard.add_hotkey('F11', self.toggle_pause)

        self.load_last_script()
        self.update_mouse_pos()

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def update_time_label(self, seconds):
        m = int(seconds // 60)
        s = seconds % 60
        self.time_label.config(text=f"{m:02d}:{s:04.1f}")

    def start_record(self):
        if self.recording:
            return
        self.events = []
        self.recording = True
        self.paused = False
        self._record_start_time = time.time()
        self.log(f"[{format_time(time.time())}] 開始錄製...")
        self._record_thread_handle = threading.Thread(target=self._record_thread, daemon=True)
        self._record_thread_handle.start()
        self.after(100, self._update_record_time)

    def _update_record_time(self):
        if self.recording:
            now = time.time()
            elapsed = now - self._record_start_time
            self.update_time_label(elapsed)
            self.after(100, self._update_record_time)
        else:
            self.update_time_label(0)

    def toggle_pause(self):
        if self.recording:
            self.paused = not self.paused
            state = "暫停" if self.paused else "繼續"
            self.log(f"[{format_time(time.time())}] 錄製{state}。")
        elif self.playing:
            self.paused = not self.paused
            state = "暫停" if self.paused else "繼續"
            self.log(f"[{format_time(time.time())}] 回放{state}。")

    def _record_thread(self):
        try:
            keyboard.start_recording()
            self._mouse_events = []
            self._recording_mouse = True
            self._record_start_time = time.time()

            from pynput.mouse import Controller
            mouse_ctrl = Controller()
            last_pos = mouse_ctrl.position

            def on_click(x, y, button, pressed):
                if self._recording_mouse and not self.paused:
                    self._mouse_events.append({
                        'type': 'mouse',
                        'event': 'down' if pressed else 'up',
                        'button': str(button).replace('Button.', ''),
                        'x': x,
                        'y': y,
                        'time': time.time()
                    })
            def on_scroll(x, y, dx, dy):
                if self._recording_mouse and not self.paused:
                    self._mouse_events.append({
                        'type': 'mouse',
                        'event': 'wheel',
                        'delta': dy,
                        'x': x,
                        'y': y,
                        'time': time.time()
                    })
            mouse_listener = pynput.mouse.Listener(
                on_click=on_click,
                on_scroll=on_scroll
            )
            mouse_listener.start()

            # 定時採樣滑鼠座標（只記錄有變化的座標）
            while self.recording:
                now = time.time()
                pos = mouse_ctrl.position
                if pos != last_pos:
                    # 錄製時
                    monitors = get_monitor_info()
                    monitor_idx, monitor = find_monitor(pos[0], pos[1], monitors)
                    rel_x = pos[0] - monitor['x']
                    rel_y = pos[1] - monitor['y']
                    self._mouse_events.append({
                        'type': 'mouse',
                        'event': 'move',
                        'monitor': monitor_idx,
                        'rel_x': rel_x,
                        'rel_y': rel_y,
                        'time': now
                    })
                    last_pos = pos
                time.sleep(MOUSE_SAMPLE_INTERVAL)  # 10ms
            self._recording_mouse = False
            mouse_listener.stop()
            k_events = keyboard.stop_recording()

            # 過濾掉錄製快捷鍵（如 F10）事件
            filtered_k_events = [
                e for e in k_events
                if not (e.name == 'f10' and e.event_type in ('down', 'up'))
            ]
            self.events = sorted(
                [{'type': 'keyboard', 'event': e.event_type, 'name': e.name, 'time': e.time} for e in filtered_k_events] +
                self._mouse_events,
                key=lambda e: e['time']
            )
            self.log(f"[{format_time(time.time())}] 錄製完成，共 {len(self.events)} 筆事件。")
            self.log(f"事件預覽: {json.dumps(self.events[:10], ensure_ascii=False, indent=2)}")
        except Exception as ex:
            self.log(f"[{format_time(time.time())}] 錄製執行緒發生錯誤: {ex}")

    def stop_record(self):
        if self.recording:
            self.recording = False
            self.log(f"[{format_time(time.time())}] 停止錄製。")
            # 不要直接 join，改用 after 輪詢
            self._wait_record_thread_finish()

    def _wait_record_thread_finish(self):
        if hasattr(self, '_record_thread_handle') and self._record_thread_handle.is_alive():
            self.after(100, self._wait_record_thread_finish)
        else:
            self.auto_save_script()

    def play_record(self):
        if self.playing or not self.events:
            return
        try:
            self.speed = float(self.speed_var.get())
        except:
            self.speed = 1.0
        self.log(f"[{format_time(time.time())}] 開始回放，速度倍率: {self.speed}")
        self._play_start_time = time.time()
        threading.Thread(target=self._play_thread, daemon=True).start()
        self.after(100, self._update_play_time)

    def _update_play_time(self):
        if self.playing:
            # 以目前回放到的事件時間為主
            idx = getattr(self, "_current_play_index", 0)
            if idx == 0:
                elapsed = 0
            else:
                elapsed = self.events[idx-1]['time'] - self.events[0]['time']
            self.update_time_label(elapsed)
            total = self.events[-1]['time'] - self.events[0]['time'] if self.events else 0
            remain = max(0, total - elapsed)
            self.countdown_label.config(text=f"{int(remain//60):02d}:{remain%60:04.1f}")
            self.after(100, self._update_play_time)
        else:
            self.update_time_label(0)
            self.countdown_label.config(text="00:00.0")

    def _play_thread(self):
        self.playing = True
        self.paused = False
        repeat = 1
        try:
            repeat = int(self.repeat_var.get())
            if repeat < 1:
                repeat = 1
        except:
            repeat = 1
        for _ in range(repeat):
            self._current_play_index = 0
            total_events = len(self.events)
            if total_events == 0:
                break
            base_time = self.events[0]['time']
            play_start = time.time()
            # 取得目前回放時的螢幕資訊
            monitors = get_monitor_info()
            while self._current_play_index < total_events:
                if self.paused:
                    time.sleep(0.05)
                    continue
                i = self._current_play_index
                e = self.events[i]
                event_offset = (e['time'] - base_time) / self.speed
                target_time = play_start + event_offset
                while True:
                    now = time.time()
                    if now >= target_time:
                        break
                    if self.paused:
                        time.sleep(0.05)
                        target_time += 0.05
                        continue
                    time.sleep(min(0.01, target_time - now))
                # 執行事件
                if e['type'] == 'keyboard':
                    if e['event'] == 'down':
                        keyboard.press(e['name'])
                    elif e['event'] == 'up':
                        keyboard.release(e['name'])
                    self.log(f"[{format_time(e['time'])}] 鍵盤: {e['event']} {e['name']}")
                elif e['type'] == 'mouse':
                    if e.get('event') == 'move':
                        # 依據錄製時的螢幕與相對座標，還原絕對座標
                        monitor = monitors[e['monitor']]
                        abs_x = monitor['x'] + e['rel_x']
                        abs_y = monitor['y'] + e['rel_y']
                        move_mouse_abs(abs_x, abs_y)
                    elif e.get('event') == 'down':
                        mouse_event_win('down', button=e.get('button', 'left'))
                        self.log(f"[{format_time(e['time'])}] 滑鼠: {e}")
                    elif e.get('event') == 'up':
                        mouse_event_win('up', button=e.get('button', 'left'))
                        self.log(f"[{format_time(e['time'])}] 滑鼠: {e}")
                self._current_play_index += 1
        self.playing = False
        self.log(f"[{format_time(time.time())}] 回放結束。")

    def refresh_script_list(self):
        try:
            scripts = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith('.json')]
            self.script_combo['values'] = scripts
            self.script_combo.current(0)
        except Exception as e:
            self.log(f"載入腳本列表時發生錯誤: {e}")

    def on_script_selected(self, event):
        script_name = self.script_var.get()
        if script_name:
            self.load_script_file(os.path.join(SCRIPTS_DIR, script_name))

    def load_last_script(self):
        try:
            with open(LAST_SCRIPT_FILE, 'r', encoding='utf-8') as f:
                script_name = f.read().strip()
                if script_name:
                    self.script_var.set(script_name)
                    self.load_script_file(os.path.join(SCRIPTS_DIR, script_name))
        except Exception as e:
            self.log(f"載入最後腳本時發生錯誤: {e}")

    def load_script_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.events = json.load(f)
            self.log(f"載入腳本檔案: {file_path}，事件數量: {len(self.events)}")
        except Exception as e:
            self.log(f"載入腳本檔案時發生錯誤: {e}")

    def auto_save_script(self):
        if not self.recording and not self.playing:
            try:
                script_name = self.script_var.get()
                if not script_name:
                    self.log("無法自動儲存腳本，因為腳本名稱為空。")
                    return
                file_path = os.path.join(SCRIPTS_DIR, script_name)
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.events, f, ensure_ascii=False, indent=2)
                with open(LAST_SCRIPT_FILE, 'w', encoding='utf-8') as f:
                    f.write(script_name)
                self.log(f"自動儲存腳本至: {file_path}")
            except Exception as e:
                self.log(f"自動儲存腳本時發生錯誤: {e}")

    def open_scripts_dir(self):
        try:
            subprocess.Popen(f'explorer "{os.path.abspath(SCRIPTS_DIR)}"')
        except Exception as e:
            self.log(f"開啟腳本資料夾時發生錯誤: {e}")

    def rename_script(self):
        try:
            old_name = self.script_var.get()
            new_name = self.rename_var.get().strip()
            if not old_name or not new_name:
                return
            old_path = os.path.join(SCRIPTS_DIR, old_name)
            new_path = os.path.join(SCRIPTS_DIR, new_name)
            os.rename(old_path, new_path)
            self.log(f"腳本檔案重新命名為: {new_name}")
            self.refresh_script_list()
        except Exception as e:
            self.log(f"更名腳本檔案時發生錯誤: {e}")

    # 加入這個方法
    def load_script(self):
        # 你可以在這裡實作選擇檔案的功能
        file_path = filedialog.askopenfilename(
            title="選擇腳本檔案",
            filetypes=[("JSON Files", "*.json")]
        )
        if file_path:
            self.load_script_file(file_path)

    def update_mouse_pos(self):
        try:
            import pyautogui
            x, y = pyautogui.position()
            self.mouse_pos_label.config(text=f"( X:{x}, Y:{y} )")
        except Exception as e:
            self.mouse_pos_label.config(text="( X:?, Y:? )")
        self.after(100, self.update_mouse_pos)

def move_mouse_abs(x, y):
    # 設定滑鼠移動的絕對位置
    ctypes.windll.user32.SetCursorPos(int(x), int(y))

def mouse_event_win(event, button='left'):
    # 模擬滑鼠事件
    btn = {
        'left': 0x0002,
        'right': 0x0008,
        'middle': 0x0020
    }.get(button, 0x0002)
    if event == 'down':
        ctypes.windll.user32.mouse_event(btn, 0, 0, 0, 0)
    elif event == 'up':
        ctypes.windll.user32.mouse_event(btn | 0x0004, 0, 0, 0, 0)

if __name__ == "__main__":
    app = RecorderApp()
    app.mainloop()