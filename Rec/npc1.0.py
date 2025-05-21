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
import win32api

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

class RecorderApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("NPC_1.0   by Lucien")
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
                    self._mouse_events.append({
                        'type': 'mouse',
                        'event': 'move',
                        'x': pos[0],
                        'y': pos[1],
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
            while self._current_play_index < total_events:
                if self.paused:
                    time.sleep(0.05)
                    continue
                i = self._current_play_index
                e = self.events[i]
                # 計算應該執行的目標時間點
                event_offset = (e['time'] - base_time) / self.speed
                target_time = play_start + event_offset
                # 精準等待到該事件時間
                while True:
                    now = time.time()
                    if now >= target_time:
                        break
                    if self.paused:
                        time.sleep(0.05)
                        target_time += 0.05  # 修正暫停時的目標時間
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
                        move_mouse_abs(e['x'], e['y'])
                    elif e.get('event') == 'down':
                        mouse_event_win('down', button=e.get('button', 'left'))
                        self.log(f"[{format_time(e['time'])}] 滑鼠: {e}")
                    elif e.get('event') == 'up':
                        mouse_event_win('up', button=e.get('button', 'left'))
                        self.log(f"[{format_time(e['time'])}] 滑鼠: {e}")
                    elif e.get('event') == 'wheel':
                        mouse_event_win('wheel', delta=e.get('delta', 0))
                        self.log(f"[{format_time(e['time'])}] 滑鼠: {e}")
                self._current_play_index += 1
        self.playing = False
        self.paused = False
        self.log(f"[{format_time(time.time())}] 回放結束。")

    # 日誌與事件模組化
    def get_events_json(self):
        return json.dumps(self.events, ensure_ascii=False, indent=2)

    def set_events_json(self, json_str):
        try:
            self.events = json.loads(json_str)
            self.log(f"[{format_time(time.time())}] 已從 JSON 載入 {len(self.events)} 筆事件。")
        except Exception as e:
            self.log(f"[{format_time(time.time())}] JSON 載入失敗: {e}")

    def auto_save_script(self):
        try:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"record_{ts}.json"
            path = os.path.join(SCRIPTS_DIR, filename)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.events, f, ensure_ascii=False, indent=2)
            self.log(f"[{format_time(time.time())}] 自動存檔：{filename}，事件數：{len(self.events)}")
            self.refresh_script_list()
            self.script_var.set(filename)
            with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f:
                f.write(filename)
        except Exception as ex:
            self.log(f"[{format_time(time.time())}] 存檔失敗: {ex}")

    def load_script(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], initialdir=SCRIPTS_DIR)
        if path:
            with open(path, "r", encoding="utf-8") as f:
                self.events = json.load(f)
            self.log(f"[{format_time(time.time())}] 腳本已載入：{os.path.basename(path)}，共 {len(self.events)} 筆事件。")
            self.refresh_script_list()
            self.script_var.set(os.path.basename(path))
            with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f:
                f.write(os.path.basename(path))

    def on_script_selected(self, event=None):
        script = self.script_var.get()
        if script:
            path = os.path.join(SCRIPTS_DIR, script)
            with open(path, "r", encoding="utf-8") as f:
                self.events = json.load(f)
            self.log(f"[{format_time(time.time())}] 腳本已載入：{script}，共 {len(self.events)} 筆事件。")
            with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f:
                f.write(script)

    def refresh_script_list(self):
        files = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith('.json')]
        self.script_combo['values'] = files

    def load_last_script(self):
        if os.path.exists(LAST_SCRIPT_FILE):
            with open(LAST_SCRIPT_FILE, "r", encoding="utf-8") as f:
                last_script = f.read().strip()
            if last_script:
                script_path = os.path.join(SCRIPTS_DIR, last_script)
                if os.path.exists(script_path):
                    with open(script_path, "r", encoding="utf-8") as f:
                        self.events = json.load(f)
                    self.script_var.set(last_script)
                    self.log(f"[{format_time(time.time())}] 已自動載入上次腳本：{last_script}，共 {len(self.events)} 筆事件。")

    def update_mouse_pos(self):
        try:
            x, y = mouse.get_position()
            self.mouse_pos_label.config(text=f"( X:{x}, Y:{y} )")
        except Exception:
            self.mouse_pos_label.config(text="( X:?, Y:? )")
        self.after(100, self.update_mouse_pos)

    def rename_script(self):
        old_name = self.script_var.get()
        new_name = self.rename_var.get().strip()
        if not old_name or not new_name:
            self.log(f"[{format_time(time.time())}] 請選擇腳本並輸入新名稱。")
            return
        if not new_name.endswith('.json'):
            new_name += '.json'
        old_path = os.path.join(SCRIPTS_DIR, old_name)
        new_path = os.path.join(SCRIPTS_DIR, new_name)
        if os.path.exists(new_path):
            self.log(f"[{format_time(time.time())}] 檔案已存在，請換個名稱。")
            return
        try:
            os.rename(old_path, new_path)
            self.log(f"[{format_time(time.time())}] 腳本已更名為：{new_name}")
            self.refresh_script_list()
            self.script_var.set(new_name)
            with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f:
                f.write(new_name)
        except Exception as e:
            self.log(f"[{format_time(time.time())}] 更名失敗: {e}")

    def open_scripts_dir(self):
        path = os.path.abspath(SCRIPTS_DIR)
        os.startfile(path)

    def _on_mouse_event(self, event):
        if self.recording and not self.paused:
            if event.__class__.__name__ == "ButtonEvent":
                self.log(f"滑鼠事件: {event.event_type} {getattr(event, 'button', 'left')} @ ({event.x},{event.y})")
                self._mouse_events.append({
                    'type': 'mouse',
                    'event': event.event_type,
                    'button': getattr(event, 'button', 'left'),
                    'x': event.x,
                    'y': event.y,
                    'time': time.time(),
                })
            # 滑鼠滾輪事件
            elif event.__class__.__name__ == "WheelEvent":
                self._mouse_events.append({
                    'type': 'mouse',
                    'event': 'wheel',
                    'delta': event.delta,
                    'x': event.x,
                    'y': event.y,
                    'time': time.time(),
                })
            # 滑鼠移動事件（通常不需額外記錄，已由 move_watcher 處理）
            # elif event.__class__.__name__ == "MoveEvent":
            #     pass

# Windows API 滑鼠控制
def move_mouse_abs(x, y):
    # 直接用全域座標移動滑鼠
    ctypes.windll.user32.SetCursorPos(int(x), int(y))

def move_mouse_abs_safe(x, y):
    # 取得所有螢幕範圍，避免滑鼠超出
    from screeninfo import get_monitors
    min_x = min(m.x for m in get_monitors())
    min_y = min(m.y for m in get_monitors())
    max_x = max(m.x + m.width for m in get_monitors())
    max_y = max(m.y + m.height for m in get_monitors())
    x = max(min_x, min(x, max_x - 1))
    y = max(min_y, min(y, max_y - 1))
    ctypes.windll.user32.SetCursorPos(int(x), int(y))

def mouse_event_win(event, x=0, y=0, button='left', delta=0):
    user32 = ctypes.windll.user32
    if event == 'down' or event == 'up':
        flags = {'left': (0x0002, 0x0004), 'right': (0x0008, 0x0010), 'middle': (0x0020, 0x0040)}
        flag = flags.get(button, (0x0002, 0x0004))[0 if event == 'down' else 1]
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
        inp = INPUT()
        inp.type = 0
        inp.mi = MOUSEINPUT(0, 0, 0, flag, 0, None)
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    elif event == 'wheel':
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
        inp = INPUT()
        inp.type = 0
        inp.mi = MOUSEINPUT(0, 0, int(delta * 120), 0x0800, 0, None)  # 0x0800 = WHEEL
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

if __name__ == "__main__":
    app = RecorderApp()
    app.mainloop()