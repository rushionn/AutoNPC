import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
import threading, time, json, os, datetime
import keyboard, mouse
import ctypes
import win32api
import tkinter.filedialog

# ====== 滑鼠控制函式放在這裡 ======
def move_mouse_abs(x, y):
    ctypes.windll.user32.SetCursorPos(int(x), int(y))

def mouse_event_win(event, x=0, y=0, button='left', delta=0):
    user32 = ctypes.windll.user32
    if not button:
        button = 'left'
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
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
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
        inp.mi = MOUSEINPUT(0, 0, int(delta * 120), 0x0800, 0, None)
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

# ====== RecorderApp 類別與其餘程式碼 ======
SCRIPTS_DIR = "scripts"
LAST_SCRIPT_FILE = "last_script.txt"
LAST_SKIN_FILE = "last_skin.txt"  # 新增這行
MOUSE_SAMPLE_INTERVAL = 0.01  # 10ms

def format_time(ts):
    return datetime.datetime.fromtimestamp(ts).strftime("%H:%M:%S")

class RecorderApp(tb.Window):
    def __init__(self):
        self.user_config = load_user_config()
        skin = self.user_config.get("skin", "darkly")
        super().__init__(themename=skin)
        self._hotkey_handlers = {}
        self.tiny_window = None

        # 統一字體 style
        self.style.configure("My.TButton", font=("Microsoft JhengHei", 9))
        self.style.configure("My.TLabel", font=("Microsoft JhengHei", 9))
        self.style.configure("My.TEntry", font=("Microsoft JhengHei", 9))
        self.style.configure("My.TCombobox", font=("Microsoft JhengHei", 9))
        self.style.configure("My.TCheckbutton", font=("Microsoft JhengHei", 9))
        self.style.configure("TinyBold.TButton", font=("Microsoft JhengHei", 9, "bold"))

        self.title("NPC_2.0   by Lucien")
        self.geometry("900x500")
        self.resizable(False, False)
        self.recording = False
        self.playing = False
        self.paused = False
        self.events = []
        self.speed = 1.0
        self._record_start_time = None
        self._play_start_time = None
        self._total_play_time = 0

        # 設定腳本資料夾
        self.script_dir = self.user_config.get("script_dir", SCRIPTS_DIR)
        if not os.path.exists(self.script_dir):
            os.makedirs(self.script_dir)

        # 快捷鍵設定，新增 tiny
        self.hotkey_map = {
            "start": "F10",
            "pause": "F11",
            "stop": "F9",
            "play": "F12",
            "tiny": "alt+`"
        }

        # ====== 上方操作區 ======
        frm_top = tb.Frame(self, padding=(10, 10, 10, 5))
        frm_top.pack(fill="x")

        self.btn_start = tb.Button(frm_top, text=f"開始錄製 ({self.hotkey_map['start']})", command=self.start_record, bootstyle=PRIMARY, width=14, style="My.TButton")
        self.btn_start.grid(row=0, column=0, padx=4)
        self.btn_pause = tb.Button(frm_top, text=f"暫停/繼續 ({self.hotkey_map['pause']})", command=self.toggle_pause, bootstyle=INFO, width=14, style="My.TButton")
        self.btn_pause.grid(row=0, column=1, padx=4)
        self.btn_stop = tb.Button(frm_top, text=f"停止 ({self.hotkey_map['stop']})", command=self.stop_all, bootstyle=WARNING, width=14, style="My.TButton")
        self.btn_stop.grid(row=0, column=2, padx=4)
        self.btn_play = tb.Button(frm_top, text=f"回放 ({self.hotkey_map['play']})", command=self.play_record, bootstyle=SUCCESS, width=10, style="My.TButton")
        self.btn_play.grid(row=0, column=3, padx=4)

        # ====== skin下拉選單 ======
        themes = ["darkly", "cyborg", "superhero", "journal","minty", "united", "morph", "lumen"]
        self.theme_var = tk.StringVar(value=self.style.theme_use())
        theme_combo = tb.Combobox(frm_top, textvariable=self.theme_var, values=themes, state="readonly", width=12, style="My.TCombobox")
        theme_combo.grid(row=0, column=8, padx=(0, 4), sticky="e")
        theme_combo.bind("<<ComboboxSelected>>", lambda e: self.change_theme())

        # TinyMode 按鈕（粗體，skin下拉選單左側）
        self.tiny_mode_btn = tb.Button(
            frm_top, text="TinyMode", style="TinyBold.TButton",
            command=self.toggle_tiny_mode, width=10
        )
        self.tiny_mode_btn.grid(row=0, column=7, padx=(0, 4), sticky="e")

        # ====== 下方操作區 ======
        frm_bottom = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_bottom.pack(fill="x")
        tb.Label(frm_bottom, text="回放速度:", style="My.TLabel").grid(row=0, column=0, padx=(0,2))
        self.speed_var = tk.StringVar(value=self.user_config.get("speed", "1.0"))
        tb.Entry(frm_bottom, textvariable=self.speed_var, width=6, style="My.TEntry").grid(row=0, column=1, padx=2)
        tb.Button(frm_bottom, text="預設路徑", command=self.use_default_script_dir, bootstyle=SECONDARY, width=10, style="My.TButton").grid(row=0, column=3, padx=4)
        tb.Button(frm_bottom, text="快捷鍵", command=self.open_hotkey_settings, bootstyle=SECONDARY, width=10, style="My.TButton").grid(row=0, column=4, padx=4)

        # ====== 重複次數設定 ======
        frm_repeat = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_repeat.pack(fill="x")
        tb.Label(frm_repeat, text="重複次數:", style="My.TLabel").grid(row=0, column=0, padx=(0,2))
        self.repeat_var = tk.StringVar(value=self.user_config.get("repeat", "1"))
        tb.Entry(frm_repeat, textvariable=self.repeat_var, width=6, style="My.TEntry").grid(row=0, column=1, padx=2)
        tb.Label(frm_repeat, text="次", style="My.TLabel").grid(row=0, column=2, padx=(0,2))

        # ====== 腳本選單區 ======
        frm_script = tb.Frame(self, padding=(10, 0, 10, 5))
        frm_script.pack(fill="x")
        tb.Label(frm_script, text="腳本選單:", style="My.TLabel").grid(row=0, column=0, sticky="w")
        self.script_var = tk.StringVar(value=self.user_config.get("last_script", ""))
        self.script_combo = tb.Combobox(frm_script, textvariable=self.script_var, width=30, state="readonly", style="My.TCombobox")
        self.script_combo.grid(row=0, column=1, sticky="w", padx=4)
        self.rename_var = tk.StringVar()
        self.rename_entry = tb.Entry(frm_script, textvariable=self.rename_var, width=20, style="My.TEntry")
        self.rename_entry.grid(row=0, column=2, padx=4)
        tb.Button(frm_script, text="修改腳本名稱", command=self.rename_script, bootstyle=WARNING, width=12, style="My.TButton").grid(row=0, column=3, padx=4)

        self.script_combo.bind("<<ComboboxSelected>>", self.on_script_selected)
        self.refresh_script_list()
        if self.script_var.get():
            self.on_script_selected()

        # ====== 日誌顯示區 ======
        frm_log = tb.Frame(self, padding=(10, 0, 10, 10))
        frm_log.pack(fill="both", expand=True)
        log_title_frame = tb.Frame(frm_log)
        log_title_frame.pack(fill="x")

        self.mouse_pos_label = tb.Label(
            log_title_frame, text="( X:0, Y:0 )",
            font=("Consolas", 12, "bold"),
            foreground="#668B9B"
        )
        self.mouse_pos_label.pack(side="left", padx=8)
        self.time_label = tb.Label(log_title_frame, text="錄製耗時: 00:00.0", font=("Consolas", 12, "bold"), foreground="#15D3BD")
        self.time_label.pack(side="right", padx=8)
        self.countdown_label = tb.Label(log_title_frame, text="單組剩餘: 00:00.0", font=("Consolas", 12, "bold"), foreground="#DB0E59")
        self.countdown_label.pack(side="right", padx=8)
        self.total_time_label = tb.Label(log_title_frame, text="總共剩餘: 00:00.0", font=("Consolas", 12, "bold"), foreground="#FF95CA")
        self.total_time_label.pack(side="right", padx=8)

        self.log_text = tb.Text(frm_log, height=24, width=110, state="disabled", font=("Microsoft JhengHei", 9))
        self.log_text.pack(fill="both", expand=True, pady=(4,0))
        log_scroll = tb.Scrollbar(frm_log, command=self.log_text.yview)
        log_scroll.pack(side="left", fill="y")
        self.log_text.config(yscrollcommand=log_scroll.set)

        # ====== 其餘初始化 ======
        self.after(100, self._delayed_init)

    def _delayed_init(self):
        self._register_hotkeys()
        self.refresh_script_list()
        self.after(100, self.load_last_script)
        self.after(200, self.update_mouse_pos)

    def save_config(self):
        self.user_config["skin"] = self.theme_var.get()
        self.user_config["last_script"] = self.script_var.get()
        self.user_config["repeat"] = self.repeat_var.get()
        self.user_config["speed"] = self.speed_var.get()
        self.user_config["script_dir"] = self.script_dir
        save_user_config(self.user_config)

    def change_theme(self):
        self.style.theme_use(self.theme_var.get())
        self.save_config()

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
        if self.recording or self.playing:
            self.paused = not self.paused
            state = "暫停" if self.paused else "繼續"
            mode = "錄製" if self.recording else "回放"
            self.log(f"[{format_time(time.time())}] {mode}{state}。")

    def _record_thread(self):
        try:
            self._mouse_events = []
            self._recording_mouse = True
            self._record_start_time = time.time()

            # 先啟動 keyboard 與 mouse 監聽
            keyboard.start_recording()

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
            import pynput.mouse
            mouse_listener = pynput.mouse.Listener(
                on_click=on_click,
                on_scroll=on_scroll
            )
            mouse_listener.start()

            # 立即記錄當下滑鼠位置（避免剛開始沒動作時漏記）
            now = time.time()
            self._mouse_events.append({
                'type': 'mouse',
                'event': 'move',
                'x': last_pos[0],
                'y': last_pos[1],
                'time': now
            })

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
                time.sleep(MOUSE_SAMPLE_INTERVAL)
            self._recording_mouse = False
            mouse_listener.stop()
            k_events = keyboard.stop_recording()

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
            self._wait_record_thread_finish()

    def stop_all(self):
        stopped = False
        if self.recording:
            self.recording = False
            stopped = True
            self.log(f"[{format_time(time.time())}] 停止錄製。")
            self._wait_record_thread_finish()
        if self.playing:
            self.playing = False
            stopped = True
            self.log(f"[{format_time(time.time())}] 停止回放。")
        if not stopped:
            self.log(f"[{format_time(time.time())}] 無進行中動作可停止。")

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
        try:
            repeat = int(self.repeat_var.get())
            if repeat < 1:
                repeat = 1
        except:
            repeat = 1
        if self.events:
            total = (self.events[-1]['time'] - self.events[0]['time']) / self.speed
            self._total_play_time = total * repeat
        else:
            self._total_play_time = 0
        self._play_start_time = time.time()
        self._play_total_time = self._total_play_time
        self.update_total_time_label(self._total_play_time)
        self.log(f"[{format_time(time.time())}] 開始回放，速度倍率: {self.speed}")
        self.playing = True
        self.paused = False
        threading.Thread(target=self._play_thread, daemon=True).start()
        self.after(100, self._update_play_time)

    def _update_play_time(self):
        if self.playing:
            idx = getattr(self, "_current_play_index", 0)
            if idx == 0:
                elapsed = 0
            else:
                elapsed = self.events[idx-1]['time'] - self.events[0]['time']
            self.update_time_label(elapsed)
            total = self.events[-1]['time'] - self.events[0]['time'] if self.events else 0
            remain = max(0, total - elapsed)
            self.countdown_label.config(text=f"{int(remain//60):02d}:{remain%60:04.1f}")
            if hasattr(self, "_play_start_time"):
                total_remain = max(0, self._total_play_time - (time.time() - self._play_start_time))
                self.update_total_time_label(total_remain)
            self.after(100, self._update_play_time)
        else:
            self.update_time_label(0)
            self.countdown_label.config(text="00:00.0")
            self.update_total_time_label(0)

    def update_total_time_label(self, seconds):
        m = int(seconds // 60)
        s = seconds % 60
        self.total_time_label.config(text=f"總運作: {m:02d}:{s:04.1f}")

    def _play_thread(self):
        self.playing = True
        self.paused = False
        try:
            repeat = int(self.repeat_var.get())
            if repeat < 1:
                repeat = 1
        except:
            repeat = 1
        for _ in range(repeat):
            self._current_play_index = 0
            total_events = len(self.events)
            if total_events == 0 or not self.playing:
                break
            base_time = self.events[0]['time']
            play_start = time.time()
            while self._current_play_index < total_events:
                if not self.playing:
                    break  # 強制中斷回放
                while self.paused:
                    if not self.playing:
                        break  # 強制中斷回放
                    time.sleep(0.05)
                    play_start += 0.05  # 修正暫停期間的基準時間
                if not self.playing:
                    break  # 強制中斷回放
                i = self._current_play_index
                e = self.events[i]
                event_offset = (e['time'] - base_time) / self.speed
                target_time = play_start + event_offset
                while True:
                    now = time.time()
                    if not self.playing:
                        break  # 強制中斷回放
                    if now >= target_time:
                        break
                    if self.paused:
                        if not self.playing:
                            break
                        time.sleep(0.05)
                        target_time += 0.05
                        continue
                    time.sleep(min(0.01, target_time - now))
                if not self.playing:
                    break  # 強制中斷回放
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
            if not self.playing:
                break  # 強制中斷回放
        self.playing = False
        self.paused = False
        self.log(f"[{format_time(time.time())}] 回放結束。")

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
            ts = datetime.datetime.now().strftime("%Y_%m%d_%H%M_%S")
            filename = f"{ts}.json"
            path = os.path.join(self.script_dir, filename)
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
        from tkinter import filedialog
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
        self.save_config()

    def refresh_script_list(self):
        files = [f for f in os.listdir(self.script_dir) if f.endswith('.json')]
        self.script_combo['values'] = files
        # 若目前選擇的腳本不存在於新資料夾，清空選擇
        if self.script_var.get() not in files:
            self.script_var.set("")

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

    def open_hotkey_settings(self):
        win = tb.Toplevel(self)
        win.title("快捷鍵設定")
        win.geometry("340x280")
        win.resizable(False, False)

        labels = {
            "start": "開始錄製",
            "pause": "暫停/繼續",
            "stop": "停止錄製",
            "play": "回放",
            "tiny": "TinyMode"
        }
        vars = {}
        entries = {}
        row = 0

        # 用於記錄目前按下的組合鍵
        pressed_keys = {}

        def on_entry_key(event, key, var):
            key_name = event.keysym.upper()
            # 如果是 F1~F12，直接設為 F1~F12
            if key_name in [f"F{i}" for i in range(1, 13)]:
                var.set(key_name)
            else:
                keys = []
                if event.state & 0x0001: keys.append("SHIFT")
                if event.state & 0x0004: keys.append("CTRL")
                if event.state & 0x0008: keys.append("ALT")
                # 避免重複加入
                if key_name not in keys and key_name not in ("SHIFT", "CONTROL", "ALT", "CAPS_LOCK"):
                    keys.append(key_name)
                var.set("+".join(keys))
            return "break"

        def on_entry_focus_in(event, var):
            var.set("請按下組合鍵")

        def on_entry_focus_out(event, key, var):
            # 若未輸入組合鍵則還原為原本設定
            if var.get() == "請按下組合鍵" or not var.get():
                var.set(self.hotkey_map[key])

        for key, label in labels.items():
            tb.Label(win, text=label, font=("Microsoft JhengHei", 11)).grid(row=row, column=0, padx=10, pady=8, sticky="w")
            var = tk.StringVar(value=self.hotkey_map[key])
            entry = tb.Entry(win, textvariable=var, width=16, font=("Consolas", 11), state="normal")
            entry.grid(row=row, column=1, padx=10)
            vars[key] = var
            entries[key] = entry
            # 綁定事件
            entry.bind("<KeyPress>", lambda e, k=key, v=var: on_entry_key(e, k, v))
            entry.bind("<FocusIn>", lambda e, v=var: on_entry_focus_in(e, v))
            entry.bind("<FocusOut>", lambda e, k=key, v=var: on_entry_focus_out(e, k, v))
            row += 1

        def save_and_apply():
            for key in self.hotkey_map:
                val = vars[key].get()
                # TinyMode 可任意組合，其餘僅允許 F1~F12
                if key != "tiny":
                    # 僅允許 F1~F12，且不能有組合鍵
                    if val and val != "請按下組合鍵":
                        if val.upper() in [f"F{i}" for i in range(1, 13)]:
                            self.hotkey_map[key] = val.upper()
                        else:
                            self.log(f"{labels[key]} 只能設定為 F1~F12，請重新設定。")
                            return
                else:
                    # TinyMode 可任意組合
                    if val and val != "請按下組合鍵":
                        self.hotkey_map[key] = val.upper()
            self._register_hotkeys()
            self._update_hotkey_labels()
            self.log("快捷鍵設定已更新。")
            win.destroy()  # 確保最後才關閉視窗

        tb.Button(win, text="儲存", command=save_and_apply, width=10, bootstyle=SUCCESS).grid(row=row, column=0, columnspan=2, pady=16)

    # 不再需要 _make_hotkey_entry_handler

    def _register_hotkeys(self):
        for handler in self._hotkey_handlers.values():
            try:
                keyboard.remove_hotkey(handler)
            except Exception:
                pass
        self._hotkey_handlers.clear()
        for key, hotkey in self.hotkey_map.items():
            try:
                handler = keyboard.add_hotkey(hotkey, getattr(self, {
                    "start": "start_record",
                    "pause": "toggle_pause",
                    "stop": "stop_all",
                    "play": "play_record",
                    "tiny": "toggle_tiny_mode"  # <--- 加這行
                }[key]))
                self._hotkey_handlers[key] = handler
            except Exception as ex:
                self.log(f"快捷鍵 {hotkey} 註冊失敗: {ex}")

    def _update_hotkey_labels(self):
        self.btn_start.config(text=f"開始錄製 ({self.hotkey_map['start']})")
        self.btn_pause.config(text=f"暫停/繼續 ({self.hotkey_map['pause']})")
        self.btn_stop.config(text=f"停止 ({self.hotkey_map['stop']})")
        self.btn_play.config(text=f"回放 ({self.hotkey_map['play']})")
        # TinyMode 按鈕同步更新
        if hasattr(self, "tiny_btns"):
            for btn, icon, key in self.tiny_btns:
                btn.config(text=f"{icon} {self.hotkey_map[key]}")

    def toggle_tiny_mode(self):
        # 切換 TinyMode 狀態
        if not hasattr(self, "tiny_mode_on"):
            self.tiny_mode_on = False
        self.tiny_mode_on = not self.tiny_mode_on
        if self.tiny_mode_on:
            if self.tiny_window is None or not self.tiny_window.winfo_exists():
                self.tiny_window = tb.Toplevel(self)
                self.tiny_window.title("NPC TinyMode")
                self.tiny_window.geometry("470x40")
                self.tiny_window.overrideredirect(True)
                self.tiny_window.resizable(False, False)
                self.tiny_window.attributes("-topmost", True)
                self.tiny_btns = []
                # 拖曳功能
                self.tiny_window.bind("<ButtonPress-1>", self._start_move_tiny)
                self.tiny_window.bind("<B1-Motion>", self._move_tiny)
                btn_defs = [
                    ("⏺", "start"),
                    ("⏸", "pause"),
                    ("⏹", "stop"),
                    ("▶︎", "play"),
                    ("⤴︎", "tiny")
                ]
                for i, (icon, key) in enumerate(btn_defs):
                    btn = tb.Button(
                        self.tiny_window,
                        text=f"{icon} {self.hotkey_map[key]}",
                        width=7, style="My.TButton",
                        command=getattr(self, {
                            "start": "start_record",
                            "pause": "toggle_pause",
                            "stop": "stop_all",
                            "play": "play_record",
                            "tiny": "toggle_tiny_mode"
                        }[key])
                    )
                    btn.grid(row=0, column=i, padx=2, pady=5)
                    self.tiny_btns.append((btn, icon, key))
                self.tiny_window.protocol("WM_DELETE_WINDOW", self._close_tiny_mode)
                self.withdraw()
        else:
            self._close_tiny_mode()

    def _close_tiny_mode(self):
        if self.tiny_window and self.tiny_window.winfo_exists():
            self.tiny_window.destroy()
        self.tiny_mode_on = False
        self.deiconify()

    def _start_move_tiny(self, event):
        self._tiny_drag_x = event.x
        self._tiny_drag_y = event.y

    def _move_tiny(self, event):
        x = self.tiny_window.winfo_x() + event.x - self._tiny_drag_x
        y = self.tiny_window.winfo_y() + event.y - self._tiny_drag_y
        self.tiny_window.geometry(f"+{x}+{y}")

    def use_default_script_dir(self):
        self.script_dir = SCRIPTS_DIR
        if not os.path.exists(self.script_dir):
            os.makedirs(self.script_dir)
        self.refresh_script_list()
        self.save_config()
        # 開啟資料夾
        os.startfile(self.script_dir)

CONFIG_FILE = "user_config.json"

def load_user_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # 預設值
    return {
        "skin": "darkly",
        "last_script": "",
        "repeat": "1",
        "speed": "1.0",
        "script_dir": SCRIPTS_DIR
    }

def save_user_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

if __name__ == "__main__":
    app = RecorderApp()
    app.mainloop()