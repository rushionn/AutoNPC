# ...existing code...

import threading

# ...existing code...

def load_script(self):
    path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], initialdir=SCRIPTS_DIR)
    if path:
        def load_job():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    events = json.load(f)
                def update_ui():
                    self.events = events
                    self.log(f"[{format_time(time.time())}] 腳本已載入：{os.path.basename(path)}，共 {len(self.events)} 筆事件。")
                    self.refresh_script_list()
                    self.script_var.set(os.path.basename(path))
                    with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f2:
                        f2.write(os.path.basename(path))
                self.after(0, update_ui)
            except Exception as ex:
                self.after(0, lambda: self.log(f"[{format_time(time.time())}] 載入失敗: {ex}"))
        threading.Thread(target=load_job, daemon=True).start()

def on_script_selected(self, event=None):
    script = self.script_var.get()
    if script:
        path = os.path.join(SCRIPTS_DIR, script)
        def load_job():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    events = json.load(f)
                def update_ui():
                    self.events = events
                    self.log(f"[{format_time(time.time())}] 腳本已載入：{script}，共 {len(self.events)} 筆事件。")
                    with open(LAST_SCRIPT_FILE, "w", encoding="utf-8") as f2:
                        f2.write(script)
                self.after(0, update_ui)
            except Exception as ex:
                self.after(0, lambda: self.log(f"[{format_time(time.time())}] 載入失敗: {ex}"))
        threading.Thread(target=load_job, daemon=True).start()

# ...existing code...
