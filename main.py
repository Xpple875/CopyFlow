import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pyperclip
import re
import sys
import time
import threading
import json
import os
from pynput import keyboard
from pynput.keyboard import Controller

# --- SYSTEM CONSTANTS ---
VERSION = "1.3.2"
CACHE_PATH = "copyflow_data.json"
STATS_PATH = "copyflow_stats.json"
PRIMARY_ACCENT = "#3b82f6"
BG_GRAY = "#1e1e1e"
PANEL_GRAY = "#252525"
TEXT_MAIN = "#ffffff"
TEXT_DIM = "#a0a0a0"

class StatsEngine:
    @staticmethod
    def log_item():
        stats = StatsEngine.get_stats()
        stats["total_processed"] += 1
        with open(STATS_PATH, "w") as f:
            json.dump(stats, f)

    @staticmethod
    def get_stats():
        if os.path.exists(STATS_PATH):
            try:
                with open(STATS_PATH, "r") as f: return json.load(f)
            except: pass
        return {"total_processed": 0}

class CopyFlowApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"CopyFlow Pro — {VERSION}")
        self.geometry("1300x950")
        self.configure(fg_color=BG_GRAY)

        self.queue = []
        self.undo_stack = []
        self.kb = Controller()
        self._f9_lock = threading.Lock()
        self.batch_running = False
        self.search_query = ""

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_sidebar()
        self.setup_main_area()
        self.setup_footer()
        self.start_hotkey_engine()
        self.load_data()
        self.protocol("WM_DELETE_WINDOW", self.on_exit)

    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=320, corner_radius=0, fg_color=PANEL_GRAY)
        self.sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")

        ctk.CTkLabel(self.sidebar, text="CopyFlow", font=("Inter", 36, "bold"), text_color=PRIMARY_ACCENT).pack(pady=(40, 20))

        self.add_side_label("AUTOMATION")
        self.auto_tab = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.sidebar, text="Auto-Tab (Excel Mode)", variable=self.auto_tab, font=("Inter", 12)).pack(pady=5, padx=25, anchor="w")

        ctk.CTkLabel(self.sidebar, text="Batch Speed (Delay)", font=("Inter", 11), text_color="#555").pack(anchor="w", padx=25)
        self.batch_delay = ctk.CTkSlider(self.sidebar, from_=0.1, to=2.0, progress_color=PRIMARY_ACCENT)
        self.batch_delay.set(0.4)
        self.batch_delay.pack(pady=10, padx=25, fill="x")

        self.btn_batch = ctk.CTkButton(self.sidebar, text="▶ START BATCH RUN", fg_color="#1e3a8a", hover_color="#1d4ed8",
                                       font=("Inter", 14, "bold"), height=45, command=self.trigger_batch_run)
        self.btn_batch.pack(pady=10, padx=25, fill="x")

        self.add_side_label("STRATEGY")
        self.mode_var = ctk.StringVar(value="Normal")
        ctk.CTkOptionMenu(self.sidebar, values=["Lenient", "Normal", "Strict"], variable=self.mode_var, fg_color="#333").pack(pady=5, padx=25, fill="x")

        self.add_side_label("UTILITIES")
        ctk.CTkButton(self.sidebar, text="Undo Last Split", fg_color="#333", command=self.handle_undo).pack(pady=5, padx=25, fill="x")
        ctk.CTkButton(self.sidebar, text="Lifetime Stats", fg_color="#333", command=self.show_stats).pack(pady=5, padx=25, fill="x")
        ctk.CTkButton(self.sidebar, text="Help & Manual", fg_color="#333", command=self.open_help).pack(pady=5, padx=25, fill="x")

        ctk.CTkButton(self.sidebar, text="Reset & Wipe", fg_color="#442222", hover_color="#ef4444", command=self.clear_all).pack(side="bottom", pady=40, padx=25, fill="x")

    def add_side_label(self, text):
        ctk.CTkLabel(self.sidebar, text=text, font=("Inter", 11, "bold"), text_color="#666").pack(anchor="w", padx=25, pady=(20, 5))

    def setup_main_area(self):
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=1, sticky="nsew", padx=40, pady=40)
        self.main.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="Queue", font=("Inter", 42, "bold"), text_color=TEXT_MAIN).pack(side="left")
        self.prog = ctk.CTkProgressBar(header, width=250, progress_color=PRIMARY_ACCENT)
        self.prog.set(0)
        self.prog.pack(side="right", padx=10)

        self.input_box = ctk.CTkTextbox(self.main, height=160, font=("Inter", 15), border_width=1, border_color="#333", fg_color="#121212")
        self.input_box.grid(row=1, column=0, sticky="ew", pady=(20, 10))

        self.btn_add = ctk.CTkButton(self.main, text="Process & Add to Queue", height=60, font=("Inter", 18, "bold"),
                                     fg_color=PRIMARY_ACCENT, hover_color="#2563eb", command=self.process_text)
        self.btn_add.grid(row=2, column=0, sticky="ew", pady=(0, 20))

        status_bar = ctk.CTkFrame(self.main, fg_color="transparent")
        status_bar.grid(row=3, column=0, sticky="ew", pady=(0, 15))
        self.search_entry = ctk.CTkEntry(status_bar, placeholder_text="Filter items...", width=280, fg_color="#181818")
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", self.filter_queue)
        self.hk_lbl = ctk.CTkLabel(status_bar, text="● F9 LISTENER: ACTIVE", text_color="#10b981", font=("Inter", 11, "bold"))
        self.hk_lbl.pack(side="right")

        self.scroll = ctk.CTkScrollableFrame(self.main, label_text="ACTIVE PIPELINE", fg_color="#1a1a1a", border_color="#222")
        self.scroll.grid(row=4, column=0, sticky="nsew")
        self.main.grid_rowconfigure(4, weight=1)

    def setup_footer(self):
        self.footer = ctk.CTkFrame(self, height=35, fg_color="#111", corner_radius=0)
        self.footer.grid(row=1, column=1, sticky="ew")
        self.stat_lbl = ctk.CTkLabel(self.footer, text="Engine Status: Optimal", font=("Inter", 11), text_color="#555")
        self.stat_lbl.pack(side="left", padx=25)

    def process_text(self):
        raw = self.input_box.get("1.0", "end-1c").strip()
        if not raw: return
        self.undo_stack.append(json.loads(json.dumps(self.queue)))

        mode = self.mode_var.get()
        pattern = r'[;,|\t•●▪▫◦→⏳✔✖\n]'
        if mode == "Strict": pattern = r'[;,|\t•●▪▫◦→⏳✔✖\n\s\/]'
        elif mode == "Lenient": pattern = r'[\t\n|]'

        items = [i.strip() for i in re.split(pattern, raw) if i.strip()]

        for t in items:
            # --- SANITIZATION LOGIC ---
            # Removes: "1)", "1.", "- ", "• ", "a)", "A." at the start of strings
            clean_t = re.sub(r'^([a-zA-Z0-9][\)\.]|[\-\•\*\●\▪\▫\◦\→])\s*', '', t)
            clean_t = clean_t.strip()
            if clean_t:
                self.queue.append({"text": clean_t, "done": False})

        self.input_box.delete("1.0", "end")
        self.refresh_ui()
        self.save_data()

    def refresh_ui(self):
        for w in self.scroll.winfo_children(): w.destroy()
        done_c = sum(1 for x in self.queue if x['done'])
        for i, item in enumerate(self.queue):
            if self.search_query and self.search_query not in item['text'].lower(): continue
            row = ctk.CTkFrame(self.scroll, fg_color="transparent")
            row.pack(fill="x", pady=2, padx=10)
            btn = ctk.CTkButton(row, text=f"{'✓ ' if item['done'] else ''}{item['text']}",
                                anchor="w", height=38,
                                fg_color="#2a2a2a" if not item['done'] else "#151515",
                                text_color=TEXT_MAIN if not item['done'] else "#444",
                                command=lambda x=i: self.mark_done(x))
            btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
            ctk.CTkButton(row, text="✕", width=38, height=38, fg_color="transparent",
                          command=lambda x=i: self.delete_item(x)).pack(side="right")
        if self.queue: self.prog.set(done_c / len(self.queue))
        self.stat_lbl.configure(text=f"Queue: {len(self.queue)} | Done: {done_c}")

    def filter_queue(self, e=None):
        self.search_query = self.search_entry.get().lower()
        self.refresh_ui()

    def start_hotkey_engine(self):
        self.hk_listener = keyboard.GlobalHotKeys({'<f9>': self.on_f9_press})
        self.hk_listener.start()

    def on_f9_press(self):
        if not self.batch_running:
            threading.Thread(target=self.run_single, daemon=True).start()

    def run_single(self):
        if not self._f9_lock.acquire(blocking=False): return
        try:
            for i, item in enumerate(self.queue):
                if not item['done']:
                    self.execute_type(item['text'])
                    self.after(0, lambda x=i: self.mark_done(x))
                    break
        finally: self._f9_lock.release()

    def trigger_batch_run(self):
        if self.batch_running:
            self.batch_running = False
            return
        if messagebox.askyesno("Batch Mode", "Start 3s countdown? Switch to Excel/Target app now."):
            self.batch_running = True
            self.btn_batch.configure(text="⏹ STOP BATCH", fg_color="#991b1b")
            threading.Thread(target=self.batch_worker, daemon=True).start()

    def batch_worker(self):
        time.sleep(3)
        delay = self.batch_delay.get()
        for i, item in enumerate(self.queue):
            if not self.batch_running: break
            if not item['done']:
                self.execute_type(item['text'])
                self.after(0, lambda x=i: self.mark_done(x))
                time.sleep(delay)
        self.batch_running = False
        self.after(0, lambda: self.btn_batch.configure(text="▶ START BATCH RUN", fg_color="#1e3a8a"))

    def execute_type(self, text):
        self.kb.type(text)
        if self.auto_tab.get():
            time.sleep(0.05)
            self.kb.press(keyboard.Key.tab)
            self.kb.release(keyboard.Key.tab)

    def show_stats(self):
        s = StatsEngine.get_stats()
        messagebox.showinfo("Lifetime Stats", f"Total Items Processed: {s['total_processed']}\nEstimated Time Saved: {round(s['total_processed']*2/60, 2)} minutes")

    def open_help(self):
        win = ctk.CTkToplevel(self)
        win.title("CopyFlow Master Manual")
        win.geometry("750x850")
        win.attributes("-topmost", True)

        txt = ctk.CTkTextbox(win, font=("Inter", 14), wrap="word")
        txt.pack(fill="both", expand=True, padx=30, pady=30)

        guide = """
COPYFLOW PRO — USER MANUAL
==========================

1. QUICK START
--------------
• Paste data -> Click 'Process' -> Go to your app -> Tap F9 to paste!

2. SMART CLEANING (NEW)
-----------------------
CopyFlow now automatically removes list artifacts.
Example: If you paste '1) USB Cable', the app saves just 'USB Cable'.
Supported: 1., 1), a., a), •, -, *, and more.

3. SPLITTING MODES
------------------
• NORMAL: Keeps phrases like 'New York' as one item.
• STRICT: Forces a split at every single space or dash.
• LENIENT: Only splits when it sees a new line.

4. BATCH AUTOMATION
-------------------
• Enable 'Auto-Tab' for Excel.
• Use 'Start Batch Run' to automate 100+ items without touching your mouse.
• Adjust 'Batch Speed' if the target app is slow.

5. SAFETY
---------
• Click 'Panic STOP' to kill a batch run.
• Close the app to stop all keyboard listening.
        """
        txt.insert("1.0", guide)
        txt.configure(state="disabled")

    def mark_done(self, idx):
        if not self.queue[idx]['done']:
            self.queue[idx]['done'] = True
            StatsEngine.log_item()
        self.refresh_ui()
        self.save_data()

    def delete_item(self, idx):
        self.queue.pop(idx)
        self.refresh_ui()
        self.save_data()

    def handle_undo(self):
        if self.undo_stack:
            self.queue = self.undo_stack.pop()
            self.refresh_ui()

    def clear_all(self):
        if messagebox.askyesno("Confirm", "Wipe queue and cache?"):
            self.queue = []
            self.refresh_ui()
            if os.path.exists(CACHE_PATH): os.remove(CACHE_PATH)

    def save_data(self):
        try:
            with open(CACHE_PATH, 'w') as f: json.dump(self.queue, f)
        except: pass

    def load_data(self):
        if os.path.exists(CACHE_PATH):
            try:
                with open(CACHE_PATH, 'r') as f: self.queue = json.load(f)
                self.refresh_ui()
            except: pass

    def on_exit(self):
        if self.hk_listener: self.hk_listener.stop()
        self.save_data()
        self.destroy()
        sys.exit(0)

if __name__ == "__main__":
    app = CopyFlowApp()
    app.mainloop()
