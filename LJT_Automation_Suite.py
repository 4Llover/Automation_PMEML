"""
LJT Automation Suite v1.5.0
Author: LJT
Copyright 2026 LJT All Rights Reserved
"""

import os, sys, time, threading, subprocess, warnings
import pandas as pd
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
    ctk.set_appearance_mode("Light")
except ImportError:
    print("Please install customtkinter: pip install customtkinter")
    sys.exit(1)

try:
    import pyautogui, pyperclip, pygetwindow as gw
    from pywinauto import Application
    from pywinauto.keyboard import send_keys
except ImportError:
    pass

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════
# 配色方案
# ═══════════════════════════════════════════════════════════════════
COLORS = {
    "bg": "#F0F4F8",
    "surface": "#FFFFFF",
    "card": "#E8EEF4",
    "input": "#F8FAFC",
    "border": "#CBD5E1",
    "primary": "#4F46E5",
    "primary_h": "#4338CA",
    "success": "#059669",
    "success_h": "#047857",
    "warning": "#D97706",
    "danger": "#DC2626",
    "danger_h": "#B91C1C",
    "text": "#1E293B",
    "text_dim": "#64748B",
    "accent": "#7C3AED",
}

FONTS = {
    "title": ("Microsoft YaHei", 28, "bold"),
    "heading": ("Microsoft YaHei", 15, "bold"),
    "label": ("Microsoft YaHei", 13),
    "entry": ("Consolas", 12),
    "button": ("Microsoft YaHei", 14, "bold"),
    "log": ("Cascadia Code", 11),
    "small": ("Microsoft YaHei", 12),
    "tiny": ("Microsoft YaHei", 11),
    "tab": ("Microsoft YaHei", 16, "bold"),
}


def card(parent, title=""):
    f = ctk.CTkFrame(parent, fg_color=COLORS["surface"],
                     border_color=COLORS["border"], border_width=1, corner_radius=12)
    f.pack(fill="x", pady=(0, 10))
    if title:
        ctk.CTkLabel(f, text=title, font=FONTS["heading"],
                     text_color=COLORS["text"]).pack(anchor="w", padx=16, pady=(14, 6))
    return f


class LogBox(ctk.CTkTextbox):
    def __init__(self, parent, height=400, **kw):
        super().__init__(parent,
                         font=FONTS["log"],
                         fg_color=COLORS["surface"],
                         text_color=COLORS["text_dim"],
                         border_color=COLORS["border"],
                         border_width=1,
                         corner_radius=10,
                         scrollbar_button_color=COLORS["card"],
                         scrollbar_button_hover_color=COLORS["border"],
                         height=height, **kw)
        self.tag_config("INFO", foreground="#6366F1")
        self.tag_config("OK", foreground=COLORS["success"])
        self.tag_config("WARN", foreground=COLORS["warning"])
        self.tag_config("ERR", foreground=COLORS["danger"])

    def add(self, msg, tag="INFO"):
        self.configure(state="normal")
        self.insert("end", msg + "\n", tag)
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("0.0", "end")
        self.configure(state="disabled")


class Btn(ctk.CTkButton):
    STYLES = {
        "primary": (COLORS["primary"], COLORS["primary_h"]),
        "success": (COLORS["success"], COLORS["success_h"]),
        "danger": (COLORS["danger"], COLORS["danger_h"]),
    }
    def __init__(self, parent, style="primary", text="", width=0, **kw):
        fg, fg_h = self.STYLES.get(style, self.STYLES["primary"])
        super().__init__(parent, text=text, font=FONTS["button"],
                         fg_color=fg, hover_color=fg_h,
                         text_color="white", height=44, corner_radius=10,
                         border_width=0, width=width, **kw)


def entry_with_label(parent, label_text, placeholder, browe_cmd=None):
    """输入框+标签+浏览按钮"""
    f = ctk.CTkFrame(parent, fg_color="transparent")
    f.pack(fill="x", pady=3)
    ctk.CTkLabel(f, text=label_text, font=FONTS["small"],
                 text_color=COLORS["text_dim"]).pack(anchor="w")
    h = ctk.CTkFrame(f, fg_color="transparent")
    h.pack(fill="x", pady=(4, 0))
    e = ctk.CTkEntry(h, font=FONTS["entry"], fg_color=COLORS["input"],
                      border_color=COLORS["border"], placeholder_text=placeholder, height=40)
    e.pack(side="left", fill="x", expand=True, padx=(0, 8))
    if browe_cmd:
        ctk.CTkButton(h, text="浏览", width=72, height=40,
                      fg_color=COLORS["card"], hover_color=COLORS["border"],
                      font=FONTS["small"], command=browe_cmd).pack(side="right")
    return e


# ═══════════════════════════════════════════════════════════════════
# 核心函数
# ═══════════════════════════════════════════════════════════════════
def count_dat_lines(fp):
    for enc in ["utf-8", "gbk"]:
        try:
            with open(fp, "r", encoding=enc) as f:
                return len([l for l in f if l.strip()])
        except: pass
    return 0

def get_dat_files(folder):
    return [(os.path.splitext(fn)[0], os.path.join(folder, fn))
            for fn in os.listdir(folder) if fn.upper().endswith(".DAT")]

def match_sample(name, flist):
    su = str(name).upper().strip()
    for fn, fp in flist:
        if fn.upper() == su: return (fn, fp)
    for fn, fp in flist:
        if fn.upper().startswith(su): return (fn, fp)
    for fn, fp in flist:
        if su.startswith(fn.upper()): return (fn, fp)
    return None

def determine_batch(n):
    if n <= 10: return "一"
    elif n <= 20: return "二"
    elif n <= 30: return "三"
    else: return "四"

def normalize_batch(v):
    m = {"一":"一","二":"二","三":"三","四":"四",1:"一",2:"二",3:"三",4:"四","1":"一","2":"二","3":"三","4":"四"}
    return m.get(v, "一")


# ═══════════════════════════════════════════════════════════════════
# 主程序
# ═══════════════════════════════════════════════════════════════════
class LJTApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("LJT Automation Suite")
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.processing = False
        self.stop_requested = False
        self.t3_data = None
        self.t3_files = []
        self._queue = []
        self.configure(fg_color=COLORS["bg"])

        # 标题栏
        hf = ctk.CTkFrame(self, fg_color=COLORS["surface"], height=80, corner_radius=0)
        hf.pack(fill="x")
        hf.pack_propagate(False)
        hb = ctk.CTkFrame(hf, fg_color="transparent")
        hb.pack(side="left", padx=28, pady=0)
        ctk.CTkLabel(hb, text="LJT", font=("Segoe UI", 30, "bold"), text_color=COLORS["primary"]).pack(side="left")
        ctk.CTkLabel(hb, text="Automation Suite", font=("Segoe UI", 18), text_color=COLORS["text_dim"]).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(hf, text="v1.5.0", font=("Consolas", 12), text_color=COLORS["border"]).pack(side="right", padx=24)

        # 标签页 - 放大醒目
        self.tabs = ctk.CTkTabview(self,
                                    fg_color=COLORS["bg"],
                                    segmented_button_fg_color=COLORS["card"],
                                    segmented_button_selected_color=COLORS["primary"],
                                    segmented_button_selected_hover_color=COLORS["primary_h"],
                                    segmented_button_unselected_color=COLORS["surface"],
                                    segmented_button_unselected_hover_color=COLORS["border"],
                                    text_color=COLORS["text_dim"],
                                    height=80)
        self.tabs.pack(fill="both", expand=True, padx=16, pady=(10, 12))

        t1 = self.tabs.add("📂 2G 数据导出")
        t2 = self.tabs.add("🔬 超导数据录入")
        t3 = self.tabs.add("💾 DOS 批量录入")

        self.init_tab1(t1)
        self.init_tab2(t2)
        self.init_tab3(t3)

        # 状态栏
        sf = ctk.CTkFrame(self, fg_color=COLORS["surface"], height=32, corner_radius=0)
        sf.pack(fill="x", side="bottom")
        sf.pack_propagate(False)
        ctk.CTkLabel(sf, text="LJT Automation Suite v1.5.0  |  Copyright 2026 LJT",
                     font=FONTS["tiny"], text_color=COLORS["border"]).pack(side="left", padx=16)
        self._status = ctk.CTkLabel(sf, text="就绪", font=FONTS["small"], text_color=COLORS["success"])
        self._status.pack(side="right", padx=16)

        self.after(100, self._pump)

    def _pump(self):
        while self._queue:
            tab, msg, tag = self._queue.pop(0)
            w = {"tab1": self.t1_log, "tab2": self.t2_log, "tab3": self.t3_log}.get(tab)
            if w: w.add(msg, tag)
        self.after(100, self._pump)

    def log(self, tab, msg, tag="INFO"):
        self._queue.append((tab, msg, tag))

    def status(self, txt, color=None):
        self._status.configure(text=txt, text_color=color or COLORS["success"])
        self.update_idletasks()

    def stop(self):
        self.stop_requested = True
        self.processing = False

    # ─────────────────────────────────────────────────────────────
    # Tab 1: 2G 数据导出
    # ─────────────────────────────────────────────────────────────
    def init_tab1(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", width=480)
        left.pack(side="left", fill="y", padx=(0, 14), pady=0)
        right = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=12)
        right.pack(side="right", fill="both", expand=True)

        # PAcquire
        c1 = card(left, "PAcquire 程序")
        h = ctk.CTkFrame(c1, fg_color="transparent")
        h.pack(fill="x", padx=16, pady=(0, 14))
        self.t1_exe = ctk.CTkEntry(h, font=FONTS["entry"], fg_color=COLORS["input"],
                                     border_color=COLORS["border"],
                                     placeholder_text="C:\\Program Files\\PAcquire\\PAcquire.exe", height=40)
        self.t1_exe.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(h, text="浏览", width=80, height=40,
                      fg_color=COLORS["card"], hover_color=COLORS["border"],
                      font=FONTS["small"], command=lambda: self._browse(self.t1_exe, [("EXE","*.exe")])).pack(side="right")

        # 目录
        c2 = card(left, "目录设置")
        ctk.CTkLabel(c2, text="输入目录（DAT文件）", font=FONTS["small"], text_color=COLORS["text_dim"]).pack(anchor="w", padx=16, pady=(0, 4))
        self.t1_in = ctk.CTkEntry(c2, font=FONTS["entry"], fg_color=COLORS["input"],
                                    border_color=COLORS["border"],
                                    placeholder_text="D:\\Data\\Input", height=40)
        self.t1_in.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(c2, text="输出目录", font=FONTS["small"], text_color=COLORS["text_dim"]).pack(anchor="w", padx=16, pady=(0, 4))
        self.t1_out = ctk.CTkEntry(c2, font=FONTS["entry"], fg_color=COLORS["input"],
                                     border_color=COLORS["border"],
                                     placeholder_text="D:\\Data\\Output", height=40)
        self.t1_out.pack(fill="x", padx=16, pady=(0, 14))

        # 提示
        c3 = card(left, "注意事项")
        for tip in ["• 先在 PAcquire 中配置 ASCII 导出格式",
                    "• 使用英文键盘布局",
                    "• 运行前关闭 CapsLock"]:
            ctk.CTkLabel(c3, text=tip, font=FONTS["small"], text_color=COLORS["warning"], anchor="w").pack(anchor="w", padx=16, pady=3)

        # 按钮
        c4 = card(left)
        bf = ctk.CTkFrame(c4, fg_color="transparent")
        bf.pack(fill="x", padx=16, pady=(0, 8))
        self.t1_start = Btn(bf, "success", text="▶ 开始导出", command=self.start_t1)
        self.t1_start.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.t1_stop = Btn(bf, "danger", text="■ 停止", state="disabled", command=self.stop)
        self.t1_stop.pack(side="left", fill="x", expand=True)
        self.t1_prog = ctk.CTkProgressBar(c4, height=10, corner_radius=5, progress_color=COLORS["success"], fg_color=COLORS["card"])
        self.t1_prog.pack(fill="x", padx=16, pady=(0, 0))

        self.t1_log = LogBox(right, height=500)
        self.t1_log.pack(fill="both", expand=True, padx=12, pady=12)

    def _browse(self, entry, ft):
        p = filedialog.askopenfilename(filetypes=ft)
        if p: entry.delete(0, "end"); entry.insert(0, p)

    def start_t1(self):
        exe = self.t1_exe.get().strip()
        inp = self.t1_in.get().strip()
        out = self.t1_out.get().strip()
        if not os.path.isfile(exe): messagebox.showerror("错误", "PAcquire 路径无效"); return
        if not os.path.isdir(inp): messagebox.showerror("错误", "输入目录无效"); return
        if not os.path.isdir(out): messagebox.showerror("错误", "输出目录无效"); return
        self.processing = True; self.stop_requested = False
        self.t1_start.configure(state="disabled"); self.t1_stop.configure(state="normal")
        self.status("处理中...", COLORS["warning"])
        self.t1_log.clear()
        threading.Thread(target=self.run_t1, args=(exe, inp, out), daemon=True).start()

    def run_t1(self, exe, inp, out):
        try:
            dats = [f for f in os.listdir(inp) if f.upper().endswith(".DAT")]
            if not dats: self.log("tab1", "未找到 DAT 文件!", "ERR"); return
            self.log("tab1", f"找到 {len(dats)} 个文件")
            proc = subprocess.Popen([exe]); time.sleep(8)
            def paste(t): pyperclip.copy(t); time.sleep(0.15); pyautogui.hotkey("ctrl", "v"); time.sleep(0.3)
            norm = lambda p: f'"{p}"' if " " in p else p
            for i, fn in enumerate(dats):
                if self.stop_requested: break
                self.after(0, lambda v=(i+1)/len(dats): self.t1_prog.set(v))
                self.log("tab1", f"[{i+1}/{len(dats)}] {fn}")
                try:
                    ws = pyautogui.getWindowsWithTitle("2G Enterprises Data Acquisition")
                    if ws: ws[0].activate(); time.sleep(2)
                    pyautogui.hotkey("alt", "f"); time.sleep(0.8)
                    pyautogui.press("down"); time.sleep(0.15); pyautogui.press("enter"); time.sleep(1)
                    paste(norm(os.path.join(inp, fn))); pyautogui.press("enter"); time.sleep(2)
                    pyautogui.hotkey("alt", "f"); time.sleep(0.8)
                    for _ in range(5): pyautogui.press("down"); time.sleep(0.15)
                    pyautogui.press("enter"); time.sleep(1)
                    paste(norm(out)); pyautogui.press("enter"); time.sleep(0.5)
                    paste(fn.upper()); pyautogui.press("enter"); time.sleep(1)
                    self.log("tab1", f"完成: {fn}", "OK")
                except Exception as e: self.log("tab1", f"错误: {fn} - {e}", "ERR")
            self.log("tab1", "全部完成!", "OK")
        except Exception as e: self.log("tab1", f"严重错误: {e}", "ERR")
        finally:
            if proc: proc.terminate()
            self.processing = False
            self.after(0, lambda: (self.t1_start.configure(state="normal"), self.t1_stop.configure(state="disabled")))
            self.after(0, lambda: self.status("就绪"))

    # ─────────────────────────────────────────────────────────────
    # Tab 2: 超导数据录入
    # ─────────────────────────────────────────────────────────────
    def init_tab2(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", width=480)
        left.pack(side="left", fill="y", padx=(0, 14), pady=0)
        right = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=12)
        right.pack(side="right", fill="both", expand=True)

        # 重要提醒
        w = card(left, "⚠ 重要提醒")
        ctk.CTkLabel(w, text="操作前请先在目标输出目录创建 holder 文件，用于锚定文件保存位置。",
                     font=FONTS["small"], text_color=COLORS["danger"], justify="left").pack(anchor="w", padx=16, pady=(0, 14))

        # PAcquire
        c1 = card(left, "PAcquire 程序")
        h = ctk.CTkFrame(c1, fg_color="transparent")
        h.pack(fill="x", padx=16, pady=(0, 14))
        self.t2_exe = ctk.CTkEntry(h, font=FONTS["entry"], fg_color=COLORS["input"],
                                    border_color=COLORS["border"],
                                    placeholder_text="C:\\Program Files\\PAcquire\\PAcquire.exe", height=40)
        self.t2_exe.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(h, text="浏览", width=80, height=40, fg_color=COLORS["card"], hover_color=COLORS["border"],
                      font=FONTS["small"], command=lambda: self._browse(self.t2_exe, [("EXE","*.exe")])).pack(side="right")

        # Excel
        c2 = card(left, "Excel 数据文件")
        h = ctk.CTkFrame(c2, fg_color="transparent")
        h.pack(fill="x", padx=16, pady=(0, 14))
        self.t2_excel = ctk.CTkEntry(h, font=FONTS["entry"], fg_color=COLORS["input"],
                                      border_color=COLORS["border"],
                                      placeholder_text="D:\\Data\\samples.xlsx", height=40)
        self.t2_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(h, text="浏览", width=80, height=40, fg_color=COLORS["card"], hover_color=COLORS["border"],
                      font=FONTS["small"], command=lambda: self._browse(self.t2_excel, [("Excel","*.xlsx *.xls")])).pack(side="right")

        # 格式
        c3 = card(left, "数据格式")
        ctk.CTkLabel(c3, text="列名：样品名 | 反倾向 | 倾角余角 | 地层倾向 | 地层倾角",
                     font=FONTS["small"], text_color=COLORS["accent"]).pack(anchor="w", padx=16, pady=(0, 14))

        # 按钮
        c4 = card(left)
        bf = ctk.CTkFrame(c4, fg_color="transparent")
        bf.pack(fill="x", padx=16, pady=(0, 8))
        self.t2_start = Btn(bf, "success", text="▶ 开始录入", command=self.start_t2)
        self.t2_start.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.t2_stop = Btn(bf, "danger", text="■ 停止", state="disabled", command=self.stop)
        self.t2_stop.pack(side="left", fill="x", expand=True)
        self.t2_prog = ctk.CTkProgressBar(c4, height=10, corner_radius=5, progress_color=COLORS["success"], fg_color=COLORS["card"])
        self.t2_prog.pack(fill="x", padx=16, pady=(0, 8))
        self.t2_stat = ctk.CTkLabel(c4, text="就绪", font=FONTS["small"], text_color=COLORS["text_dim"])
        self.t2_stat.pack(anchor="w", padx=16)

        self.t2_log = LogBox(right, height=500)
        self.t2_log.pack(fill="both", expand=True, padx=12, pady=12)

    def start_t2(self):
        exe = self.t2_exe.get().strip()
        excel = self.t2_excel.get().strip()
        if not os.path.exists(exe): messagebox.showerror("错误", "PAcquire 路径无效"); return
        if not os.path.exists(excel): messagebox.showerror("错误", "Excel 文件无效"); return
        self.processing = True; self.stop_requested = False
        self.t2_start.configure(state="disabled"); self.t2_stop.configure(state="normal")
        self.status("处理中...", COLORS["warning"])
        self.t2_log.clear()
        threading.Thread(target=self.run_t2, args=(exe, excel), daemon=True).start()

    def run_t2(self, exe, excel):
        try:
            df = pd.read_excel(excel)
            for cn, en in [("样品名","SampleName"),("反倾向","AntiTrend"),("倾角余角","IncAngle"),("地层倾向","Strike"),("地层倾角","Dip")]:
                if cn in df.columns and en not in df.columns:
                    df = df.rename(columns={cn: en})
            cols = ["SampleName", "AntiTrend", "IncAngle", "Strike", "Dip"]
            miss = [c for c in cols if c not in df.columns]
            if miss: self.log("tab2", f"缺少列: {miss}", "ERR"); return
            total = len(df); self.log("tab2", f"加载 {total} 条记录")
            try: app = Application(backend="win32").start(exe)
            except: app = Application(backend="uia").start(exe)
            time.sleep(8)
            win = None
            for t in ["2G Enterprises Data Acquisition", "Data Acquisition"]:
                try: win = app.window(title=t); win.wait("exists visible", timeout=15); break
                except: pass
            if not win: self.log("tab2", "无法连接 PAcquire", "ERR"); return
            ok = 0
            for i, row in df.iterrows():
                if self.stop_requested: break
                self.after(0, lambda v=int((i+1)/total*100): (self.t2_prog.set(v/100), self.t2_stat.configure(text=f"{i+1}/{total}")))
                try:
                    win.menu_select("File->New..."); time.sleep(1.5)
                    sw = app.window(title="Sample Information"); sw.wait("exists visible", timeout=15)
                    sw.type_keys("%n"); time.sleep(0.3); send_keys(str(row["SampleName"]).upper())
                    sw.type_keys("%d"); time.sleep(0.3); send_keys(str(row["AntiTrend"]))
                    sw.type_keys("%p"); time.sleep(0.3); send_keys(str(row["IncAngle"]))
                    sw.type_keys("%a"); time.sleep(0.3); send_keys(str(row["Strike"]))
                    sw.type_keys("%u"); time.sleep(0.3); send_keys(str(row["Dip"]))
                    sw.type_keys("%o"); time.sleep(1.5)
                    ok += 1
                    self.log("tab2", f"完成: {row['SampleName']}", "OK")
                except Exception as e: self.log("tab2", f"错误: {row.get('SampleName','?')}", "ERR")
            self.after(0, lambda: (self.t2_prog.set(1), self.t2_stat.configure(text=f"完成: {ok}/{total}")))
        except Exception as e: self.log("tab2", f"严重错误: {e}", "ERR")
        finally:
            self.processing = False
            self.after(0, lambda: (self.t2_start.configure(state="normal"), self.t2_stop.configure(state="disabled")))
            self.after(0, lambda: self.status("就绪"))

    # ─────────────────────────────────────────────────────────────
    # Tab 3: DOS 批量录入
    # ─────────────────────────────────────────────────────────────
    def init_tab3(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", width=480)
        left.pack(side="left", fill="y", padx=(0, 14), pady=0)
        right = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=12)
        right.pack(side="right", fill="both", expand=True)

        # 文件选择（精简版，无DAT浏览按钮）
        c1 = card(left, "文件选择")
        ctk.CTkLabel(c1, text="Excel 文件（含样品数据）", font=FONTS["small"], text_color=COLORS["text_dim"]).pack(anchor="w", padx=16, pady=(0, 4))
        h = ctk.CTkFrame(c1, fg_color="transparent")
        h.pack(fill="x", padx=16, pady=(4, 10))
        self.t3_excel = ctk.CTkEntry(h, font=FONTS["entry"], fg_color=COLORS["input"],
                                      border_color=COLORS["border"],
                                      placeholder_text="D:\\Data\\samples.xlsx", height=40)
        self.t3_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(h, text="浏览", width=80, height=40, fg_color=COLORS["card"], hover_color=COLORS["border"],
                      font=FONTS["small"], command=lambda: self._browse(self.t3_excel, [("Excel","*.xlsx *.xls")])).pack(side="right")

        ctk.CTkLabel(c1, text="DAT 文件夹路径（手动输入）", font=FONTS["small"], text_color=COLORS["text_dim"]).pack(anchor="w", padx=16, pady=(0, 4))
        self.t3_folder = ctk.CTkEntry(c1, font=FONTS["entry"], fg_color=COLORS["input"],
                                       border_color=COLORS["border"],
                                       placeholder_text="D:\\Data\\DAT", height=40)
        self.t3_folder.pack(fill="x", padx=16, pady=(4, 14))

        # 操作指南（合并）
        c2 = card(left, "使用指南")
        ctk.CTkLabel(c2, text="1.选Excel 2.输DAT路径 3.扫描 4.预处理 5.开DOSBox 6.开始",
                     font=FONTS["small"], text_color=COLORS["text"], anchor="w", wraplength=420).pack(anchor="w", padx=16, pady=(0, 6))
        ctk.CTkLabel(c2, text="⚠ DOS需先到Edit界面 | 运行中禁止操作键盘鼠标",
                     font=FONTS["small"], text_color=COLORS["warning"], anchor="w", wraplength=420).pack(anchor="w", padx=16, pady=(0, 14))

        # 按钮行
        c5 = card(left)
        bf = ctk.CTkFrame(c5, fg_color="transparent")
        bf.pack(fill="x", padx=16, pady=(0, 6))
        self.t3_scan = Btn(bf, "primary", text="① 扫描", width=100, command=self.scan_t3)
        self.t3_scan.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.t3_pre = Btn(bf, "primary", text="② 预处理", width=100, command=self.preprocess_t3)
        self.t3_pre.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.t3_start = Btn(bf, "success", text="③ 开始", width=100, state="disabled", command=self.start_t3)
        self.t3_start.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.t3_stop = Btn(bf, "danger", text="■", width=50, state="disabled", command=self.stop)
        self.t3_stop.pack(side="left")

        self.t3_prog = ctk.CTkProgressBar(c5, height=10, corner_radius=5, progress_color=COLORS["success"], fg_color=COLORS["card"])
        self.t3_prog.pack(fill="x", padx=16, pady=(0, 6))
        self.t3_stat = ctk.CTkLabel(c5, text="未加载数据", font=FONTS["small"], text_color=COLORS["text_dim"])
        self.t3_stat.pack(anchor="w", padx=16)

        self.t3_log = LogBox(right, height=520)
        self.t3_log.pack(fill="both", expand=True, padx=12, pady=12)

    def scan_t3(self):
        folder = self.t3_folder.get().strip()
        if not os.path.isdir(folder): messagebox.showerror("错误", "DAT 文件夹路径无效"); return
        self.t3_files = get_dat_files(folder)
        self.t3_log.clear()
        self.t3_log.add(f"找到 {len(self.t3_files)} 个 DAT 文件")
        for name, path in self.t3_files[:10]:
            lines = count_dat_lines(path)
            batch = determine_batch(lines)
            self.t3_log.add(f"  {name}: {lines}行 -> 批次{batch}")
        if len(self.t3_files) > 10: self.t3_log.add(f"  ... 还有 {len(self.t3_files)-10} 个")

    def preprocess_t3(self):
        excel = self.t3_excel.get().strip()
        folder = self.t3_folder.get().strip()
        if not os.path.isfile(excel): messagebox.showerror("错误", "请选择 Excel 文件"); return
        if not os.path.isdir(folder): messagebox.showerror("错误", "DAT 文件夹路径无效"); return
        if not self.t3_files: self.t3_files = get_dat_files(folder)
        try:
            df = pd.read_excel(excel)
            self.t3_log.clear()
            self.t3_log.add(f"列名: {list(df.columns)}")
            if "SampleName" not in df.columns and "样品名" not in df.columns:
                messagebox.showerror("错误", "需要「样品名」列"); return
            for cn, en in [("样品名","SampleName"),("反倾向","A"),("倾角余角","B"),("地层倾向","S"),("地层倾角","D")]:
                if cn in df.columns and en not in df.columns:
                    df = df.rename(columns={cn: en})
            if "Batch" not in df.columns and "批次" not in df.columns:
                df["Batch"] = "一"
            elif "批次" in df.columns and "Batch" not in df.columns:
                df = df.rename(columns={"批次": "Batch"})
            matched, no_dat = 0, []
            for idx, row in df.iterrows():
                sname = row["SampleName"] if "SampleName" in row.index else ""
                if pd.isna(sname) or not str(sname).strip(): continue
                m = match_sample(str(sname), self.t3_files)
                if m:
                    _, fp = m
                    df.at[idx, "Batch"] = determine_batch(count_dat_lines(fp))
                    matched += 1
                else:
                    no_dat.append(str(sname))
            if no_dat:
                self.t3_log.add(f"删除 {len(no_dat)} 个无 DAT 文件的样品:", "WARN")
                for s in no_dat[:5]: self.t3_log.add(f"  - {s}", "WARN")
                df = df[~df["SampleName"].astype(str).isin(no_dat)].reset_index(drop=True)
            self.t3_data = df
            self.t3_log.add(f"预处理完成: {len(df)} 个样品")
            self.t3_log.add(f"匹配成功: {matched} 个")
            bc = df["Batch"].value_counts()
            for b, lbl in [("一","1-10"),("二","11-20"),("三","21-30"),("四","30+")]:
                cnt = bc.get(b, 0)
                if cnt: self.t3_log.add(f"  批次{b}({lbl}行): {cnt}个")
            self.t3_log.add("数据预览(前3条):")
            for _, r in df.head(3).iterrows():
                sn = str(r.get("SampleName","")).strip()
                sb = str(r.get("Batch","")).strip()
                sa = str(r.get("A","")).strip() if pd.notna(r.get("A")) else ""
                sb2 = str(r.get("B","")).strip() if pd.notna(r.get("B")) else ""
                ss = str(r.get("S","")).strip() if pd.notna(r.get("S")) else ""
                sd = str(r.get("D","")).strip() if pd.notna(r.get("D")) else ""
                self.t3_log.add(f"  {sn} | B={sb} | A={sa} B={sb2} S={ss} D={sd}")
            df.to_excel(excel.replace(".xlsx", "_processed.xlsx"), index=False)
            self.t3_stat.configure(text=f"共{len(df)}个 | 匹配{matched}个")
            self.t3_start.configure(state="normal")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.t3_log.add(f"错误: {e}", "ERR")

    def start_t3(self):
        if self.t3_data is None:
            messagebox.showerror("错误", "请先运行预处理"); return
        self.processing = True; self.stop_requested = False
        self.t3_start.configure(state="disabled"); self.t3_stop.configure(state="normal")
        self.status("处理中...", COLORS["warning"])
        self.t3_log.clear()
        threading.Thread(target=self.run_t3, daemon=True).start()

    def run_t3(self):
        try:
            total = len(self.t3_data)
            self.log("tab3", "开始 DOS 录入...")
            self.log("tab3", "请确保 DOSBox 在 Edit 界面!", "WARN")
            time.sleep(2)
            try:
                ws = gw.getWindowsWithTitle("DOSBox")
                if ws: ws[0].activate(); time.sleep(1); self.log("tab3", "DOSBox 已激活")
            except: pass
            delay = 1.0
            for i, (idx, row) in enumerate(self.t3_data.iterrows()):
                if self.stop_requested: self.log("tab3", "用户停止", "WARN"); break
                self.after(0, lambda v=(i+1)/total: self.t3_prog.set(v))
                try:
                    sn = str(row["SampleName"]).strip() if "SampleName" in row.index else "?"
                    sb = normalize_batch(str(row["Batch"]).strip()) if "Batch" in row.index else "一"
                    sa = str(row["A"]).strip() if "A" in row.index and pd.notna(row["A"]) else ""
                    sb2 = str(row["B"]).strip() if "B" in row.index and pd.notna(row["B"]) else ""
                    ss = str(row["S"]).strip() if "S" in row.index and pd.notna(row["S"]) else ""
                    sd = str(row["D"]).strip() if "D" in row.index and pd.notna(row["D"]) else ""
                except: sn, sb, sa, sb2, ss, sd = "?", "一", "", "", "", ""
                extra = {"一":2,"二":3,"三":4,"四":5}.get(sb, 2)
                self.log("tab3", f"[{i+1}/{total}] {sn}(B{sb}) A={sa} B={sb2} S={ss} D={sd} ex={extra}")
                try:
                    pyautogui.press("e"); time.sleep(delay)
                    pyautogui.press("y"); time.sleep(delay)
                    pyautogui.press("a"); time.sleep(delay)
                    if sa: pyautogui.write(sa, interval=0.1)
                    time.sleep(delay); pyautogui.press("enter"); time.sleep(delay)
                    pyautogui.press("b"); time.sleep(delay)
                    if sb2: pyautogui.write(sb2, interval=0.1)
                    time.sleep(delay); pyautogui.press("enter"); time.sleep(delay)
                    pyautogui.press("s"); time.sleep(delay)
                    if ss: pyautogui.write(ss, interval=0.1)
                    time.sleep(delay); pyautogui.press("enter"); time.sleep(delay)
                    pyautogui.press("d"); time.sleep(delay)
                    if sd: pyautogui.write(sd, interval=0.1)
                    time.sleep(delay); pyautogui.press("enter"); time.sleep(delay)
                    for _ in range(extra):
                        pyautogui.press("enter"); time.sleep(delay)
                    time.sleep(delay)
                    pyautogui.press("y"); time.sleep(delay)
                    pyautogui.press("y"); time.sleep(delay)
                    pyautogui.press("r"); time.sleep(delay)
                    pyautogui.press("n"); time.sleep(delay)
                    pyautogui.press("enter"); time.sleep(delay*2)
                    self.log("tab3", f"  OK: {sn}", "OK")
                except Exception as e: self.log("tab3", f"  错误: {sn} - {e}", "ERR")
            self.log("tab3", "全部完成!", "OK")
        except Exception as e: self.log("tab3", f"严重错误: {e}", "ERR")
        finally:
            self.processing = False
            self.after(0, lambda: (self.t3_start.configure(state="normal"), self.t3_stop.configure(state="disabled")))
            self.after(0, lambda: self.status("就绪"))


if __name__ == "__main__":
    app = LJTApp()
    app.mainloop()
