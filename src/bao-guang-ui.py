"""
bao-guang-ui.py — 播放量抓取工具图形界面
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import importlib.util
import sys
import os
import io
import re
import queue
import builtins
import json
from datetime import datetime, timedelta

# ── 动态加载 bao-guang.py（文件名含连字符，无法直接 import）──────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_CFG_PATH = os.path.join(_HERE, ".ui-config.json")


def _load_config() -> dict:
    try:
        with open(_CFG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(cfg: dict):
    try:
        with open(_CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
_spec = importlib.util.spec_from_file_location(
    "baoguang", os.path.join(_HERE, "bao-guang.py"))
_bg = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_bg)

# ── 颜色 & 字体 ──────────────────────────────────────────────────────
_BG     = "#F0F2F7"
_CARD   = "#FFFFFF"
_ACCENT = "#4F6EF7"
_TEXT   = "#1A202C"
_SUB    = "#718096"
_OK     = "#38A169"
_ERR    = "#E53E3E"
_WARN   = "#D69E2E"
_BORDER = "#E2E8F0"
_LOG_BG = "#1A202C"
_LOG_FG = "#A0AEC0"

_FN = "Microsoft YaHei UI"
F_H  = (_FN, 17, "bold")
F_S  = (_FN, 10, "bold")
F_N  = (_FN, 10)
F_SM = (_FN, 9)
F_M  = ("Consolas", 9)
F_B  = (_FN, 10, "bold")


# ── 工具函数 ──────────────────────────────────────────────────────────

def _make_card(parent: tk.Widget, title: str = "") -> tk.Frame:
    outer = tk.Frame(parent, bg=_CARD,
                     highlightthickness=1, highlightbackground=_BORDER)
    if title:
        tk.Label(outer, text=title, bg=_CARD, fg=_TEXT,
                 font=F_S).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Frame(outer, bg=_BORDER, height=1).pack(fill="x", pady=(8, 0))
    return outer


def _make_btn(parent, text, cmd, bg, fg="#FFFFFF",
              width=None, state="normal"):
    kw = dict(text=text, command=cmd, bg=bg, fg=fg, font=F_B, relief="flat",
              cursor="hand2", state=state, activebackground=bg, activeforeground=fg,
              padx=14, pady=7)
    if width is not None:
        kw["width"] = width
    return tk.Button(parent, **kw)


class _Capture(io.StringIO):
    """将 stdout 的行输出转发到回调函数"""
    def __init__(self, cb):
        super().__init__()
        self._cb = cb
        self._buf = ""

    def write(self, s):
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line:
                self._cb(line)

    def flush(self):
        if self._buf.strip():
            self._cb(self._buf)
            self._buf = ""


# ── Cron 选择器对话框 ────────────────────────────────────────────────

class _CronDialog(tk.Toplevel):
    """可视化 Cron 表达式构建器（分 时 日 月 周）"""

    _MONTH_NAMES = ["1月","2月","3月","4月","5月","6月",
                    "7月","8月","9月","10月","11月","12月"]
    _DOW_NAMES   = ["周日","周一","周二","周三","周四","周五","周六"]

    _FIELD_CFG = [
        {"label": "分钟", "min": 0,  "max": 59, "step_max": 30, "cols": 10},
        {"label": "小时", "min": 0,  "max": 23, "step_max": 12, "cols": 8},
        {"label": "日期", "min": 1,  "max": 31, "step_max": 15, "cols": 8},
        {"label": "月份", "min": 1,  "max": 12, "step_max": 6,  "cols": 4},
        {"label": "星期", "min": 0,  "max": 6,  "step_max": 3,  "cols": 7},
    ]

    def __init__(self, parent, current_expr: str):
        super().__init__(parent)
        self.title("Cron 选择器")
        self.configure(bg=_BG)
        self.resizable(True, False)
        self.grab_set()
        self.transient(parent)

        self.result: str | None = None
        self._fields = self._parse_expr(current_expr)
        self._preview_var = tk.StringVar()

        self._build()
        self._update_preview()

        self.update_idletasks()
        # 居中于父窗口
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h   = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w)//2}+{py + (ph - h)//2}")
        self.focus_set()

    # ── 解析表达式 ──────────────────────────────────────────────────

    def _parse_expr(self, expr: str) -> list[dict]:
        parts = (expr.strip().split() + ["*"] * 5)[:5]
        fields = []
        for part, cfg in zip(parts, self._FIELD_CFG):
            d: dict = dict(cfg)
            d["mode"]       = tk.StringVar(value="any")
            d["step"]       = tk.IntVar(value=1)
            d["range_from"] = tk.IntVar(value=cfg["min"])
            d["range_to"]   = tk.IntVar(value=cfg["max"])
            d["checks"]     = {}   # int → BooleanVar（构建 UI 时填充）
            d["pre_sel"]    = set()

            if part == "*":
                d["mode"].set("any")
            elif part.startswith("*/"):
                try:
                    d["step"].set(int(part[2:]))
                    d["mode"].set("step")
                except ValueError:
                    pass
            elif re.fullmatch(r'\d+-\d+', part):
                a, b = part.split("-")
                d["range_from"].set(int(a))
                d["range_to"].set(int(b))
                d["mode"].set("range")
            else:
                sel: set[int] = set()
                for seg in part.split(","):
                    try:
                        sel.add(int(seg))
                    except ValueError:
                        pass
                if sel:
                    d["pre_sel"] = sel
                    d["mode"].set("specific")

            fields.append(d)
        return fields

    # ── 构建对话框 ──────────────────────────────────────────────────

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=(14, 6))

        tab_titles = ["  分钟  ", "  小时  ", "  日期  ", "  月份  ", "  星期  "]
        for i, (title, field) in enumerate(zip(tab_titles, self._fields)):
            tab = tk.Frame(nb, bg=_CARD, padx=14, pady=10)
            nb.add(tab, text=title)
            self._build_tab(tab, field, i)

        # 预览栏
        pf = tk.Frame(self, bg=_BG)
        pf.pack(fill="x", padx=14, pady=(2, 6))
        tk.Label(pf, text="表达式：", bg=_BG, fg=_SUB,
                 font=F_SM).pack(side="left")
        tk.Label(pf, textvariable=self._preview_var,
                 bg=_BG, fg=_ACCENT,
                 font=("Consolas", 11, "bold")).pack(side="left")

        # 按钮行
        bf = tk.Frame(self, bg=_BG)
        bf.pack(fill="x", padx=14, pady=(0, 14))
        _make_btn(bf, "确定", self._ok,      _ACCENT, width=8).pack(side="right")
        _make_btn(bf, "取消", self.destroy,  _SUB,    width=8).pack(side="right", padx=(0, 8))

    def _build_tab(self, frame: tk.Frame, field: dict, idx: int):
        upd = self._update_preview

        def rb_kw(val):
            return dict(variable=field["mode"], value=val,
                        bg=_CARD, fg=_TEXT, font=F_N,
                        activebackground=_CARD, selectcolor=_CARD,
                        command=upd)

        # ── 任意 ────────────────────────────────────────────────────
        tk.Radiobutton(frame, text="任意（*）", **rb_kw("any")).pack(anchor="w", pady=(0, 6))

        # ── 每隔 ────────────────────────────────────────────────────
        sf = tk.Frame(frame, bg=_CARD)
        sf.pack(anchor="w", pady=(0, 6))
        tk.Radiobutton(sf, text="每隔", **rb_kw("step")).pack(side="left")
        sp_step = tk.Spinbox(sf, from_=1, to=field["step_max"],
                             textvariable=field["step"],
                             width=4, font=F_N, bg="#F7FAFC", relief="flat",
                             highlightthickness=1, highlightbackground=_BORDER,
                             justify="center", command=upd)
        sp_step.pack(side="left", padx=6, ipady=4)
        sp_step.bind("<KeyRelease>", lambda *_: upd())
        tk.Label(sf, text=field["label"], bg=_CARD, fg=_TEXT,
                 font=F_N).pack(side="left")

        # ── 指定值 ──────────────────────────────────────────────────
        spf = tk.Frame(frame, bg=_CARD)
        spf.pack(anchor="w", pady=(0, 6))
        tk.Radiobutton(spf, text="指定值", **rb_kw("specific")).pack(anchor="w")
        chk_wrap = tk.Frame(spf, bg=_CARD)
        chk_wrap.pack(anchor="w", padx=(20, 0))
        self._build_checkboxes(chk_wrap, field, idx)

        # ── 范围 ────────────────────────────────────────────────────
        rf = tk.Frame(frame, bg=_CARD)
        rf.pack(anchor="w")
        tk.Radiobutton(rf, text="范围", **rb_kw("range")).pack(side="left")
        sp_from = tk.Spinbox(rf, from_=field["min"], to=field["max"],
                             textvariable=field["range_from"],
                             width=4, font=F_N, bg="#F7FAFC", relief="flat",
                             highlightthickness=1, highlightbackground=_BORDER,
                             justify="center", command=upd)
        sp_from.pack(side="left", padx=6, ipady=4)
        sp_from.bind("<KeyRelease>", lambda *_: upd())
        tk.Label(rf, text="至", bg=_CARD, fg=_TEXT, font=F_N).pack(side="left")
        sp_to = tk.Spinbox(rf, from_=field["min"], to=field["max"],
                           textvariable=field["range_to"],
                           width=4, font=F_N, bg="#F7FAFC", relief="flat",
                           highlightthickness=1, highlightbackground=_BORDER,
                           justify="center", command=upd)
        sp_to.pack(side="left", padx=6, ipady=4)
        sp_to.bind("<KeyRelease>", lambda *_: upd())

    def _build_checkboxes(self, parent: tk.Frame, field: dict, idx: int):
        pre = field["pre_sel"]
        cols = field["cols"]

        if idx == 4:  # 星期：名称标签
            items = list(enumerate(self._DOW_NAMES))
        elif idx == 3:  # 月份：名称标签
            items = [(v, name) for v, name in enumerate(self._MONTH_NAMES, start=1)]
        else:
            items = [(v, str(v)) for v in range(field["min"], field["max"] + 1)]

        for pos, (val, label) in enumerate(items):
            bv = tk.BooleanVar(value=(val in pre))
            bv.trace_add("write", lambda *_: self._update_preview())
            field["checks"][val] = bv
            tk.Checkbutton(
                parent, text=label, variable=bv,
                bg=_CARD, fg=_TEXT, font=F_SM,
                activebackground=_CARD, selectcolor=_CARD,
                command=self._update_preview,
            ).grid(row=pos // cols, column=pos % cols, sticky="w", padx=2, pady=1)

    # ── 生成 & 预览 ────────────────────────────────────────────────

    def _field_str(self, field: dict) -> str:
        mode = field["mode"].get()
        if mode == "step":
            return f"*/{field['step'].get()}"
        if mode == "range":
            return f"{field['range_from'].get()}-{field['range_to'].get()}"
        if mode == "specific":
            sel = sorted(v for v, bv in field["checks"].items() if bv.get())
            if not sel or len(sel) == field["max"] - field["min"] + 1:
                return "*"
            return ",".join(str(v) for v in sel)
        return "*"

    def _update_preview(self, *_):
        self._preview_var.set(" ".join(self._field_str(f) for f in self._fields))

    def _ok(self):
        self.result = self._preview_var.get()
        self.destroy()


# ── 主应用 ───────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("曝光量抓取工具")
        self.geometry("1280x860")
        self.minsize(900, 760)
        self.configure(bg=_BG)

        self._running      = False
        self._stop_evt     = threading.Event()
        self._log_q        = queue.Queue()
        self._next_vars    = [tk.StringVar(value="—") for _ in range(5)]
        self._status_var   = tk.StringVar(value="就绪")
        self._cfg          = _load_config()
        self._cron_var     = tk.StringVar(value=self._cfg.get("cron", "0 8 * * *"))
        self._overlap_var  = tk.StringVar(value=self._cfg.get("overlap", "skip"))
        self._scrape_cv    = threading.Condition(threading.Lock())
        self._scrape_active = False
        self._scrape_queue  = queue.Queue()

        self._build_ui()
        self._poll_log()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._worker_thread = threading.Thread(
            target=self._worker_loop, daemon=True, name="scrape-worker")
        self._worker_thread.start()

    # ── 构建界面 ─────────────────────────────────────────────────────

    def _build_ui(self):
        # 顶部标题栏
        hdr = tk.Frame(self, bg=_ACCENT, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📊  曝光量抓取工具",
                 bg=_ACCENT, fg="white", font=F_H,
                 padx=24).pack(side="left", fill="y")

        body = tk.Frame(self, bg=_BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # 左列：配置区（宽度自适应内容）
        left = tk.Frame(body, bg=_BG)
        left.pack(side="left", fill="y", padx=(0, 14))

        # 右列：日志区（填满剩余宽度和高度）
        right = tk.Frame(body, bg=_BG)
        right.pack(side="left", fill="both", expand=True)

        # ── 卡片 1：Google 文档地址 ──────────────────────────────────
        c1 = _make_card(left, "🔗  Google 文档地址")
        c1.pack(fill="x", pady=(0, 10))
        f1 = tk.Frame(c1, bg=_CARD)
        f1.pack(fill="x", padx=16, pady=12)

        saved_url = self._cfg.get("google_sheet_url", _bg.GOOGLE_SHEET_URL or "")
        self._url_var = tk.StringVar(value=saved_url)
        url_entry = tk.Entry(
            f1, textvariable=self._url_var,
            font=F_N, bg="#F7FAFC", fg=_TEXT,
            relief="flat", insertbackground=_TEXT,
            highlightthickness=1,
            highlightbackground=_BORDER,
            highlightcolor=_ACCENT)
        url_entry.pack(fill="x", ipady=8)
        tk.Label(f1,
                 text="示例：https://docs.google.com/spreadsheets/d/…/edit",
                 bg=_CARD, fg=_SUB, font=F_SM).pack(anchor="w", pady=(4, 0))

        # ── 卡片 2：平台登录 ─────────────────────────────────────────
        c_login = _make_card(left, "🔑  平台登录")
        c_login.pack(fill="x", pady=(0, 10))
        fl = tk.Frame(c_login, bg=_CARD)
        fl.pack(fill="x", padx=16, pady=12)

        _LOGIN_PLATFORMS = [
            ("Instagram", "instagram", "#E1306C"),
            ("Facebook",  "facebook",  "#1877F2"),
            ("X (Twitter)", "x",       "#000000"),
            ("TikTok",    "tiktok",    "#010101"),
            ("Threads",   "threads",   "#333333"),
        ]
        self._login_btns: dict[str, tk.Button] = {}
        for label, key, color in _LOGIN_PLATFORMS:
            btn = tk.Button(
                fl, text=label,
                command=lambda k=key: self._open_login(k),
                bg=color, fg="white", font=F_B,
                relief="flat", cursor="hand2",
                padx=14, pady=7,
                activebackground=color, activeforeground="white")
            btn.pack(side="left", padx=(0, 8))
            self._login_btns[key] = btn

        tk.Label(fl, text="点击打开对应平台浏览器完成登录",
                 bg=_CARD, fg=_SUB, font=F_SM).pack(side="left", padx=(8, 0))

        # ── 卡片 3：定时设置 ─────────────────────────────────────────
        c2 = _make_card(left, "⏰  定时设置（Cron 表达式）")
        c2.pack(fill="x", pady=(0, 10))
        f2 = tk.Frame(c2, bg=_CARD)
        f2.pack(fill="x", padx=16, pady=14)

        # cron 输入行
        row = tk.Frame(f2, bg=_CARD)
        row.pack(anchor="w", fill="x")
        tk.Label(row, text="Cron：", bg=_CARD, fg=_TEXT,
                 font=F_N).pack(side="left")
        cron_entry = tk.Entry(
            row, textvariable=self._cron_var,
            width=22, font=("Consolas", 10),
            bg="#F7FAFC", fg=_TEXT, relief="flat",
            insertbackground=_TEXT,
            highlightthickness=1,
            highlightbackground=_BORDER,
            highlightcolor=_ACCENT)
        cron_entry.pack(side="left", padx=(4, 8), ipady=6)
        tk.Button(
            row, text="选择器",
            command=self._open_cron_selector,
            bg=_ACCENT, fg="white", font=F_SM,
            relief="flat", cursor="hand2",
            padx=10, pady=5,
            activebackground=_ACCENT, activeforeground="white",
        ).pack(side="left", padx=(0, 10))
        tk.Label(row, text="分 时 日 月 周（* 表示任意）",
                 bg=_CARD, fg=_SUB, font=F_SM).pack(side="left")

        # 快捷预设
        row_p = tk.Frame(f2, bg=_CARD)
        row_p.pack(anchor="w", pady=(8, 0))
        tk.Label(row_p, text="预设：", bg=_CARD, fg=_SUB,
                 font=F_SM).pack(side="left")
        _PRESETS = [
            ("每天 8 点",  "0 8 * * *"),
            ("每天 0 点",  "0 0 * * *"),
            ("每 2 小时",  "0 */2 * * *"),
            ("每 30 分",   "*/30 * * * *"),
            ("工作日 9 点","0 9 * * 1-5"),
        ]
        for label, expr in _PRESETS:
            tk.Button(
                row_p, text=label,
                command=lambda e=expr: self._set_cron(e),
                bg=_BORDER, fg=_TEXT, font=F_SM,
                relief="flat", cursor="hand2",
                padx=8, pady=3,
                activebackground=_ACCENT, activeforeground="white",
            ).pack(side="left", padx=(0, 4))

        # 下次运行预览（5条）
        row2 = tk.Frame(f2, bg=_CARD)
        row2.pack(anchor="w", pady=(10, 0))
        tk.Label(row2, text="计划时间：",
                 bg=_CARD, fg=_SUB, font=F_SM).grid(row=0, column=0, sticky="w",
                                                     rowspan=5, padx=(0, 6))
        _NUMS = ["①", "②", "③", "④", "⑤"]
        for i, var in enumerate(self._next_vars):
            tk.Label(row2, text=_NUMS[i], bg=_CARD, fg=_SUB,
                     font=F_SM).grid(row=i, column=1, sticky="w", padx=(0, 4))
            tk.Label(row2, textvariable=var,
                     bg=_CARD, fg=_ACCENT, font=F_SM,
                     width=20, anchor="w").grid(row=i, column=2, sticky="w")

        # 并发处理选项
        row_ov = tk.Frame(f2, bg=_CARD)
        row_ov.pack(anchor="w", pady=(8, 0))
        tk.Label(row_ov, text="并发处理：",
                 bg=_CARD, fg=_SUB, font=F_SM).pack(side="left")
        for _lbl, _val in [("跳过本次", "skip"), ("排队等待", "queue")]:
            tk.Radiobutton(
                row_ov, text=_lbl,
                variable=self._overlap_var, value=_val,
                bg=_CARD, fg=_TEXT, font=F_SM,
                activebackground=_CARD, selectcolor=_CARD,
            ).pack(side="left", padx=(0, 16))

        # 实时更新预览
        self._cron_var.trace_add("write", lambda *_: self._update_cron_preview())
        self._update_cron_preview()

        # ── 操作按钮行 ───────────────────────────────────────────────
        c_btn = _make_card(left, "")
        c_btn.pack(fill="x", pady=(0, 10))
        btn_row = tk.Frame(c_btn, bg=_CARD)
        btn_row.pack(fill="x", padx=16, pady=12)

        self._b_run = _make_btn(btn_row, "▶  立即运行",
                                self._run_once, _OK)
        self._b_run.pack(side="left", padx=(0, 8))

        self._b_start = _make_btn(btn_row, "⏰  启动定时",
                                  self._start_timer, _ACCENT)
        self._b_start.pack(side="left", padx=(0, 8))

        self._b_stop = _make_btn(btn_row, "⏹  停止",
                                 self._stop_all, _ERR,
                                 state="disabled")
        self._b_stop.pack(side="left")

        # 右侧状态
        sf = tk.Frame(btn_row, bg=_CARD)
        sf.pack(side="right")
        tk.Label(sf, text="状态：", bg=_CARD, fg=_SUB,
                 font=F_SM).pack(side="left")
        self._sl = tk.Label(sf, textvariable=self._status_var,
                            bg=_CARD, fg=_OK, font=F_B)
        self._sl.pack(side="left")

        # ── 右列：运行日志 ───────────────────────────────────────────
        c3 = _make_card(right, "📋  运行日志")
        c3.pack(fill="both", expand=True)

        log_wrap = tk.Frame(c3, bg=_LOG_BG)
        log_wrap.pack(fill="both", expand=True, padx=16, pady=(8, 0))

        self._lt = tk.Text(
            log_wrap, bg=_LOG_BG, fg=_LOG_FG,
            font=F_M, relief="flat", state="disabled",
            wrap="word", selectbackground=_ACCENT,
            padx=10, pady=8)
        sb = tk.Scrollbar(log_wrap, command=self._lt.yview,
                          bg="#2D3748", troughcolor="#2D3748",
                          relief="flat")
        self._lt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._lt.pack(fill="both", expand=True)

        self._lt.tag_config("ts",  foreground="#5C7CFA")
        self._lt.tag_config("ok",  foreground="#68D391")
        self._lt.tag_config("w",   foreground="#ECC94B")
        self._lt.tag_config("err", foreground="#FC8181")
        self._lt.tag_config("n",   foreground=_LOG_FG)

        foot = tk.Frame(c3, bg=_CARD)
        foot.pack(fill="x", padx=16, pady=8)
        lbl_clear = tk.Label(foot, text="清空日志",
                             bg=_CARD, fg=_SUB, font=F_SM, cursor="hand2")
        lbl_clear.pack(side="right")
        lbl_clear.bind("<Button-1>", lambda _: self._clear_log())

    # ── 日志 ─────────────────────────────────────────────────────────

    def log(self, msg: str):
        if not msg or not msg.strip():
            return
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_q.put((ts, msg.rstrip()))

    def _poll_log(self):
        try:
            while True:
                ts, msg = self._log_q.get_nowait()
                low = msg.lower()
                if any(k in low for k in ["✅", "完成", "成功", "已写入", "excel"]):
                    tag = "ok"
                elif any(k in low for k in ["❌", "警告", "跳过", "失败", "出错"]):
                    tag = "w"
                elif any(k in low for k in ["error", "错误", "exception"]):
                    tag = "err"
                else:
                    tag = "n"
                self._lt.config(state="normal")
                self._lt.insert("end", f"[{ts}] ", "ts")
                self._lt.insert("end", msg + "\n", tag)
                self._lt.see("end")
                self._lt.config(state="disabled")
        except queue.Empty:
            pass
        self.after(60, self._poll_log)

    def _clear_log(self):
        self._lt.config(state="normal")
        self._lt.delete("1.0", "end")
        self._lt.config(state="disabled")

    # ── 控制逻辑 ─────────────────────────────────────────────────────

    def _set_status(self, text: str, color: str):
        self._status_var.set(text)
        self._sl.config(fg=color)

    def _set_btns(self, running: bool):
        s_idle = "disabled" if running else "normal"
        s_stop = "normal"   if running else "disabled"
        self._b_run.config(state=s_idle)
        self._b_start.config(state=s_idle)
        self._b_stop.config(state=s_stop)

    # ── Cron 工具 ────────────────────────────────────────────────────

    @staticmethod
    def _cron_next(expr: str, after: datetime = None) -> datetime | None:
        """计算 cron 表达式的下一次触发时间（标准5字段：分 时 日 月 周）"""
        try:
            parts = expr.strip().split()
            if len(parts) != 5:
                return None
            m_f, h_f, dom_f, mon_f, dow_f = parts

            def matches(field: str, val: int) -> bool:
                for seg in field.split(","):
                    if seg == "*":
                        return True
                    if "/" in seg:
                        base, step = seg.split("/", 1)
                        step = int(step)
                        start = 0 if base == "*" else int(base)
                        if val >= start and (val - start) % step == 0:
                            return True
                    elif "-" in seg:
                        a, b = seg.split("-", 1)
                        if int(a) <= val <= int(b):
                            return True
                    elif int(seg) == val:
                        return True
                return False

            t = (after or datetime.now()).replace(second=0, microsecond=0)
            t += timedelta(minutes=1)
            for _ in range(527040):  # 最多搜索一年
                # cron 周：0=周日…6=周六；Python weekday：0=周一…6=周日
                py_dow = t.weekday()
                cron_dow = (py_dow + 1) % 7
                if (matches(mon_f, t.month) and
                        matches(dom_f, t.day) and
                        matches(dow_f, cron_dow) and
                        matches(h_f, t.hour) and
                        matches(m_f, t.minute)):
                    return t
                t += timedelta(minutes=1)
        except Exception:
            pass
        return None

    def _open_cron_selector(self):
        dlg = _CronDialog(self, self._cron_var.get())
        self.wait_window(dlg)
        if dlg.result:
            self._cron_var.set(dlg.result)

    def _set_cron(self, expr: str):
        self._cron_var.set(expr)

    def _update_schedule_display(self):
        """计算接下来5次触发时间并更新显示标签"""
        expr = self._cron_var.get()
        t = None
        for var in self._next_vars:
            t = self._cron_next(expr, after=t)
            var.set(t.strftime("%Y-%m-%d %H:%M:%S") if t else "—")

    def _update_cron_preview(self):
        expr = self._cron_var.get()
        if self._cron_next(expr):
            self._update_schedule_display()
        else:
            for i, var in enumerate(self._next_vars):
                var.set("（表达式无效）" if i == 0 else "—")

    def _apply_url(self):
        url = self._url_var.get().strip()
        _bg.GOOGLE_SHEET_URL = url
        self._cfg.update({
            "google_sheet_url": url,
            "cron": self._cron_var.get().strip(),
            "overlap": self._overlap_var.get(),
        })
        _save_config(self._cfg)

    def _open_login(self, platform: str):
        """在后台线程中打开平台浏览器登录页，避免阻塞 UI"""
        btn = self._login_btns.get(platform)
        if btn:
            btn.config(state="disabled", text="打开中...")

        def worker():
            try:
                _bg.open_login_page(platform)
                self.log(f"✅ {platform.upper()} 浏览器已打开，请完成登录")
            except Exception as e:
                self.log(f"❌ 打开 {platform.upper()} 失败: {e}")
            finally:
                _labels = {
                    "instagram": "Instagram",
                    "facebook":  "Facebook",
                    "x":         "X (Twitter)",
                    "tiktok":    "TikTok",
                    "threads":   "Threads",
                }
                if btn:
                    self.after(0, lambda: btn.config(
                        state="normal", text=_labels.get(platform, platform)))

        threading.Thread(target=worker, daemon=True).start()

    # ── 持久化 Playwright 工作线程 ───────────────────────────────────

    def _worker_loop(self):
        """所有 Playwright/抓取操作在此线程执行，避免跨线程使用浏览器上下文"""
        while True:
            task = self._scrape_queue.get()
            if task is None:
                break
            try:
                self._exec_scrape()
            finally:
                with self._scrape_cv:
                    self._scrape_active = False
                    self._scrape_cv.notify_all()

    # 立即运行（单次）
    def _run_once(self):
        if self._running:
            return
        self._apply_url()
        self._running = True
        self._stop_evt.clear()
        self._set_status("运行中", _WARN)
        self._set_btns(True)
        for v in self._next_vars: v.set("—")
        threading.Thread(target=self._single_task, daemon=True).start()

    def _single_task(self):
        with self._scrape_cv:
            self._scrape_active = True
        self._scrape_queue.put("scrape")
        with self._scrape_cv:
            while self._scrape_active:
                self._scrape_cv.wait(timeout=1.0)
        self._running = False
        self.after(0, lambda: self._set_status("就绪", _OK))
        self.after(0, lambda: self._set_btns(False))

    # 启动定时循环
    def _start_timer(self):
        if self._running:
            return
        self._apply_url()
        self._running = True
        self._stop_evt.clear()
        self._set_status("定时运行", _ACCENT)
        self._set_btns(True)
        threading.Thread(target=self._timer_loop, daemon=True).start()

    def _timer_loop(self):
        """纯计时线程：按 cron 触发，抓取在独立线程执行"""
        while not self._stop_evt.is_set():
            next_t = self._cron_next(self._cron_var.get())
            if next_t is None:
                self.log("❌ Cron 表达式无效，定时已停止")
                break

            label = next_t.strftime("%Y-%m-%d %H:%M:%S")
            self.after(0, self._update_schedule_display)
            self.log(f"⏰ 下次运行：{label}")

            while datetime.now() < next_t and not self._stop_evt.is_set():
                time.sleep(1)
            if self._stop_evt.is_set():
                break

            # 触发 — 检查是否有任务正在运行
            with self._scrape_cv:
                if self._scrape_active:
                    if self._overlap_var.get() == "skip":
                        self.log("⚠ 上次任务尚未完成，已跳过本次触发")
                        continue
                    else:
                        self.log("⏳ 上次任务尚未完成，排队中...")
                        while self._scrape_active and not self._stop_evt.is_set():
                            self._scrape_cv.wait(timeout=1.0)
                        if self._stop_evt.is_set():
                            break
                        self.log("▶ 排队任务开始执行")
                self._scrape_active = True

            self._scrape_queue.put("scrape")

        # 等待正在进行的抓取结束再更新 UI
        with self._scrape_cv:
            while self._scrape_active:
                self._scrape_cv.wait(timeout=1.0)

        self._running = False
        self.after(0, lambda: self._set_status("已停止", _SUB))
        self.after(0, lambda: self._set_btns(False))
        self.after(0, lambda: [v.set("—") for v in self._next_vars])

    def _exec_scrape(self):
        """抓取体：重定向 stdout，替换 input()"""
        cap = _Capture(self.log)
        old_out    = sys.stdout
        old_input  = builtins.input
        sys.stdout    = cap
        builtins.input = self._gui_input
        try:
            self.log("▶ 开始执行抓取任务...")
            _bg.main()
            self.log("✅ 抓取完成")
        except Exception as e:
            self.log(f"❌ 运行出错: {e}")
        finally:
            sys.stdout    = old_out
            builtins.input = old_input

    def _gui_input(self, prompt: str = "") -> str:
        """替换 bao-guang.py 中的 input()，弹出 GUI 对话框等待用户确认"""
        evt = threading.Event()

        def _show():
            tip = prompt.strip() if prompt.strip() else (
                "请在浏览器中完成登录操作，完成后点击「继续」。")
            messagebox.showinfo("需要操作", tip, parent=self)
            evt.set()

        self.after(0, _show)
        evt.wait()
        return ""

    def _stop_all(self):
        self._stop_evt.set()
        self.log("⏹ 已发出停止信号，等待当前任务结束...")
        self._set_status("停止中...", _WARN)

    def _on_close(self):
        self._stop_evt.set()
        self._scrape_queue.put(None)   # 停止 worker 线程
        try:
            _bg._close_all_contexts()  # 关闭所有浏览器
        except Exception:
            pass
        self.destroy()

# pyinstaller --onedir --windowed --add-data "bao-guang.py;." --collect-all playwright --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import winreg --hidden-import webbrowser --name "曝光量抓取工具" bao-guang-ui.py
if __name__ == "__main__":
    App().mainloop()
