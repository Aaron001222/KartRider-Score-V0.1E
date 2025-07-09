import tkinter as tk
from tkinter import ttk
import matplotlib
import warnings
import os
import sys
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from tkinter.simpledialog import askstring
import PIL.ImageGrab
import io
import win32clipboard

PLAYER_COLORS = [
    ("紅色", "#FF4C4C"),
    ("藍色", "#4C6FFF"),
    ("綠色", "#4CFF4C"),
    ("黃色", "#FFF44C"),
    ("紫色", "#B44CFF"),
    ("橘色", "#FFB84C"),
    ("青色", "#4CFFF4"),
    ("粉色", "#FFB6C1"),
    ("黑色", "#222831"),
    ("白色", "#EEEEEE"),
    ("灰色", "#CCCCCC")
]
DEFAULT_SCORES = [10, 7, 5, 4, 3, 1, 0, -1, -5]

DARK_BG = "#181A20"
DARK_PANEL = "#23262F"
DARK_ENTRY = "#23262F"
DARK_HEADER = "#111216"
DARK_BORDER = "#393E46"
DARK_TEXT = "#FFFFFF"
DARK_SUBTEXT = "#B0B3B8"

class ColorDot(tk.Canvas):
    def __init__(self, master, color="#FFFFFF", **kwargs):
        try:
            bg = master.cget("background")
        except Exception:
            bg = DARK_BG
        super().__init__(master, width=18, height=18, highlightthickness=0, background=bg)
        self.dot = self.create_oval(3, 3, 15, 15, fill=color, outline="#888", width=1)
    def set_color(self, color):
        self.itemconfig(self.dot, fill=color)

class CustomCombobox(ttk.Combobox):
    # 盡量讓下拉選單內外都黑底白字
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.configure(style="Dark.TCombobox")
        self.bind("<FocusIn>", self._on_focus)
        self.bind("<FocusOut>", self._on_focus)
    def _on_focus(self, event=None):
        self.configure(style="Dark.TCombobox")

class ScoreSystem(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KartRider Score v0.1E (好熊寶製作)")
        if hasattr(sys, "_MEIPASS"):
            # PyInstaller 執行時的臨時資料夾
            icon_path = os.path.join(sys._MEIPASS, "NL.ico")
        else:
            icon_path = os.path.join(os.path.dirname(__file__), "NL.ico")
        self.ico_path = icon_path
        self.iconbitmap(self.ico_path)
        self.configure(bg=DARK_BG)
        self.players = []
        self.score_define_vars = []
        self.id_vars = []
        self.color_vars = []
        self.rank_vars = []
        self.target_score_var = tk.IntVar(value=60)  # 新增：目標總分
        # self.target_rank_vars = []  # 移除目標名次
        self.color_dots = []
        self.result_rows = []
        self.total_scores = [0]*8
        self.last_scores = [0]*8
        self.round_var = tk.IntVar(value=1)
        self.max_round_var = tk.IntVar(value=35)  # 新增：總局數
        self.history = []  # 新增：分數歷史堆疊
        self.save_history()  # 只在這裡存一次
        self.rank_button_vars = [tk.IntVar(value=0) for _ in range(8)]  # 新增：按鈕式排名狀態
        self.rank_button_widgets = []  # 新增：存放按鈕元件
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("TLabel", background=DARK_BG, foreground=DARK_TEXT)
        style.configure("TEntry", fieldbackground=DARK_ENTRY, foreground=DARK_TEXT, borderwidth=2, relief="groove")
        style.configure("TButton", background=DARK_PANEL, foreground=DARK_TEXT, borderwidth=2, relief="groove")
        style.configure("TLabelframe", background=DARK_BG, foreground=DARK_TEXT, borderwidth=2, relief="ridge")
        style.configure("TLabelframe.Label", background=DARK_BG, foreground=DARK_TEXT)
        # 自訂 Combobox 樣式
        style.configure("Dark.TCombobox", fieldbackground=DARK_HEADER, background=DARK_HEADER, foreground=DARK_TEXT, selectbackground=DARK_HEADER, selectforeground=DARK_TEXT)
        style.map("Dark.TCombobox",
            fieldbackground=[('readonly', DARK_HEADER)],
            background=[('readonly', DARK_HEADER)],
            foreground=[('readonly', DARK_TEXT)],
            selectbackground=[('readonly', DARK_HEADER)],
            selectforeground=[('readonly', DARK_TEXT)]
        )

        # 主Frame橫向分割
        main_frame = tk.Frame(self, bg=DARK_BG)
        main_frame.pack(fill="both", expand=True)
        left_frame = tk.Frame(main_frame, bg=DARK_BG)
        left_frame.pack(side="left", fill="y", padx=10, pady=10)
        right_frame = tk.Frame(main_frame, bg=DARK_BG)
        right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # 局數顯示（無『局數』標題，只剩總局數設定框）
        round_frame = tk.Frame(left_frame, bg=DARK_BG)
        round_frame.pack(fill="x", pady=5)
        # 總局數設定框在最上方
        round_set_row = tk.Frame(round_frame, bg=DARK_BG)
        round_set_row.pack(padx=10, pady=2, anchor="w")
        tk.Label(round_set_row, text="總局數:", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10)).pack(side="left")
        max_round_entry = tk.Entry(round_set_row, textvariable=self.max_round_var, width=4, bg=DARK_ENTRY, fg=DARK_TEXT, insertbackground=DARK_TEXT, relief="groove", highlightbackground=DARK_BORDER)
        max_round_entry.pack(side="left", padx=2)
        # 新增：目標總分設定
        tk.Label(round_set_row, text="目標總分:", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10)).pack(side="left", padx=(10,0))
        target_score_entry = tk.Entry(round_set_row, textvariable=self.target_score_var, width=6, bg=DARK_ENTRY, fg=DARK_TEXT, insertbackground=DARK_TEXT, relief="groove", highlightbackground=DARK_BORDER)
        target_score_entry.pack(side="left", padx=2)
        # 當前局數顯示加回外框
        round_label_frame = ttk.LabelFrame(left_frame, text="當前局數", style="TLabelframe")
        round_label_frame.pack(fill="x", pady=5)
        self.round_label_var = tk.StringVar()
        self.update_round_label()  # 初始化
        self.round_label = ttk.Label(round_label_frame, textvariable=self.round_label_var, font=("Times New Roman", 16, "bold"), background=DARK_BG, foreground=DARK_TEXT)
        self.round_label.pack(padx=10, pady=5)

        # 玩家資訊區塊（grid 排版，按鈕區塊和玩家8同一行右側齊平，間隔一致）
        player_frame = ttk.LabelFrame(left_frame, text="玩家資訊", style="TLabelframe")
        player_frame.pack(fill="x", pady=5)
        player_frame.configure(style="TLabelframe")
        # 新增：欄位標題列
        header_row = tk.Frame(player_frame, bg=DARK_BG)
        header_row.grid(row=0, column=0, sticky="ew", pady=(0,2))
        tk.Label(header_row, text="", bg=DARK_BG, width=5).grid(row=0, column=0, padx=2, sticky="ew")
        tk.Label(header_row, text="顏色", bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 10, "bold"), width=12, anchor="w").grid(row=0, column=1, padx=2, sticky="ew")
        tk.Label(header_row, text="玩家名稱", bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 10, "bold"), width=10, anchor="w").grid(row=0, column=2, padx=2, sticky="ew")
        tk.Label(header_row, text="本輪排名", bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 9, "bold"), width=8, anchor="center").grid(row=0, column=3, padx=2, sticky="ew")
        tk.Label(header_row, text="按鈕排名", bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 9, "bold"), width=8, anchor="center").grid(row=0, column=4, padx=2, sticky="ew")
        player_rows = []
        for i in range(8):
            rowf = tk.Frame(player_frame, bg=DARK_BG)
            rowf.grid(row=i+1, column=0, sticky="ew", pady=2)
            player_rows.append(rowf)
            id_var = tk.StringVar()
            color_var = tk.StringVar(value=PLAYER_COLORS[i][0])
            rank_var = tk.StringVar(value="X")  # 預設為X
            self.id_vars.append(id_var)
            self.color_vars.append(color_var)
            self.rank_vars.append(rank_var)
            tk.Label(rowf, text=f"玩家{i+1}", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10), width=5).grid(row=0, column=0, padx=2, sticky="ew")
            # 顏色下拉+圓點同一格
            color_frame = tk.Frame(rowf, bg=DARK_BG)
            color_cb = CustomCombobox(color_frame, values=[name for name, _ in PLAYER_COLORS], textvariable=color_var, width=7, state="readonly")
            color_cb.pack(side="left", padx=(0,2))
            dot = ColorDot(color_frame, color=dict(PLAYER_COLORS)[color_var.get()])
            dot.pack(side="left", padx=(0,2))
            color_frame.grid(row=0, column=1, padx=2, sticky="ew")
            self.color_dots.append(dot)
            def update_dot(var=color_var, d=dot):
                d.set_color(dict(PLAYER_COLORS)[var.get()])
            color_var.trace_add('write', lambda *args, var=color_var, d=dot: update_dot(var, d))
            # 玩家名稱輸入框
            tk.Entry(rowf, textvariable=id_var, width=10, bg=DARK_ENTRY, fg=DARK_TEXT, insertbackground=DARK_TEXT, relief="groove", highlightbackground=DARK_BORDER).grid(row=0, column=2, padx=2, sticky="ew")
            # 本輪排名（下拉選單）
            rank_cb = CustomCombobox(rowf, values=["X"] + [str(x+1) for x in range(8)], textvariable=rank_var, width=4, state="readonly")
            rank_cb.grid(row=0, column=3, padx=2, sticky="ew")
            # 按鈕式排名
            btn = tk.Button(rowf, text="X", width=4, bg=DARK_HEADER, fg=DARK_TEXT, relief="groove", font=("Arial", 10, "bold"))
            btn.grid(row=0, column=4, padx=2, sticky="ew")
            self.rank_button_widgets.append(btn)
            def on_left_click(event, idx=i):
                self.assign_rank_by_button(idx)
            def on_right_click(event, idx=i):
                self.clear_rank_by_button(idx)
            btn.bind("<Button-1>", on_left_click)
            btn.bind("<Button-3>", on_right_click)
            # Combobox 變動時同步按鈕
            def on_rank_cb_change(*args, idx=i):
                self.sync_button_with_combobox(idx)
            rank_var.trace_add('write', on_rank_cb_change)

        # 讓每一行高度一致
        for i in range(8):
            player_frame.grid_rowconfigure(i, weight=1)
        player_frame.grid_columnconfigure(0, weight=0)
        player_frame.grid_columnconfigure(1, weight=1)

        # 2x2 按鈕區塊用 rowspan=8 貼齊底部
        button_grid = tk.Frame(player_frame, bg=DARK_BG)
        btn_font = ("Arial", 11, "bold")
        undo_btn = tk.Button(button_grid, text="上一步", command=self.undo_score, bg=DARK_PANEL, fg=DARK_TEXT, font=btn_font, relief="groove", highlightbackground=DARK_BORDER)
        history_btn = tk.Button(button_grid, text="歷史紀錄", command=self.show_history, bg=DARK_PANEL, fg=DARK_TEXT, font=btn_font, relief="groove", highlightbackground=DARK_BORDER)
        curve_btn = tk.Button(button_grid, text="分數曲線", command=self.show_score_curve, bg=DARK_PANEL, fg=DARK_TEXT, font=btn_font, relief="groove", highlightbackground=DARK_BORDER)
        capture_btn = tk.Button(button_grid, text="擷取", command=self.show_capture, bg=DARK_PANEL, fg=DARK_TEXT, font=btn_font, relief="groove", highlightbackground=DARK_BORDER)
        undo_btn.grid(row=0, column=0, padx=4, pady=2, sticky="nsew")
        history_btn.grid(row=0, column=1, padx=4, pady=2, sticky="nsew")
        curve_btn.grid(row=1, column=0, padx=4, pady=2, sticky="nsew")
        capture_btn.grid(row=1, column=1, padx=4, pady=2, sticky="nsew")
        button_grid.grid_columnconfigure(0, weight=1)
        button_grid.grid_columnconfigure(1, weight=1)
        button_grid.grid_rowconfigure(0, weight=1)
        button_grid.grid_rowconfigure(1, weight=1)
        # 用 rowspan=8 讓下緣齊平
        button_grid.grid(row=0, column=1, rowspan=8, sticky="se", padx=4, pady=2)
        # 讓按鈕貼齊底部
        button_grid.pack_propagate(False)

        # 分數定義
        score_frame = ttk.LabelFrame(left_frame, text="名次分數定義", style="TLabelframe")
        score_frame.pack(fill="x", pady=5)
        score_frame.configure(style="TLabelframe")
        score_row = tk.Frame(score_frame, bg=DARK_BG)
        score_row.pack()
        for i in range(8):
            var = tk.IntVar(value=DEFAULT_SCORES[i])
            self.score_define_vars.append(var)
            tk.Label(score_row, text=f"第{i+1}名", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10)).pack(side="left", padx=2)
            tk.Entry(score_row, textvariable=var, width=4, bg=DARK_ENTRY, fg=DARK_TEXT, insertbackground=DARK_TEXT, relief="groove", highlightbackground=DARK_BORDER).pack(side="left", padx=2)
        # 新增未完成(X)欄位
        var = tk.IntVar(value=DEFAULT_SCORES[8])
        self.score_define_vars.append(var)
        tk.Label(score_row, text="未完成", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10)).pack(side="left", padx=2)
        tk.Entry(score_row, textvariable=var, width=4, bg=DARK_ENTRY, fg=DARK_TEXT, insertbackground=DARK_TEXT, relief="groove", highlightbackground=DARK_BORDER).pack(side="left", padx=2)

        # 計算按鈕
        calc_btn = tk.Button(left_frame, text="計算分數並排序", command=self.calculate_scores, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 11, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        calc_btn.pack(pady=10, fill="x")

        # 結果顯示（仿表格）
        self.result_frame = tk.Frame(right_frame, bg=DARK_BG, bd=0, highlightthickness=0)
        self.result_frame.grid(row=0, column=0, pady=60, sticky="n")  # 往下移動，置中
        self.draw_result_table()
        # 新增：一鍵複製按鈕放在表格區域內
        self.copy_btn = tk.Button(self.result_frame, text="一鍵複製", command=self.copy_result_table_image, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        self.copy_btn.grid(row=9, column=0, columnspan=4, sticky="e", pady=8, padx=8)
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        # 讓整個主視窗右鍵都能清空排名
        self.bind_all("<Button-3>", self.clear_all_ranks_event)

    def draw_result_table(self):
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        header_bg = DARK_HEADER
        header_fg = DARK_TEXT
        header_font = ("Times New Roman", 12, "bold")
        tk.Label(self.result_frame, text="排名", bg=header_bg, fg=header_fg, font=header_font, width=6, pady=6, bd=0, relief="flat").grid(row=0, column=0, sticky="nsew")
        tk.Label(self.result_frame, text="選手", bg=header_bg, fg=header_fg, font=header_font, width=16, pady=6, bd=0, relief="flat").grid(row=0, column=1, sticky="nsew")
        tk.Label(self.result_frame, text="總分", bg=header_bg, fg=header_fg, font=header_font, width=8, pady=6, bd=0, relief="flat").grid(row=0, column=2, sticky="nsew")
        tk.Label(self.result_frame, text="上一輪得分", bg=header_bg, fg=header_fg, font=header_font, width=10, pady=6, bd=0, relief="flat").grid(row=0, column=3, sticky="nsew")
        self.result_rows = []
        for i in range(8):
            row = []
            lbl_rank = tk.Label(self.result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=6, pady=4, bd=0, relief="flat")
            lbl_rank.grid(row=i+1, column=0, sticky="nsew", padx=(0,1), pady=1)
            cell = tk.Frame(self.result_frame, bg=DARK_PANEL)
            dot = ColorDot(cell, color="#FFFFFF")
            dot.pack(side="left", padx=4)
            lbl_name = tk.Label(cell, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12))
            lbl_name.pack(side="left", padx=2)
            cell.grid(row=i+1, column=1, sticky="nsew", padx=(0,1), pady=1)
            lbl_total = tk.Label(self.result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=8, pady=4, bd=0, relief="flat")
            lbl_total.grid(row=i+1, column=2, sticky="nsew", padx=(0,1), pady=1)
            lbl_last = tk.Label(self.result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=10, pady=4, bd=0, relief="flat")
            lbl_last.grid(row=i+1, column=3, sticky="nsew", padx=(0,1), pady=1)
            row.append(lbl_rank)
            row.append(dot)
            row.append(lbl_name)
            row.append(lbl_total)
            row.append(lbl_last)
            self.result_rows.append(row)
        self.result_frame.grid_columnconfigure(0, weight=1)
        self.result_frame.grid_columnconfigure(1, weight=2)
        self.result_frame.grid_columnconfigure(2, weight=1)
        self.result_frame.grid_columnconfigure(3, weight=1)

    def update_round_label(self):
        # 當前局數從0開始，總局數直接顯示
        current_round = self.round_var.get() - 1
        if current_round < 0:
            current_round = 0
        max_round = self.max_round_var.get()
        self.round_label_var.set(f"當前局數: {current_round}/{max_round}")

    def calculate_scores(self):
        score_map = {}
        for i, var in enumerate(self.score_define_vars):
            if i < 8:
                score_map[str(i+1)] = var.get()
            else:
                score_map["X"] = var.get()
        players = []
        for i in range(8):
            pid = self.id_vars[i].get() or f"玩家{i+1}"
            color_name = self.color_vars[i].get()
            color_code = dict(PLAYER_COLORS).get(color_name, "#393E46")
            rank = self.rank_vars[i].get()
            score = score_map.get(rank, 0)
            self.last_scores[i] = score
            self.total_scores[i] += score
            players.append({
                "id": pid,
                "color_name": color_name,
                "color_code": color_code,
                "rank": int(rank) if rank.isdigit() else 99,
                "score": score,
                "total": self.total_scores[i],
                "last": score
            })
        players.sort(key=lambda x: (-x["total"], x["rank"]))
        # 取得目標分數與第一名分數
        target_score = self.target_score_var.get()
        first_score = self.score_define_vars[0].get() if self.score_define_vars else 0
        for i, p in enumerate(players):
            row = self.result_rows[i]
            row[0].config(text=str(i+1))
            row[1].set_color(p['color_code'])
            row[2].config(text=p['id'], bg=DARK_PANEL, fg=DARK_TEXT)
            # 根據分數狀態改變總分欄背景色與文字顏色
            bg = DARK_PANEL
            fg = DARK_TEXT
            if target_score > 0 and first_score > 0:
                if p['total'] >= target_score:
                    bg = "#3CB371"  # 綠色
                    fg = DARK_TEXT
                elif target_score - p['total'] <= first_score:
                    bg = "#FFD700"  # 黃色
                    fg = "#222831"  # 黑色
            row[3].config(text=str(p['total']), bg=bg, fg=fg)
            row[4].config(text=str(p['last']))
        self.round_var.set(self.round_var.get() + 1)
        self.save_history()  # 只在這裡存！
        self.update_round_label()  # 新增：每次計算後更新顯示
        self.clear_all_ranks()  # 新增：計算後自動清空排名
        self.update_result_table()  # 新增：每次計算後刷新顏色提示

    def save_history(self):
        # 存目前所有分數、局數、last_scores
        snapshot = {
            "total_scores": self.total_scores.copy(),
            "last_scores": self.last_scores.copy(),
            "round": self.round_var.get()
        }
        self.history.append(snapshot)

    def undo_score(self):
        if len(self.history) < 2:
            return  # 沒有可回復的狀態
        self.history.pop()  # 移除目前狀態
        snapshot = self.history[-1]
        self.total_scores = snapshot["total_scores"].copy()
        self.last_scores = snapshot["last_scores"].copy()
        self.round_var.set(snapshot["round"])
        self.update_result_table()
        self.update_round_label()  # 新增：回溯時也更新顯示

    def update_result_table(self):
        # 重新顯示結果表格
        players = []
        for i in range(8):
            pid = self.id_vars[i].get() or f"玩家{i+1}"
            color_name = self.color_vars[i].get()
            color_code = dict(PLAYER_COLORS).get(color_name, "#393E46")
            players.append({
                "id": pid,
                "color_name": color_name,
                "color_code": color_code,
                "rank": i+1,
                "score": self.last_scores[i],
                "total": self.total_scores[i],
                "last": self.last_scores[i]
            })
        players.sort(key=lambda x: (-x["total"], x["rank"]))
        # 取得目標分數與第一名分數
        target_score = self.target_score_var.get()
        first_score = self.score_define_vars[0].get() if self.score_define_vars else 0
        for i, p in enumerate(players):
            row = self.result_rows[i]
            row[0].config(text=str(i+1))
            row[1].set_color(p['color_code'])
            row[2].config(text=p['id'], bg=DARK_PANEL, fg=DARK_TEXT)
            # 根據分數狀態改變總分欄文字顏色（不改背景）
            fg = DARK_TEXT
            if target_score > 0 and first_score > 0:
                if p['total'] >= target_score:
                    fg = "#3CB371"  # 綠色
                elif target_score - p['total'] <= first_score:
                    fg = "#FFD700"  # 黃色
            row[3].config(text=str(p['total']), bg=DARK_PANEL, fg=fg)
            row[4].config(text=str(p['last']))

    def show_history(self):
        if not self.history:
            return
        win = tk.Toplevel(self)
        win.title("歷史紀錄")
        win.iconbitmap(self.ico_path)
        win.configure(bg=DARK_BG)
        # 標題
        tk.Label(win, text="局數", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=0, column=0, sticky="nsew")
        for i in range(8):
            tk.Label(win, text=self.id_vars[i].get() or f"玩家{i+1}", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=0, column=i+1, sticky="nsew")
        # 每局分數
        for r, snap in enumerate(self.history[1:]):  # 跳過第一個快照
            tk.Label(win, text=str(r+1), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=6).grid(row=r+1, column=0, sticky="nsew")
            for i in range(8):
                score = snap["last_scores"][i]
                tk.Label(win, text=str(score), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=10).grid(row=r+1, column=i+1, sticky="nsew")

    def show_score_curve(self):
        import matplotlib
        # 中文用微軟正黑體，英文數字仍可用 Times New Roman
        matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Times New Roman', 'SimHei', 'Arial Unicode MS', 'sans-serif']
        matplotlib.rcParams['font.family'] = ['Microsoft JhengHei', 'Times New Roman', 'sans-serif']
        matplotlib.rcParams['axes.unicode_minus'] = False
        if len(self.history) < 2:
            return
        win = tk.Toplevel(self)
        win.title("分數曲線")
        win.iconbitmap(self.ico_path)
        win.configure(bg=DARK_BG)
        fig, ax = plt.subplots(figsize=(8, 4), dpi=100)
        rounds = len(self.history)
        for i in range(8):
            scores = []
            total = 0
            for snap in self.history:
                total += snap["last_scores"][i]
                scores.append(total)
            color_name = self.color_vars[i].get()
            color_code = dict(PLAYER_COLORS).get(color_name, "#888888")
            label = self.id_vars[i].get() or f"玩家{i+1}"
            ax.plot(range(rounds), scores, marker='o', label=label, color=color_code, linewidth=2)
        ax.set_xlabel("局數", color=DARK_TEXT, fontname='Microsoft JhengHei')
        ax.set_ylabel("總分", color=DARK_TEXT, fontname='Microsoft JhengHei')
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        title = askstring("標題設定", "請輸入分數曲線標題：", parent=win)
        if not title:
            title = "玩家分數曲線"
        ax.set_title(title, color=DARK_TEXT, fontname='Microsoft JhengHei')
        ax.set_facecolor(DARK_BG)
        fig.patch.set_facecolor(DARK_BG)
        ax.tick_params(axis='x', colors=DARK_TEXT, labelsize=10)
        ax.tick_params(axis='y', colors=DARK_TEXT, labelsize=10)
        legend = ax.legend(facecolor=DARK_PANEL, edgecolor=DARK_BORDER, labelcolor=DARK_TEXT, prop={'family': 'Microsoft JhengHei', 'size': 10})
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def show_capture(self):
        win = tk.Toplevel(self)
        win.title("擷取總分")
        win.iconbitmap(self.ico_path)
        win.configure(bg=DARK_BG)
        block_width = 68
        block_height = 30
        # 不設定 geometry，讓內容自動撐開

        players = []
        for i in range(8):
            pid = self.id_vars[i].get() or f"玩家{i+1}"
            color_name = self.color_vars[i].get()
            color_code = dict(PLAYER_COLORS).get(color_name, "#888888")
            total = self.total_scores[i]
            players.append({
                "id": pid,
                "color_name": color_name,
                "color_code": color_code,
                "total": total
            })
        players.sort(key=lambda x: -x["total"])

        light_colors = ["黃色", "粉色", "白色", "綠色", "青色", "橘色"]

        for i, p in enumerate(players):
            frame = tk.Frame(win, bg=DARK_BG)
            frame.grid(row=0, column=i, padx=2, pady=0)
            fg_color = "#222831" if p["color_name"] in light_colors else DARK_TEXT
            color_block = tk.Frame(frame, bg=p["color_code"], width=block_width, height=block_height, bd=2, relief="ridge")
            color_block.pack_propagate(False)
            color_block.pack()
            score_label = tk.Label(color_block, text=str(p["total"]), font=("Times New Roman", 16, "bold"), fg=fg_color, bg=p["color_code"])
            score_label.pack(expand=True, fill='both')
            name_label = tk.Label(frame, text=p["id"], font=("Times New Roman", 10, "bold"), fg=DARK_TEXT, bg=DARK_BG)
            name_label.pack(pady=(0,0))

    def assign_rank_by_button(self, idx):
        # 分配下一個可用名次給 idx，若已分配則不動作
        if self.rank_button_vars[idx].get() > 0:
            return
        used = set(var.get() for var in self.rank_button_vars if var.get() > 0)
        for rank in range(1, 9):
            if rank not in used:
                self.rank_button_vars[idx].set(rank)
                self.rank_vars[idx].set(str(rank))
                break
        self.update_rank_buttons()

    def clear_rank_by_button(self, idx):
        # 右鍵清除該玩家名次，並自動調整其他玩家名次
        self.rank_button_vars[idx].set(0)
        self.rank_vars[idx].set("X")
        self.update_rank_buttons()

    def sync_button_with_combobox(self, idx):
        # Combobox 變動時同步按鈕
        val = self.rank_vars[idx].get()
        if val == "X":
            self.rank_button_vars[idx].set(0)
        else:
            try:
                self.rank_button_vars[idx].set(int(val))
            except Exception:
                self.rank_button_vars[idx].set(0)
        self.update_rank_buttons()

    def update_rank_buttons(self):
        # 更新所有按鈕顯示
        for i, btn in enumerate(self.rank_button_widgets):
            val = self.rank_button_vars[i].get()
            btn.config(text=str(val) if val > 0 else "X")

    def clear_all_ranks_event(self, event=None):
        # 右鍵全視窗清空所有本輪排名
        self.clear_all_ranks()

    def clear_all_ranks(self):
        # 清空所有本輪排名（下拉選單與按鈕）
        for i in range(8):
            self.rank_button_vars[i].set(0)
            self.rank_vars[i].set("X")
        self.update_rank_buttons()

    def copy_result_table_image(self):
        # 取得 result_frame 在螢幕上的座標
        self.result_frame.update()
        # 截圖前先隱藏按鈕
        self.copy_btn.grid_remove()
        self.result_frame.update()
        x = self.result_frame.winfo_rootx()
        y = self.result_frame.winfo_rooty()
        w = self.result_frame.winfo_width()
        h = self.result_frame.winfo_height()
        bbox = (x, y, x + w, y + h)
        img = PIL.ImageGrab.grab(bbox)
        # 截圖後再顯示按鈕
        self.copy_btn.grid(row=9, column=0, columnspan=4, sticky="e", pady=8, padx=8)
        # 複製到剪貼簿（Windows 專用）
        output = io.BytesIO()
        img.convert('RGB').save(output, 'BMP')
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

if __name__ == "__main__":
    app = ScoreSystem()
    app.mainloop() 