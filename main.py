import tkinter as tk
from tkinter import ttk
import matplotlib
import warnings
import os
import sys
# 新增：高 DPI 支援
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from tkinter.simpledialog import askstring
import PIL.ImageGrab
import io
import win32clipboard
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk, ImageGrab
from io import BytesIO

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

SELECT_GREEN = "#3CB371"

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
        
        # 初始化變數
        self.mode = tk.StringVar(value="個人")
        self.total_rounds = tk.StringVar(value="35")
        self.current_round = tk.StringVar(value="0/35")
        self.player_names = [tk.StringVar() for _ in range(8)]
        self.player_colors = PLAYER_COLORS
        self.team_scores = {}  # 2V2隊伍分數
        self.team_last_scores = {}  # 2V2隊伍上輪得分
        self.group_scores = {"隊伍A": 0, "隊伍B": 0}  # 團體賽分數
        self.group_last_scores = {"隊伍A": 0, "隊伍B": 0}  # 團體賽上輪得分
        
        # 初始化其他變數
        self.score_define_vars = []
        self.id_vars = []
        self.team_result_rows = []
        self.total_scores = [0] * 8  # 恢復為列表
        self.last_scores = [0] * 8   # 恢復為列表
        
        # 新增其他必要的變數初始化
        self.color_vars = []  # 玩家顏色變數
        self.rank_vars = []   # 玩家排名變數
        self.color_dots = []  # 玩家顏色點
        self.target_score_var = tk.IntVar(value=60)  # 目標總分
        self.round_var = tk.IntVar(value=1)
        self.max_round_var = tk.IntVar(value=35)  # 新增：總局數
        self.history = []  # 新增：分數歷史堆疊
        # 初始化歷史記錄（包含團體賽資料）
        self.group_scores = [0, 0]  # 初始化團體賽分數
        self.group_last_scores = [0, 0]
        self.save_history()  # 只在這裡存一次
        self.rank_button_vars = [tk.IntVar(value=0) for _ in range(8)]  # 新增：按鈕式排名狀態
        self.rank_button_widgets = []  # 新增：存放按鈕元件
        self.image_folder = None  # 圖片資料夾路徑
        self.displayed_image = None  # 保持參考避免被GC
        self.folder_path_var = tk.StringVar(value="未設定")
        self.game_mode = tk.StringVar(value="個人賽")  # 新增：比賽模式
        self.teams = []  # 新增：隊伍資訊
        self.team_names = []  # 新增：隊名
        self.team_colors = [] # 新增：隊伍顏色
        self.group_names = [] # 新增：團體賽隊伍名稱
        self.group_colors = [] # 新增：團體賽隊伍顏色
        self.group_scores = [] # 新增：團體賽隊伍總分
        self.group_last_scores = [] # 新增：團體賽隊伍上輪得分
        self.setup_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        self.destroy()
        import sys
        sys.exit(0)

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

        # 先宣告圖片顯示區大小
        self.image_box_width = 150
        self.image_box_height = 230

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
        # 圖片路徑設定按鈕與路徑顯示
        self.folder_path_btn = tk.Button(round_set_row, text="圖片路徑設定", command=self.set_image_folder, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        self.folder_path_btn.pack(side="left", padx=(16,2))
        self.folder_path_label = tk.Label(round_set_row, textvariable=self.folder_path_var, bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 10), anchor="w")
        self.folder_path_label.pack(side="left", padx=(2,0), fill="x", expand=False)
        # 當前局數顯示加回外框
        round_label_frame = ttk.LabelFrame(left_frame, text="當前局數", style="TLabelframe")
        round_label_frame.pack(fill="x", pady=5)
        self.round_label_var = tk.StringVar()
        self.update_round_label()  # 初始化
        self.round_label = ttk.Label(round_label_frame, textvariable=self.round_label_var, font=("Times New Roman", 16, "bold"), background=DARK_BG, foreground=DARK_TEXT)
        self.round_label.pack(padx=10, pady=5)

        # 玩家資訊與截圖區左右分開（用 grid）
        player_and_image_frame = tk.Frame(left_frame, bg=DARK_BG)
        player_and_image_frame.pack(fill="x", pady=5)
        player_and_image_frame.grid_columnconfigure(0, weight=1)
        player_and_image_frame.grid_rowconfigure(0, weight=1)
        # 玩家資訊區（左）
        player_frame = ttk.LabelFrame(player_and_image_frame, text="玩家資訊", style="TLabelframe")
        player_frame.grid(row=0, column=0, sticky="nsew")
        # 欄位標題列
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
        button_grid.grid(row=0, column=1, rowspan=8, sticky="se", padx=4, pady=2)
        button_grid.pack_propagate(False)
        # 分數定義
        score_frame = ttk.LabelFrame(left_frame, text="名次分數定義", style="TLabelframe")
        score_frame.pack(fill="x", pady=5)
        score_frame.configure(style="TLabelframe")
        # 模式選擇
        mode_frame = tk.Frame(score_frame, bg=DARK_BG)
        mode_frame.pack(fill="x", pady=5)
        tk.Label(mode_frame, text="比賽模式:", bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10)).pack(side="left")
        mode_cb = CustomCombobox(mode_frame, values=["個人賽", "2V2組隊", "團體賽"], textvariable=self.game_mode, width=10, state="readonly")
        mode_cb.pack(side="left", padx=5)
        mode_cb.bind("<<ComboboxSelected>>", self.on_mode_change)
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
        self.result_frame.grid(row=0, column=0, pady=(20,0), sticky="n")  # 調整上方間距為20
        self.draw_result_table()
        # 新增：一鍵複製按鈕放在表格區域內
        self.copy_btn = tk.Button(self.result_frame, text="一鍵複製", command=self.copy_result_table_image, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        self.copy_btn.grid(row=9, column=0, columnspan=4, sticky="e", pady=8, padx=8)
        # 隊伍計分顯示區域（紅框部分）
        self.team_result_frame = tk.Frame(right_frame, bg=DARK_BG, bd=2, relief="groove", highlightthickness=1, highlightbackground="red")
        self.team_result_frame.grid(row=1, column=0, pady=(20,10), sticky="ew")  # 調整上方間距為20
        self.team_result_rows = []
        self.draw_team_result_table()
        # 團體賽選隊按鈕區
        self.group_select_var = tk.StringVar(value="")
        self.group_select_frame = tk.Frame(right_frame, bg=DARK_BG)
        self.group_select_frame.grid(row=2, column=0, sticky="w", padx=0, pady=(0,0))
        self.group_select_btns = []
        self.update_group_select_buttons()
        # 2V2組隊一鍵複製按鈕
        self.copy_team_btn = tk.Button(self.team_result_frame, text="一鍵複製", command=self.copy_team_result_table_image, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        self.copy_team_btn.grid(row=5, column=0, columnspan=4, sticky="e", pady=8, padx=8)
        self.copy_team_btn.grid_remove()  # 預設隱藏
        # 預設隱藏隊伍計分區
        self.team_result_frame.grid_remove()
        
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_rowconfigure(1, weight=0)
        right_frame.grid_columnconfigure(0, weight=1)

        # 讓整個主視窗右鍵都能清空排名
        self.bind_all("<Button-3>", self.clear_all_ranks_event)

        # 主視窗左下角顯示最後載入圖片路徑
        self.last_image_path_var = tk.StringVar(value="")
        self.last_image_path_label = tk.Label(self, textvariable=self.last_image_path_var, bg=DARK_BG, fg=DARK_SUBTEXT, font=("Arial", 8), anchor="w")
        self.last_image_path_label.pack(side="bottom", anchor="w", padx=8, pady=(0,2), fill="x")

        # 截圖框區（右）
        self.image_box_width = 150
        self.image_box_height = 240
        image_column_frame = tk.Frame(player_and_image_frame, bg=DARK_BG)
        image_column_frame.grid(row=0, column=1, padx=8, pady=(32,2), sticky="n")
        self.image_frame = tk.Frame(image_column_frame, width=self.image_box_width, height=self.image_box_height, bg=DARK_BG, relief="groove", borderwidth=2)
        self.image_frame.pack(side="top")
        self.image_frame.pack_propagate(False)
        self.image_label = tk.Label(self.image_frame, bg=DARK_BG)
        self.image_label.place(x=0, y=0, width=self.image_box_width, height=self.image_box_height)
        load_img_btn = tk.Button(image_column_frame, text="載入截圖", command=self.show_latest_image, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        load_img_btn.pack(side="top", pady=8)

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

        # 如果是2V2組隊模式，計算隊伍分數
        if self.game_mode.get() == "2V2組隊" and self.teams:
            self.team_result_frame.grid()  # 顯示隊伍計分區
            self.copy_team_btn.grid()  # 顯示一鍵複製
            self.calculate_team_scores()
        # 團體賽：根據選擇給隊伍+1分，並清空選擇
        elif self.game_mode.get() == "團體賽" and hasattr(self, "group_names") and self.group_names:
            sel = self.group_select_var.get()
            if sel in ("0", "1"):
                idx = int(sel)
                self.group_last_scores = [0, 0]
                self.group_last_scores[idx] = 1
                self.group_scores[idx] += 1
                self.group_select_var.set("")
                self.update_group_select_buttons()
            self.calculate_group_scores()
        # 所有模式都在最後統一存一次歷史
        self.save_history()

        # 清空截圖框
        if self.image_label is not None:
            self.image_label.config(image="")
            self.displayed_image = None

    def save_history(self):
        # 存目前所有分數、局數、last_scores
        snapshot = {
            "total_scores": self.total_scores.copy(),
            "last_scores": self.last_scores.copy(),
            "round": self.round_var.get(),
            # 團體賽分數
            "group_scores": self.group_scores.copy(),
            "group_last_scores": self.group_last_scores.copy()
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
        # 團體賽分數回復
        if snapshot["group_scores"] is not None and hasattr(self, "group_scores"):
            self.group_scores = snapshot["group_scores"].copy()
        if snapshot["group_last_scores"] is not None and hasattr(self, "group_last_scores"):
            self.group_last_scores = snapshot["group_last_scores"].copy()
        self.update_result_table()
        self.update_round_label()  # 新增：回溯時也更新顯示
        # 2V2組隊模式下同步隊伍分數顯示
        if self.game_mode.get() == "2V2組隊" and self.teams:
            self.team_result_frame.grid()
            self.calculate_team_scores()
            self.copy_team_btn.grid()
        # 團體賽模式下同步分數顯示
        elif self.game_mode.get() == "團體賽" and hasattr(self, "group_names"):
            self.team_result_frame.grid()
            self.calculate_group_scores()
            self.group_select_frame.grid()
            self.group_select_var.set("")  # 清空選擇
        else:
            self.team_result_frame.grid_remove()
            self.copy_team_btn.grid_remove()

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
        # 比賽模式下拉選單
        mode_var = tk.StringVar(value=self.game_mode.get())
        mode_cb = ttk.Combobox(win, values=["個人賽", "2V2組隊", "團體賽"], textvariable=mode_var, state="readonly", width=10)
        mode_cb.grid(row=0, column=0, columnspan=10, pady=6)

        table_frame = tk.Frame(win, bg=DARK_BG)
        table_frame.grid(row=1, column=0, sticky="nsew")

        def draw_table():
            for widget in table_frame.winfo_children():
                widget.destroy()
            if mode_var.get() == "2V2組隊" and self.teams and self.team_names:
                # 2V2隊伍歷史分數表：只顯示成員為2人的有效隊伍
                valid_teams = []
                for i, team in enumerate(self.teams):
                    if len(team) == 2:
                        valid_teams.append((i, team))
                
                if not valid_teams:
                    tk.Label(table_frame, text="請至少設定一組2人隊伍", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), pady=20).grid(row=0, column=0, columnspan=5)
                    return
                
                tk.Label(table_frame, text="局數", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=0, column=0, sticky="nsew")
                for col, (i, team) in enumerate(valid_teams):
                    name = self.team_names[i] if i < len(self.team_names) else f"隊伍{i+1}"
                    tk.Label(table_frame, text=name, bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=0, column=col+1, sticky="nsew")
                
                team_totals = [0] * len(valid_teams)
                for r, snap in enumerate(self.history[1:]):
                    tk.Label(table_frame, text=str(r+1), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=6).grid(row=r+1, column=0, sticky="nsew")
                    for col, (i, team) in enumerate(valid_teams):
                        s1 = snap["last_scores"][team[0]]
                        s2 = snap["last_scores"][team[1]]
                        team_score = s1 + s2
                        team_totals[col] += team_score
                        tk.Label(table_frame, text=str(team_score), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=10).grid(row=r+1, column=col+1, sticky="nsew")
                
                # 總分列
                tk.Label(table_frame, text="總分", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=len(self.history), column=0, sticky="nsew")
                for col, total in enumerate(team_totals):
                    tk.Label(table_frame, text=str(total), bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=len(self.history), column=col+1, sticky="nsew")
            elif mode_var.get() == "團體賽" and self.group_names:
                # 團體賽歷史分數表：顯示兩隊的總分
                tk.Label(table_frame, text="局數", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=0, column=0, sticky="nsew")
                for i in range(2):
                    name = self.group_names[i] if i < len(self.group_names) else f"隊伍{i+1}"
                    tk.Label(table_frame, text=name, bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=0, column=i+1, sticky="nsew")
                
                group_totals = [0, 0]
                for r, snap in enumerate(self.history[1:]):
                    tk.Label(table_frame, text=str(r+1), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=6).grid(row=r+1, column=0, sticky="nsew")
                    if snap["group_last_scores"] is not None:
                        for i in range(2):
                            score = snap["group_last_scores"][i]
                            group_totals[i] += score
                            tk.Label(table_frame, text=str(score), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=10).grid(row=r+1, column=i+1, sticky="nsew")
                
                # 總分列
                tk.Label(table_frame, text="總分", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=len(self.history), column=0, sticky="nsew")
                for i in range(2):
                    tk.Label(table_frame, text=str(group_totals[i]), bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=len(self.history), column=i+1, sticky="nsew")
            else:
                # 個人賽/團體賽：原本顯示
                tk.Label(table_frame, text="局數", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=0, column=0, sticky="nsew")
                for i in range(8):
                    tk.Label(table_frame, text=self.id_vars[i].get() or f"玩家{i+1}", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=0, column=i+1, sticky="nsew")
                for r, snap in enumerate(self.history[1:]):
                    tk.Label(table_frame, text=str(r+1), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=6).grid(row=r+1, column=0, sticky="nsew")
                    for i in range(8):
                        score = snap["last_scores"][i]
                        tk.Label(table_frame, text=str(score), bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 10), width=10).grid(row=r+1, column=i+1, sticky="nsew")
                # 總分列
                tk.Label(table_frame, text="總分", bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=6).grid(row=len(self.history), column=0, sticky="nsew")
                for i in range(8):
                    total = sum(snap["last_scores"][i] for snap in self.history[1:])
                    tk.Label(table_frame, text=str(total), bg=DARK_HEADER, fg=DARK_TEXT, font=("Times New Roman", 11, "bold"), width=10).grid(row=len(self.history), column=i+1, sticky="nsew")

        mode_cb.bind("<<ComboboxSelected>>", lambda e: draw_table())
        draw_table()

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
        # 比賽模式下拉選單
        mode_var = tk.StringVar(value=self.game_mode.get())
        mode_cb = ttk.Combobox(win, values=["個人賽", "2V2組隊", "團體賽"], textvariable=mode_var, state="readonly", width=10)
        mode_cb.pack(pady=6)

        fig, ax = plt.subplots(figsize=(8, 4), dpi=100)
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.get_tk_widget().pack(fill='both', expand=True)

        def plot_curve():
            ax.clear()
            rounds = len(self.history)
            if mode_var.get() == "2V2組隊" and self.teams and self.team_names:
                # 2V2組隊：只畫成員為2人的有效隊伍
                valid_team_count = 0
                for i, team in enumerate(self.teams):
                    if len(team) == 2:
                        scores = []
                        total = 0
                        for snap in self.history:
                            s1 = snap["last_scores"][team[0]]
                            s2 = snap["last_scores"][team[1]]
                            total += s1 + s2
                            scores.append(total)
                        label = self.team_names[i] if i < len(self.team_names) else f"隊伍{i+1}"
                        color_code = dict(PLAYER_COLORS).get(self.team_colors[i], "#888888") if hasattr(self, "team_colors") and i < len(self.team_colors) else "#888888"
                        ax.plot(range(rounds), scores, marker='o', label=label, color=color_code, linewidth=2)
                        valid_team_count += 1
                if valid_team_count == 0:
                    ax.text(0.5, 0.5, "請至少設定一組2人隊伍", color=DARK_TEXT, fontsize=16, ha='center', va='center', transform=ax.transAxes)
                    canvas.draw()
                    return
            elif mode_var.get() == "團體賽" and self.group_names:
                # 團體賽：畫兩隊曲線，名稱顏色固定index
                for i in range(2):
                    scores = []
                    total = 0
                    for snap in self.history:
                        if snap["group_last_scores"] is not None:
                            total += snap["group_last_scores"][i]
                        scores.append(total)
                    # 固定用初始名稱與顏色
                    color_name = self.group_colors[i]
                    color_code = dict(PLAYER_COLORS).get(color_name, "#888888")
                    label = self.group_names[i]
                    ax.plot(range(rounds), scores, marker='o', label=label, color=color_code, linewidth=2)
            else:
                # 個人賽/團體賽：畫8人曲線，名稱顏色固定index
                for i in range(8):
                    scores = []
                    total = 0
                    for snap in self.history:
                        total += snap["last_scores"][i]
                        scores.append(total)
                    # 固定用初始名稱與顏色
                    color_name = self.color_vars[i].get()
                    color_code = dict(PLAYER_COLORS).get(color_name, "#888888")
                    label = self.id_vars[i].get() or f"玩家{i+1}"
                    ax.plot(range(rounds), scores, marker='o', label=label, color=color_code, linewidth=2)
            ax.set_xlabel("局數", color=DARK_TEXT, fontname='Microsoft JhengHei')
            ax.set_ylabel("總分", color=DARK_TEXT, fontname='Microsoft JhengHei')
            ax.xaxis.set_major_locator(MaxNLocator(integer=True))
            title = "分數曲線"
            ax.set_title(title, color=DARK_TEXT, fontname='Microsoft JhengHei')
            ax.set_facecolor(DARK_BG)
            fig.patch.set_facecolor(DARK_BG)
            ax.tick_params(axis='x', colors=DARK_TEXT, labelsize=10)
            ax.tick_params(axis='y', colors=DARK_TEXT, labelsize=10)
            legend = ax.legend(facecolor=DARK_PANEL, edgecolor=DARK_BORDER, labelcolor=DARK_TEXT, prop={'family': 'Microsoft JhengHei', 'size': 10})
            canvas.draw()

        mode_cb.bind("<<ComboboxSelected>>", lambda e: plot_curve())
        plot_curve()

    def show_capture(self):
        win = tk.Toplevel(self)
        win.title("擷取總分")
        win.iconbitmap(self.ico_path)
        win.configure(bg=DARK_BG)
        block_width = 68
        block_height = 30
        # 不設定 geometry，讓內容自動撐開

        light_colors = ["黃色", "粉色", "白色", "綠色", "青色", "橘色"]

        if self.game_mode.get() == "2V2組隊" and self.teams and self.team_names:
            # 只顯示成員為2人的有效隊伍
            teams = []
            for i, team in enumerate(self.teams):
                if len(team) == 2:
                    total = self.total_scores[team[0]] + self.total_scores[team[1]]
                    color_code = dict(PLAYER_COLORS).get(self.team_colors[i], "#888888") if hasattr(self, "team_colors") and i < len(self.team_colors) else "#888888"
                    teams.append({
                        "name": self.team_names[i],
                        "color_code": color_code,
                        "color_name": self.team_colors[i] if hasattr(self, "team_colors") and i < len(self.team_colors) else "",
                        "total": total
                    })
            if not teams:
                import tkinter.messagebox as mb
                mb.showwarning("隊伍設定不完整", "請至少設定一組2人隊伍，才能擷取隊伍總分！")
                win.destroy()
                return
            teams.sort(key=lambda x: -x["total"])
            for i, t in enumerate(teams):
                frame = tk.Frame(win, bg=DARK_BG)
                frame.grid(row=0, column=i, padx=2, pady=0)
                color_block = tk.Frame(frame, bg=t["color_code"], width=block_width, height=block_height, bd=2, relief="ridge")
                color_block.pack_propagate(False)
                color_block.pack()
                fg_color = "#222831" if t["color_name"] in light_colors else DARK_TEXT
                score_label = tk.Label(color_block, text=str(t["total"]), font=("Times New Roman", 16, "bold"), fg=fg_color, bg=t["color_code"])
                score_label.pack(expand=True, fill='both')
                name_label = tk.Label(frame, text=t["name"], font=("Times New Roman", 10, "bold"), fg=DARK_TEXT, bg=DARK_BG)
                name_label.pack(pady=(0,0))
            return
        elif self.game_mode.get() == "團體賽" and hasattr(self, "group_names") and self.group_names:
            # 團體賽：顯示兩隊總分
            for i in range(2):
                frame = tk.Frame(win, bg=DARK_BG)
                frame.grid(row=0, column=i, padx=2, pady=0)
                color_code = dict(PLAYER_COLORS).get(self.group_colors[i], "#888888")
                color_block = tk.Frame(frame, bg=color_code, width=block_width, height=block_height, bd=2, relief="ridge")
                color_block.pack_propagate(False)
                color_block.pack()
                fg_color = "#222831" if self.group_colors[i] in light_colors else DARK_TEXT
                score_label = tk.Label(color_block, text=str(self.group_scores[i]), font=("Times New Roman", 16, "bold"), fg=fg_color, bg=color_code)
                score_label.pack(expand=True, fill='both')
                name_label = tk.Label(frame, text=self.group_names[i], font=("Times New Roman", 10, "bold"), fg=DARK_TEXT, bg=DARK_BG)
                name_label.pack(pady=(0,0))
            return

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
        self.result_frame.update_idletasks()
        self.result_frame.update()
        # 截圖前先隱藏按鈕
        self.copy_btn.grid_remove()
        self.result_frame.update_idletasks()
        self.result_frame.update()
        x = self.result_frame.winfo_rootx()
        y = self.result_frame.winfo_rooty()
        w = self.result_frame.winfo_width() + 4  # 微調寬度
        h = self.result_frame.winfo_height() + 4  # 微調高度
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

    def copy_team_result_table_image(self):
        """複製隊伍分數表格為圖片（仿照個人賽一鍵複製）"""
        self.team_result_frame.update_idletasks()
        # 截圖前先隱藏按鈕
        self.copy_team_btn.grid_remove()
        self.team_result_frame.update_idletasks()
        self.team_result_frame.update()
        x = self.team_result_frame.winfo_rootx()
        y = self.team_result_frame.winfo_rooty()
        w = self.team_result_frame.winfo_width() + 4  # 微調寬度
        h = self.team_result_frame.winfo_height() + 4  # 微調高度
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            # 複製到剪貼簿
            import win32clipboard
            from io import BytesIO
            output = BytesIO()
            img.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            self.bell()
        except ImportError:
            import tkinter.messagebox as mb
            mb.showerror("缺少套件", "請先安裝 Pillow 套件\n\npip install pillow")
        except Exception as e:
            import tkinter.messagebox as mb
            mb.showerror("複製失敗", f"無法複製圖片到剪貼簿\n{e}")
        finally:
            # 截圖後再顯示按鈕
            self.copy_team_btn.grid(row=5, column=0, columnspan=4, sticky="e", pady=8, padx=8)

    def set_image_folder(self):
        folder = filedialog.askdirectory(title="選擇圖片資料夾")
        if folder:
            self.image_folder = folder
            self.folder_path_var.set(folder)
            # 不自動載入，需按載入截圖

    def show_latest_image(self):
        if not self.image_folder:
            return
        # 找最新的1600x900圖片
        import glob
        import os
        image_files = glob.glob(os.path.join(self.image_folder, '*.png')) + glob.glob(os.path.join(self.image_folder, '*.jpg')) + glob.glob(os.path.join(self.image_folder, '*.jpeg'))
        if not image_files:
            messagebox.showwarning("找不到圖片", "資料夾內沒有圖片檔案！")
            return
        # 選擇更新日期最新的
        image_files = [f for f in image_files if os.path.isfile(f)]
        if not image_files:
            messagebox.showwarning("找不到圖片", "資料夾內沒有圖片檔案！")
            return
        latest_file = max(image_files, key=os.path.getmtime)
        self.last_image_path_var.set(latest_file)
        try:
            img = Image.open(latest_file)
            # 根據原始尺寸選擇截圖範圍
            if img.size == (1600, 900):
                # 1600x900：使用原有截圖範圍
                crop_box = (0, 130, 300, 715)
            elif img.size == (1920, 1080):
                # 1920x1080：使用新的截圖範圍
                crop_box = (0, 215, 300, 800)
            else:
                messagebox.showwarning("尺寸錯誤", f"圖片尺寸需為1600x900或1920x1080，實際為{img.size}")
                return
            # 裁切圖片
            crop_img = img.crop(crop_box)
            # 縮放到框框大小 150x200
            try:
                crop_img = crop_img.resize((self.image_box_width, self.image_box_height), Image.Resampling.LANCZOS)
            except AttributeError:
                messagebox.showerror("Pillow 版本過舊", "請升級 Pillow 至 9.1.0 以上以支援高品質縮放。\n\npip install --upgrade Pillow")
                return
            # 轉成Tkinter可用格式
            tk_img = ImageTk.PhotoImage(crop_img)
            self.displayed_image = tk_img  # 避免被GC
            if self.image_label is not None:
                self.image_label.config(image=tk_img)
        except Exception as e:
            messagebox.showerror("圖片載入錯誤", str(e))

    def on_mode_change(self, event):
        mode = self.game_mode.get()
        if mode == "2V2組隊":
            self.setup_2v2_teams()
            self.team_result_frame.grid()  # 顯示隊伍計分區
            self.copy_team_btn.grid()  # 顯示一鍵複製
            self.group_select_frame.grid_remove()
        elif mode == "團體賽":
            self.setup_group_teams()
            self.team_result_frame.grid()  # 顯示團體分數區
            self.copy_team_btn.grid()  # 顯示一鍵複製
            self.group_select_frame.grid()  # 顯示選隊按鈕
            self.update_group_select_buttons()
        else:
            self.team_result_frame.grid_remove()  # 隱藏隊伍計分區
            self.copy_team_btn.grid_remove()  # 隱藏一鍵複製
            self.group_select_frame.grid_remove()
        self.update_result_table()

    def setup_2v2_teams(self):
        # 2V2組隊設定視窗（4組隊伍，每組2人）
        team_window = tk.Toplevel(self)
        team_window.title("2V2組隊設定")
        team_window.configure(bg=DARK_BG)
        team_window.geometry("500x400")
        team_window.transient(self)
        team_window.grab_set()

        # 隊伍設定區域
        team_frame = tk.Frame(team_window, bg=DARK_BG)
        team_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 建立4個隊伍設定區塊
        team_frames = []
        team_name_vars = []
        team_players = []
        team_color_vars = []
        
        for team_idx in range(4):
            # 隊伍框架
            team_frame_item = ttk.LabelFrame(team_frame, text=f"第{team_idx+1}隊", style="TLabelframe")
            team_frame_item.pack(fill="x", pady=5)
            team_frames.append(team_frame_item)
            
            # 隊名設定
            team_name_var = tk.StringVar(value=f"隊伍{chr(65+team_idx)}")  # A, B, C, D
            team_name_vars.append(team_name_var)
            tk.Label(team_frame_item, text="隊名:", bg=DARK_BG, fg=DARK_TEXT).pack(side="left", padx=5)
            tk.Entry(team_frame_item, textvariable=team_name_var, width=10, bg=DARK_ENTRY, fg=DARK_TEXT).pack(side="left", padx=5)
            
            # 隊伍顏色下拉選單
            color_var = tk.StringVar(value=PLAYER_COLORS[team_idx][0])
            team_color_vars.append(color_var)
            tk.Label(team_frame_item, text="顏色:", bg=DARK_BG, fg=DARK_TEXT).pack(side="left", padx=5)
            color_cb = CustomCombobox(team_frame_item, values=[name for name, _ in PLAYER_COLORS], textvariable=color_var, width=7, state="readonly")
            color_cb.pack(side="left", padx=2)
            dot = ColorDot(team_frame_item, color=dict(PLAYER_COLORS)[color_var.get()])
            dot.pack(side="left", padx=(0,2))
            def update_dot(var=color_var, d=dot):
                d.set_color(dict(PLAYER_COLORS)[var.get()])
            color_var.trace_add('write', lambda *args, var=color_var, d=dot: update_dot(var, d))

            # 隊伍成員選擇（每隊2人）
            team_player_vars = []
            for i in range(2):
                player_var = tk.StringVar()
                # 顯示玩家名稱，如果沒有輸入則顯示預設編號
                player_names = []
                for j in range(8):
                    player_name = self.id_vars[j].get().strip()
                    if player_name:
                        player_names.append(player_name)
                    else:
                        player_names.append(f"玩家{j+1}")
                player_cb = CustomCombobox(team_frame_item, values=player_names, textvariable=player_var, width=8, state="readonly")
                player_cb.pack(side="left", padx=5)
                team_player_vars.append(player_var)
            team_players.append(team_player_vars)

        # 確認按鈕
        def confirm_teams():
            # 儲存隊伍資訊
            self.team_names = [var.get() for var in team_name_vars]
            self.teams = []
            self.team_colors = [var.get() for var in team_color_vars]
            for team_player_vars in team_players:
                team = [int(p.get().replace("玩家", "")) - 1 for p in team_player_vars if p.get()]
                self.teams.append(team)
            
            team_window.destroy()
            
            # 顯示設定結果
            result_text = "隊伍設定完成！\n"
            for i, (team_name, team) in enumerate(zip(self.team_names, self.teams)):
                result_text += f"{team_name}: 玩家{team[0]+1}, 玩家{team[1]+1}\n"
            messagebox.showinfo("設定完成", result_text)

        confirm_btn = tk.Button(team_frame, text="確認設定", command=confirm_teams, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 11, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        confirm_btn.pack(pady=10)

    def setup_group_teams(self):
        # 團體賽設定視窗（2隊，名稱+顏色）
        group_window = tk.Toplevel(self)
        group_window.title("團體賽隊伍設定")
        group_window.configure(bg=DARK_BG)
        group_window.geometry("400x200")
        group_window.transient(self)
        group_window.grab_set()

        group_names = []
        group_color_vars = []
        color_choices = ["紅色", "藍色"]
        for i in range(2):
            frame = tk.Frame(group_window, bg=DARK_BG)
            frame.pack(fill="x", pady=12)
            tk.Label(frame, text=f"第{i+1}隊名稱:", bg=DARK_BG, fg=DARK_TEXT).pack(side="left", padx=8)
            name_var = tk.StringVar(value=f"隊伍{chr(65+i)}")  # A, B, C, D
            group_names.append(name_var)
            tk.Entry(frame, textvariable=name_var, width=10, bg=DARK_ENTRY, fg=DARK_TEXT).pack(side="left", padx=8)
            tk.Label(frame, text="顏色:", bg=DARK_BG, fg=DARK_TEXT).pack(side="left", padx=8)
            color_var = tk.StringVar(value=color_choices[i])
            group_color_vars.append(color_var)
            color_cb = CustomCombobox(frame, values=color_choices, textvariable=color_var, width=6, state="readonly")
            color_cb.pack(side="left", padx=2)
            dot = ColorDot(frame, color=dict(PLAYER_COLORS)[color_var.get()])
            dot.pack(side="left", padx=(0,2))
            def update_dot(var=color_var, d=dot):
                d.set_color(dict(PLAYER_COLORS)[var.get()])
            color_var.trace_add('write', lambda *args, var=color_var, d=dot: update_dot(var, d))

        def confirm_groups():
            self.group_names = [var.get() for var in group_names]
            self.group_colors = [var.get() for var in group_color_vars]
            # 初始化分數
            self.group_scores = [0, 0]
            self.group_last_scores = [0, 0]
            # 初始化歷史記錄
            if len(self.history) == 0:
                self.save_history()
            self.update_group_select_buttons()
            self.group_select_frame.grid()
            group_window.destroy()

        confirm_btn = tk.Button(group_window, text="確認設定", command=confirm_groups, bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 11, "bold"), relief="groove", highlightbackground=DARK_BORDER)
        confirm_btn.pack(pady=10)

    def draw_team_result_table(self):
        # 清除舊的隊伍結果表格
        for widget in self.team_result_frame.winfo_children():
            widget.destroy()
        self.team_result_rows = []
        
        # 隊伍結果表格標題
        header_bg = DARK_HEADER
        header_fg = DARK_TEXT
        header_font = ("Times New Roman", 14, "bold")
        tk.Label(self.team_result_frame, text="隊伍排名", bg=header_bg, fg=header_fg, font=header_font, width=6, pady=6, bd=0, relief="flat").grid(row=0, column=0, sticky="nsew")
        tk.Label(self.team_result_frame, text="隊伍名稱", bg=header_bg, fg=header_fg, font=header_font, width=16, pady=6, bd=0, relief="flat").grid(row=0, column=1, sticky="nsew")
        tk.Label(self.team_result_frame, text="總分", bg=header_bg, fg=header_fg, font=header_font, width=8, pady=6, bd=0, relief="flat").grid(row=0, column=2, sticky="nsew")
        tk.Label(self.team_result_frame, text="上輪得分", bg=header_bg, fg=header_fg, font=header_font, width=10, pady=6, bd=0, relief="flat").grid(row=0, column=3, sticky="nsew")
        
        # 隊伍結果行（最多4隊）
        for i in range(4):
            row = []
            lbl_rank = tk.Label(self.team_result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=6, pady=4, bd=0, relief="flat")
            lbl_rank.grid(row=i+1, column=0, sticky="nsew", padx=(0,1), pady=1)
            name_cell = tk.Frame(self.team_result_frame, bg=DARK_PANEL)
            dot = ColorDot(name_cell, color="#888888")
            dot.pack(side="left", padx=4)
            lbl_name = tk.Label(name_cell, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12))
            lbl_name.pack(side="left", padx=2)
            name_cell.grid(row=i+1, column=1, sticky="nsew", padx=(0,1), pady=1)
            lbl_total = tk.Label(self.team_result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=8, pady=4, bd=0, relief="flat")
            lbl_total.grid(row=i+1, column=2, sticky="nsew", padx=(0,1), pady=1)
            lbl_last = tk.Label(self.team_result_frame, text="", bg=DARK_PANEL, fg=DARK_TEXT, font=("Times New Roman", 12), width=10, pady=4, bd=0, relief="flat")
            lbl_last.grid(row=i+1, column=3, sticky="nsew", padx=(0,1), pady=1)
            row.extend([lbl_rank, dot, lbl_name, lbl_total, lbl_last])
            self.team_result_rows.append(row)
        
        # 設定欄位寬度
        self.team_result_frame.grid_columnconfigure(0, weight=1)
        self.team_result_frame.grid_columnconfigure(1, weight=2)
        self.team_result_frame.grid_columnconfigure(2, weight=1)
        self.team_result_frame.grid_columnconfigure(3, weight=1)

        # 如果是團體賽模式，添加一鍵複製按鈕
        if self.game_mode.get() == "團體賽":
            self.group_copy_btn = tk.Button(self.team_result_frame, text="一鍵複製", command=self.copy_group_result_table_image, 
                                          bg=DARK_PANEL, fg=DARK_TEXT, font=("Arial", 10, "bold"), 
                                          relief="groove", highlightbackground=DARK_BORDER)
            self.group_copy_btn.grid(row=5, column=0, columnspan=4, sticky="e", pady=8, padx=8)

    def calculate_team_scores(self):
        if not self.teams or not self.team_names:
            return
        
        # 計算每隊總分
        team_scores = []
        for i, team in enumerate(self.teams):
            if len(team) == 2:  # 確保隊伍有2人
                player1_score = self.total_scores[team[0]]
                player2_score = self.total_scores[team[1]]
                team_total = player1_score + player2_score
                team_last = self.last_scores[team[0]] + self.last_scores[team[1]]
                team_scores.append({
                    "name": self.team_names[i],
                    "color": dict(PLAYER_COLORS).get(self.team_colors[i], "#888888") if hasattr(self, "team_colors") and i < len(self.team_colors) else "#888888",
                    "total": team_total,
                    "last": team_last,
                    "rank": i + 1
                })
        
        # 按總分排序
        team_scores.sort(key=lambda x: x["total"], reverse=True)
        
        # 更新隊伍排名顯示
        for i, team in enumerate(team_scores):
            if i < len(self.team_result_rows):
                row = self.team_result_rows[i]
                row[0].config(text=str(i + 1))  # 排名
                row[1].set_color(team["color"])  # 顏色圓點
                row[2].config(text=team["name"])  # 隊伍名稱
                row[3].config(text=str(team["total"]))  # 總分
                row[4].config(text=str(team["last"]))  # 上輪得分
        # 清空未使用的行
        for i in range(len(team_scores), len(self.team_result_rows)):
            row = self.team_result_rows[i]
            row[0].config(text="")
            row[1].set_color("#888888")
            row[2].config(text="")
            row[3].config(text="")
            row[4].config(text="")

    def calculate_group_scores(self):
        if not self.group_names:
            return
        
        # 計算每隊總分
        group_scores = []
        for i, group_name in enumerate(self.group_names):
            total = self.group_scores[i]
            last = self.group_last_scores[i]
            group_scores.append({
                "name": group_name,
                "color": dict(PLAYER_COLORS).get(self.group_colors[i], "#888888"),
                "total": total,
                "last": last,
                "rank": i + 1
            })
        
        # 按總分排序
        group_scores.sort(key=lambda x: x["total"], reverse=True)
        
        # 更新隊伍排名顯示
        for i, group in enumerate(group_scores):
            if i < len(self.team_result_rows):
                row = self.team_result_rows[i]
                row[0].config(text=str(i + 1))  # 排名
                row[1].set_color(group["color"])  # 顏色圓點
                row[2].config(text=group["name"])  # 隊伍名稱
                row[3].config(text=str(group["total"]))  # 總分
                row[4].config(text=str(group["last"]))  # 上輪得分
        # 清空未使用的行
        for i in range(len(group_scores), len(self.team_result_rows)):
            row = self.team_result_rows[i]
            row[0].config(text="")
            row[1].set_color("#888888")
            row[2].config(text="")
            row[3].config(text="")
            row[4].config(text="")

    def update_group_select_buttons(self):
        # 只在團體賽顯示
        for widget in self.group_select_frame.winfo_children():
            widget.destroy()
        self.group_select_btns = []
        if hasattr(self, "group_names") and self.game_mode.get() == "團體賽" and self.group_names:
            for i, name in enumerate(self.group_names):
                color = dict(PLAYER_COLORS).get(self.group_colors[i], "#888888") if hasattr(self, "group_colors") and i < len(self.group_colors) else "#888888"
                btn = tk.Radiobutton(self.group_select_frame, text=f"{name}", variable=self.group_select_var, value=str(i), bg=DARK_BG, fg=color, selectcolor=SELECT_GREEN, font=("Arial", 11, "bold"), indicatoron=False, width=8)
                btn.pack(side="left", padx=6, pady=2)
                self.group_select_btns.append(btn)
        else:
            self.group_select_var.set("")

    def copy_group_result_table_image(self):
        """複製團體賽結果表格為圖片"""
        # 獲取團體賽結果表格的位置和大小
        x = self.team_result_frame.winfo_rootx()
        y = self.team_result_frame.winfo_rooty()
        width = self.team_result_frame.winfo_width()
        height = self.team_result_frame.winfo_height()

        # 截取螢幕上的表格區域
        image = ImageGrab.grab(bbox=(x, y, x + width, y + height))

        # 將圖片複製到剪貼簿
        output = BytesIO()
        image.save(output, 'BMP')
        data = output.getvalue()[14:]  # 去除檔頭
        output.close()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

    def draw_2v2_team_scores(self):
        """繪製2V2模式的隊伍分數"""
        # 清除現有的行
        for row in self.team_result_rows[1:]:  # 保留表頭
            for widget in row:
                widget.destroy()
        self.team_result_rows = self.team_result_rows[:1]  # 只保留表頭行

        # 獲取有效的隊伍（至少有一個玩家）
        valid_teams = []
        for i in range(0, 8, 2):
            team_idx = i // 2
            team_name = f"Team {team_idx + 1}"
            if self.player_names[i].get() or self.player_names[i+1].get():
                team_score = self.team_scores.get(team_name, 0)
                team_last_score = self.team_last_scores.get(team_name, 0)
                valid_teams.append((team_name, team_score, team_last_score))

        # 按分數排序
        valid_teams.sort(key=lambda x: x[1], reverse=True)

        # 顯示隊伍分數
        for i, (team_name, total_score, last_score) in enumerate(valid_teams):
            row = []
            # 等位排名
            rank_label = tk.Label(self.team_result_frame, text=str(i+1), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            rank_label.grid(row=i+2, column=0, padx=5, pady=2)
            row.append(rank_label)
            
            # 隊伍名稱
            name_label = tk.Label(self.team_result_frame, text=team_name, bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            name_label.grid(row=i+2, column=1, padx=5, pady=2)
            row.append(name_label)
            
            # 總分
            total_label = tk.Label(self.team_result_frame, text=str(total_score), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            total_label.grid(row=i+2, column=2, padx=5, pady=2)
            row.append(total_label)
            
            # 上輪得分
            last_label = tk.Label(self.team_result_frame, text=str(last_score), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            last_label.grid(row=i+2, column=3, padx=5, pady=2)
            row.append(last_label)
            
            self.team_result_rows.append(row)

    def draw_group_scores(self):
        """繪製團體模式的隊伍分數"""
        # 清除現有的行
        for row in self.team_result_rows[1:]:  # 保留表頭
            for widget in row:
                widget.destroy()
        self.team_result_rows = self.team_result_rows[:1]  # 只保留表頭行

        # 獲取團隊資訊並排序
        teams = []
        team_names = ["隊伍A", "隊伍B"]
        for team_name in team_names:
            team_score = self.group_scores.get(team_name, 0)
            team_last_score = self.group_last_scores.get(team_name, 0)
            teams.append((team_name, team_score, team_last_score))
        teams.sort(key=lambda x: x[1], reverse=True)

        # 顯示團隊分數
        for i, (team_name, total_score, last_score) in enumerate(teams):
            row = []
            # 等位排名
            rank_label = tk.Label(self.team_result_frame, text=str(i+1), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            rank_label.grid(row=i+2, column=0, padx=5, pady=2)
            row.append(rank_label)
            
            # 團隊名稱
            name_label = tk.Label(self.team_result_frame, text=team_name, bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            name_label.grid(row=i+2, column=1, padx=5, pady=2)
            row.append(name_label)
            
            # 總分
            total_label = tk.Label(self.team_result_frame, text=str(total_score), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            total_label.grid(row=i+2, column=2, padx=5, pady=2)
            row.append(total_label)
            
            # 上輪得分
            last_label = tk.Label(self.team_result_frame, text=str(last_score), bg=DARK_BG, fg=DARK_TEXT, font=("Arial", 10))
            last_label.grid(row=i+2, column=3, padx=5, pady=2)
            row.append(last_label)
            
            self.team_result_rows.append(row)

if __name__ == "__main__":
    app = ScoreSystem()
    app.mainloop() 