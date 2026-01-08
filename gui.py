#!/usr/bin/env python3
"""
PowerPoint Image Grid Generator - GUI Interface

This application provides a graphical interface for creating PowerPoint presentations
with images arranged in a grid layout.

Features:
- Interactive REAL-TIME PREVIEW of the slide layout
- Layout Mode (Grid vs Flow)
- Flow Alignment (Left/Center/Right) and Vertical Alignment (Top/Center/Bottom)
- Fixed image size specification
- Crop row/column filters and detailed crop size settings
- Precise Gap Control (cm or scale)
- Image Aspect Ratio & Fit Mode control
- Per-Crop Alignment & Position Control
- Save/Load configuration to/from YAML files

Usage:
  python gui.py                    - Launch GUI application
  python gui.py config.yaml        - Launch with pre-loaded config
"""

import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from typing import Optional

from PIL import Image, ImageTk

from core import (
    GridConfig,
    GapConfig,
    CropRegion,
    CropDisplayConfig,
    LayoutMetrics,
    load_config,
    save_config,
    get_sorted_images,
    create_grid_presentation,
    calculate_grid_metrics,
    calculate_item_bounds,
    calculate_size_fit_static,
    should_apply_crop,
    pt_to_cm,
)


# =============================================================================
# Crop Editor Window
# =============================================================================


class CropEditor(tk.Toplevel):
    """Window for visually selecting crop regions on an image."""

    def __init__(self, parent, image_path: str, callback):
        super().__init__(parent)
        self.title("Crop Editor")
        self.geometry("900x700")
        self.callback = callback

        self.image_path = image_path
        self.orig_img = Image.open(image_path)
        self.orig_w, self.orig_h = self.orig_img.size

        # UI Components
        self.canvas_frame = ttk.Frame(self)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.bottom_frame = ttk.Frame(self, padding=5)
        self.bottom_frame.pack(fill=tk.X)

        ttk.Label(self.bottom_frame, text="Name:").pack(side=tk.LEFT)
        self.var_name = tk.StringVar(value="Region")
        ttk.Entry(self.bottom_frame, textvariable=self.var_name, width=10).pack(
            side=tk.LEFT, padx=5
        )

        ttk.Button(self.bottom_frame, text="保存 (Save)", command=self.on_save).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(self.bottom_frame, text="キャンセル", command=self.destroy).pack(
            side=tk.RIGHT
        )

        # State
        self.rect_id = None
        self.start_x = None
        self.start_y = None
        self.cur_rect = None  # (x, y, w, h) in original coords

        self.display_scale = 1.0
        self.tk_img = None
        self.off_x = 0
        self.off_y = 0

        # Bindings
        self.bind("<Configure>", self.on_resize_window)
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        self.after(100, self.redraw_image)

    def on_resize_window(self, event):
        if event.widget == self:
            self.redraw_image()

    def redraw_image(self):
        c_w = self.canvas.winfo_width()
        c_h = self.canvas.winfo_height()
        if c_w < 50 or c_h < 50:
            return

        scale_w = c_w / self.orig_w
        scale_h = c_h / self.orig_h
        self.display_scale = min(scale_w, scale_h) * 0.9

        new_w = int(self.orig_w * self.display_scale)
        new_h = int(self.orig_h * self.display_scale)

        resized = self.orig_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        self.off_x = (c_w - new_w) // 2
        self.off_y = (c_h - new_h) // 2

        self.canvas.create_image(
            self.off_x, self.off_y, anchor=tk.NW, image=self.tk_img
        )

        # Redraw existing rect if present
        if self.cur_rect:
            x, y, w, h = self.cur_rect
            sx = x * self.display_scale + self.off_x
            sy = y * self.display_scale + self.off_y
            ex = (x + w) * self.display_scale + self.off_x
            ey = (y + h) * self.display_scale + self.off_y
            self.rect_id = self.canvas.create_rectangle(
                sx, sy, ex, ey, outline="red", width=2
            )

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x,
            self.start_y,
            outline="red",
            width=2,
        )

    def on_drag(self, event):
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        end_x, end_y = event.x, event.y

        # Normalize to image coords
        x1 = min(self.start_x, end_x) - self.off_x
        y1 = min(self.start_y, end_y) - self.off_y
        x2 = max(self.start_x, end_x) - self.off_x
        y2 = max(self.start_y, end_y) - self.off_y

        # Convert to original scale
        ox1 = max(0, int(x1 / self.display_scale))
        oy1 = max(0, int(y1 / self.display_scale))
        ox2 = min(self.orig_w, int(x2 / self.display_scale))
        oy2 = min(self.orig_h, int(y2 / self.display_scale))

        w = ox2 - ox1
        h = oy2 - oy1

        if w > 0 and h > 0:
            self.cur_rect = (ox1, oy1, w, h)

    def on_save(self):
        if self.cur_rect:
            self.callback(
                self.cur_rect[0],
                self.cur_rect[1],
                self.cur_rect[2],
                self.cur_rect[3],
                self.var_name.get(),
            )
        self.destroy()


# =============================================================================
# Main Application
# =============================================================================


class ImageGridApp:
    """Main GUI application for PowerPoint grid generation."""

    def __init__(self, root, initial_config: Optional[str] = None):
        self.root = root
        self.root.title("PowerPoint Grid Generator GUI")
        self.root.geometry("1400x950")

        self._init_variables()
        self._create_widgets()
        self._add_preview_tracers()

        # Load initial config if provided
        if initial_config:
            try:
                config = load_config(initial_config)
                self._apply_config_to_gui(config)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load config: {e}")

        self._update_preview()

    def _init_variables(self):
        """Initialize all tkinter variables."""
        # Output settings
        self.output_path = tk.StringVar(value="output.pptx")

        # NEW: input mode (folders vs images)
        self.input_mode = tk.StringVar(value="folders")  # 'folders' | 'images'

        # Grid settings
        self.rows = tk.IntVar(value=3)
        self.cols = tk.IntVar(value=3)
        self.arrangement = tk.StringVar(value="row")
        self.layout_mode = tk.StringVar(value="flow")
        self.flow_align = tk.StringVar(value="left")
        self.flow_vertical_align = tk.StringVar(value="center")

        # Slide settings
        self.slide_w = tk.DoubleVar(value=33.867)
        self.slide_h = tk.DoubleVar(value=19.05)

        # Margins
        self.margin_l = tk.DoubleVar(value=1.0)
        self.margin_t = tk.DoubleVar(value=1.0)
        self.margin_r = tk.DoubleVar(value=1.0)
        self.margin_b = tk.DoubleVar(value=1.0)

        # Gap settings
        self.gap_h_val = tk.DoubleVar(value=0.5)
        self.gap_h_mode = tk.StringVar(value="cm")
        self.gap_v_val = tk.DoubleVar(value=0.5)
        self.gap_v_mode = tk.StringVar(value="cm")
        self.gap_mc_val = tk.DoubleVar(value=0.15)
        self.gap_mc_mode = tk.StringVar(value="cm")
        self.gap_cc_val = tk.DoubleVar(value=0.15)
        self.gap_cc_mode = tk.StringVar(value="cm")
        self.gap_cb_val = tk.DoubleVar(value=0.0)
        self.gap_cb_mode = tk.StringVar(value="cm")

        # Image settings
        self.image_size_mode = tk.StringVar(value="fit")
        self.image_fit_mode = tk.StringVar(value="fit")
        self.image_w = tk.DoubleVar(value=10.0)
        self.image_h = tk.DoubleVar(value=7.5)

        # Crop settings
        self.crop_pos = tk.StringVar(value="right")
        self.crop_size_mode = tk.StringVar(value="scale")
        self.crop_size_val = tk.DoubleVar(value=0.0)
        self.crop_scale_val = tk.DoubleVar(value=0.4)

        # Border settings
        self.show_crop_border = tk.BooleanVar(value=True)
        self.crop_border_w = tk.DoubleVar(value=1.5)
        self.crop_border_shape = tk.StringVar(value="rectangle")
        self.show_zoom_border = tk.BooleanVar(value=True)
        self.zoom_border_w = tk.DoubleVar(value=1.5)
        self.zoom_border_shape = tk.StringVar(value="rectangle")

        # Preview settings
        self.dummy_ratio_w = tk.DoubleVar(value=1.0)
        self.dummy_ratio_h = tk.DoubleVar(value=1.0)

        # Data
        self.folders = []
        self.images = []  # NEW: explicit image list
        self.crop_regions = []
        self.crop_rows_filter = tk.StringVar(value="")
        self.crop_cols_filter = tk.StringVar(value="")

        # Selected region editing
        self.sel_idx = None
        self.r_name = tk.StringVar()
        self.r_x = tk.IntVar()
        self.r_y = tk.IntVar()
        self.r_w = tk.IntVar()
        self.r_h = tk.IntVar()
        self.r_color = (255, 0, 0)
        self.r_align = tk.StringVar(value="auto")
        self.r_offset = tk.DoubleVar(value=0.0)
        self.r_gap = tk.StringVar(value="")

    def _create_widgets(self):
        """Create all UI widgets."""
        # Main paned window
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.Frame(self.paned)
        self.paned.add(self.left_frame, weight=1)

        self.right_frame = ttk.Frame(self.paned)
        self.paned.add(self.right_frame, weight=2)

        # Top toolbar
        top_frame = ttk.Frame(self.left_frame, padding=5)
        top_frame.pack(fill=tk.X)
        ttk.Button(top_frame, text="設定読込", command=self._load_config_gui).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(top_frame, text="設定保存", command=self._save_config).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            top_frame, text="PPTX生成", command=self._generate, style="Accent.TButton"
        ).pack(side=tk.RIGHT, padx=2)

        # Notebook for settings tabs
        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_basic = ttk.Frame(self.notebook, padding=10)
        self.tab_layout = ttk.Frame(self.notebook, padding=10)
        self.tab_crop = ttk.Frame(self.notebook, padding=10)
        self.tab_style = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_basic, text="基本・フォルダ")
        self.notebook.add(self.tab_layout, text="レイアウト")
        self.notebook.add(self.tab_crop, text="クロップ設定")
        self.notebook.add(self.tab_style, text="装飾")

        self._setup_tab_basic()
        self._setup_tab_layout()
        self._setup_tab_crop()
        self._setup_tab_style()

        # Preview area
        f_prev_set = ttk.Frame(self.right_frame, padding=5)
        f_prev_set.pack(fill=tk.X)
        ttk.Label(f_prev_set, text="ダミー画像比率 (W:H)").pack(side=tk.LEFT, padx=5)
        ttk.Entry(f_prev_set, textvariable=self.dummy_ratio_w, width=5).pack(
            side=tk.LEFT
        )
        ttk.Label(f_prev_set, text=":").pack(side=tk.LEFT)
        ttk.Entry(f_prev_set, textvariable=self.dummy_ratio_h, width=5).pack(
            side=tk.LEFT
        )

        ttk.Label(
            self.right_frame,
            text="リアルタイムプレビュー",
            font=("Helvetica", 12, "bold"),
        ).pack(pady=5)

        canvas_frame = ttk.Frame(self.right_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.preview_canvas = tk.Canvas(canvas_frame, bg="#e0e0e0")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self._update_preview())

    def _setup_tab_basic(self):
        """Setup basic settings tab."""
        # Output file
        f_out = ttk.LabelFrame(self.tab_basic, text="出力ファイル", padding=5)
        f_out.pack(fill=tk.X, pady=5)
        ttk.Entry(f_out, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5
        )
        ttk.Button(f_out, text="参照", command=self._browse_output).pack(side=tk.RIGHT)

        # Input mode
        f_mode = ttk.LabelFrame(self.tab_basic, text="入力モード", padding=5)
        f_mode.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(
            f_mode,
            text="フォルダ",
            variable=self.input_mode,
            value="folders",
            command=self._refresh_input_listbox,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_mode,
            text="画像リスト（個別追加）",
            variable=self.input_mode,
            value="images",
            command=self._refresh_input_listbox,
        ).pack(side=tk.LEFT, padx=5)

        # Input folders / images
        f_folders = ttk.LabelFrame(
            self.tab_basic, text="入力（フォルダ/画像）", padding=5
        )
        f_folders.pack(fill=tk.BOTH, expand=True, pady=5)

        btn_frame = ttk.Frame(f_folders)
        btn_frame.pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="追加", command=self._add_input).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="削除", command=self._remove_input).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="クリア", command=self._clear_inputs).pack(
            side=tk.LEFT, padx=2
        )

        self.folder_listbox = tk.Listbox(f_folders, selectmode=tk.EXTENDED)
        self.folder_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Grid configuration
        f_grid = ttk.LabelFrame(self.tab_basic, text="グリッド構成", padding=5)
        f_grid.pack(fill=tk.X, pady=5)

        ttk.Label(f_grid, text="Rows:").grid(row=0, column=0, padx=5)
        ttk.Entry(f_grid, textvariable=self.rows, width=5).grid(row=0, column=1)
        ttk.Label(f_grid, text="Cols:").grid(row=0, column=2, padx=5)
        ttk.Entry(f_grid, textvariable=self.cols, width=5).grid(row=0, column=3)

        ttk.Label(f_grid, text="並び:").grid(row=1, column=0, padx=5)
        ttk.Radiobutton(
            f_grid, text="行(Row)", variable=self.arrangement, value="row"
        ).grid(row=1, column=1)
        ttk.Radiobutton(
            f_grid, text="列(Col)", variable=self.arrangement, value="col"
        ).grid(row=1, column=2)

    def _setup_tab_layout(self):
        """Setup layout settings tab."""
        # Layout mode
        f_mode = ttk.LabelFrame(self.tab_layout, text="レイアウトモード", padding=5)
        f_mode.pack(fill=tk.X, pady=5)

        ttk.Radiobutton(
            f_mode, text="Flow (詰める)", variable=self.layout_mode, value="flow"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_mode, text="Grid (整列)", variable=self.layout_mode, value="grid"
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(f_mode, text="| Flow Align:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(
            f_mode,
            textvariable=self.flow_align,
            values=["left", "center", "right"],
            width=8,
        ).pack(side=tk.LEFT)

        ttk.Label(f_mode, text="| V-Align:").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Combobox(
            f_mode,
            textvariable=self.flow_vertical_align,
            values=["top", "center", "bottom"],
            width=8,
        ).pack(side=tk.LEFT)

        # Slide size
        f_slide = ttk.LabelFrame(self.tab_layout, text="スライドサイズ", padding=5)
        f_slide.pack(fill=tk.X, pady=5)

        ttk.Label(f_slide, text="W:").pack(side=tk.LEFT)
        ttk.Entry(f_slide, textvariable=self.slide_w, width=6).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Label(f_slide, text="H:").pack(side=tk.LEFT)
        ttk.Entry(f_slide, textvariable=self.slide_h, width=6).pack(
            side=tk.LEFT, padx=5
        )

        # Margins
        f_mg = ttk.LabelFrame(self.tab_layout, text="余白(cm)", padding=5)
        f_mg.pack(fill=tk.X, pady=5)

        for label, var in [
            ("L:", self.margin_l),
            ("T:", self.margin_t),
            ("R:", self.margin_r),
            ("B:", self.margin_b),
        ]:
            ttk.Label(f_mg, text=label).pack(side=tk.LEFT)
            ttk.Entry(f_mg, textvariable=var, width=4).pack(side=tk.LEFT)

        # Gap settings
        f_gap = ttk.LabelFrame(self.tab_layout, text="グリッド間隔", padding=5)
        f_gap.pack(fill=tk.X, pady=5)

        ttk.Label(f_gap, text="横(H):").grid(row=0, column=0, sticky=tk.E)
        ttk.Entry(f_gap, textvariable=self.gap_h_val, width=5).grid(row=0, column=1)
        ttk.Radiobutton(f_gap, text="cm", variable=self.gap_h_mode, value="cm").grid(
            row=0, column=2
        )
        ttk.Radiobutton(
            f_gap, text="Scale", variable=self.gap_h_mode, value="scale"
        ).grid(row=0, column=3)

        ttk.Label(f_gap, text="縦(V):").grid(row=1, column=0, sticky=tk.E)
        ttk.Entry(f_gap, textvariable=self.gap_v_val, width=5).grid(row=1, column=1)
        ttk.Radiobutton(f_gap, text="cm", variable=self.gap_v_mode, value="cm").grid(
            row=1, column=2
        )
        ttk.Radiobutton(
            f_gap, text="Scale", variable=self.gap_v_mode, value="scale"
        ).grid(row=1, column=3)

        # Image size settings
        f_img = ttk.LabelFrame(self.tab_layout, text="画像サイズ設定", padding=5)
        f_img.pack(fill=tk.X, pady=5)

        f_size_m = ttk.Frame(f_img)
        f_size_m.pack(fill=tk.X)
        ttk.Radiobutton(
            f_size_m, text="Fit Mode", variable=self.image_size_mode, value="fit"
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_size_m, text="Fixed Size", variable=self.image_size_mode, value="fixed"
        ).pack(side=tk.LEFT)

        f_fit = ttk.Frame(f_img)
        f_fit.pack(fill=tk.X, pady=2)
        ttk.Label(f_fit, text="基準(Fit):").pack(side=tk.LEFT)
        for text, val in [("Fit", "fit"), ("Width", "width"), ("Height", "height")]:
            ttk.Radiobutton(
                f_fit, text=text, variable=self.image_fit_mode, value=val
            ).pack(side=tk.LEFT)

        f_fixed = ttk.Frame(f_img)
        f_fixed.pack(fill=tk.X, pady=2)
        ttk.Label(f_fixed, text="W:").pack(side=tk.LEFT)
        ttk.Entry(f_fixed, textvariable=self.image_w, width=5).pack(side=tk.LEFT)
        ttk.Label(f_fixed, text="H:").pack(side=tk.LEFT)
        ttk.Entry(f_fixed, textvariable=self.image_h, width=5).pack(side=tk.LEFT)

    def _setup_tab_crop(self):
        """Setup crop settings tab."""
        # Region list
        f_reg = ttk.LabelFrame(self.tab_crop, text="領域リスト", padding=5)
        f_reg.pack(fill=tk.X, pady=5)

        r_btn_frame = ttk.Frame(f_reg)
        r_btn_frame.pack(fill=tk.X)
        ttk.Button(
            r_btn_frame, text="画像から指定 (Editor)", command=self._open_crop_editor
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            r_btn_frame, text="追加 (数値)", command=self._add_region_dialog
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(r_btn_frame, text="削除", command=self._remove_region).pack(
            side=tk.LEFT, padx=2
        )

        self.region_tree = ttk.Treeview(
            f_reg, columns=("name", "xywh", "align"), show="headings", height=4
        )
        self.region_tree.heading("name", text="Name")
        self.region_tree.heading("xywh", text="Coord")
        self.region_tree.heading("align", text="Align")
        self.region_tree.pack(fill=tk.BOTH, expand=True)
        self.region_tree.bind("<<TreeviewSelect>>", self._on_region_select)

        # Selected region detail
        f_detail = ttk.LabelFrame(
            self.tab_crop, text="選択した領域の詳細設定", padding=5
        )
        f_detail.pack(fill=tk.X, pady=5)

        f_r1 = ttk.Frame(f_detail)
        f_r1.pack(fill=tk.X)
        ttk.Label(f_r1, text="Name:").pack(side=tk.LEFT)
        ttk.Entry(f_r1, textvariable=self.r_name, width=10).pack(side=tk.LEFT, padx=5)
        self.btn_r_color = tk.Button(
            f_r1, text="Color", width=5, command=self._pick_region_color
        )
        self.btn_r_color.pack(side=tk.LEFT, padx=5)

        f_r2 = ttk.Frame(f_detail)
        f_r2.pack(fill=tk.X)
        for label, var in [
            ("X:", self.r_x),
            ("Y:", self.r_y),
            ("W:", self.r_w),
            ("H:", self.r_h),
        ]:
            ttk.Label(f_r2, text=label).pack(side=tk.LEFT)
            ttk.Entry(f_r2, textvariable=var, width=5).pack(side=tk.LEFT)

        f_r3 = ttk.Frame(f_detail)
        f_r3.pack(fill=tk.X)
        ttk.Label(f_r3, text="Align:").pack(side=tk.LEFT)
        ttk.Combobox(
            f_r3,
            textvariable=self.r_align,
            values=["auto", "start", "center", "end"],
            width=7,
        ).pack(side=tk.LEFT)
        ttk.Label(f_r3, text="Offset:").pack(side=tk.LEFT)
        ttk.Entry(f_r3, textvariable=self.r_offset, width=5).pack(side=tk.LEFT)
        ttk.Label(f_r3, text="Gap:").pack(side=tk.LEFT)
        ttk.Entry(f_r3, textvariable=self.r_gap, width=5).pack(side=tk.LEFT)

        f_r4 = ttk.Frame(f_detail)
        f_r4.pack(fill=tk.X, pady=5)
        ttk.Button(f_r4, text="更新 (Update)", command=self._update_region_detail).pack(
            anchor=tk.E
        )

        # Global crop settings
        f_glob = ttk.LabelFrame(self.tab_crop, text="全体・配置設定", padding=5)
        f_glob.pack(fill=tk.X, pady=5)

        ttk.Label(f_glob, text="配置方向:").grid(row=0, column=0, sticky=tk.E)
        f_pos = ttk.Frame(f_glob)
        f_pos.grid(row=0, column=1, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(
            f_pos, text="Right", variable=self.crop_pos, value="right"
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_pos, text="Bottom", variable=self.crop_pos, value="bottom"
        ).pack(side=tk.LEFT)

        # Gap settings
        for r, (label, val_var, mode_var) in enumerate(
            [
                ("Main-Crop:", self.gap_mc_val, self.gap_mc_mode),
                ("Crop-Crop:", self.gap_cc_val, self.gap_cc_mode),
                ("Crop-Bottom:", self.gap_cb_val, self.gap_cb_mode),
            ],
            1,
        ):
            ttk.Label(f_glob, text=label).grid(row=r, column=0, sticky=tk.E)
            ttk.Entry(f_glob, textvariable=val_var, width=5).grid(row=r, column=1)
            ttk.Radiobutton(f_glob, text="cm", variable=mode_var, value="cm").grid(
                row=r, column=2
            )
            ttk.Radiobutton(
                f_glob, text="Scale", variable=mode_var, value="scale"
            ).grid(row=r, column=3)

        # Size settings
        ttk.Label(f_glob, text="サイズ:").grid(row=4, column=0, sticky=tk.E)
        f_sz = ttk.Frame(f_glob)
        f_sz.grid(row=4, column=1, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(
            f_sz, text="倍率(Scale)", variable=self.crop_size_mode, value="scale"
        ).pack(side=tk.LEFT)
        ttk.Entry(f_sz, textvariable=self.crop_scale_val, width=5).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_sz, text="固定(cm)", variable=self.crop_size_mode, value="size"
        ).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Entry(f_sz, textvariable=self.crop_size_val, width=5).pack(side=tk.LEFT)

        # Filter settings
        f_filter = ttk.LabelFrame(self.tab_crop, text="適用対象 (空欄=全て)", padding=5)
        f_filter.pack(fill=tk.X, pady=5)

        ttk.Label(f_filter, text="行:").pack(side=tk.LEFT)
        ttk.Entry(f_filter, textvariable=self.crop_rows_filter, width=10).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Label(f_filter, text="列:").pack(side=tk.LEFT)
        ttk.Entry(f_filter, textvariable=self.crop_cols_filter, width=10).pack(
            side=tk.LEFT, padx=5
        )

    def _setup_tab_style(self):
        """Setup style settings tab."""
        f_style = ttk.LabelFrame(self.tab_style, text="枠線設定", padding=5)
        f_style.pack(fill=tk.X, pady=5)

        # Source image border
        f_src = ttk.LabelFrame(f_style, text="元画像上の枠線 (Source Image)", padding=5)
        f_src.pack(fill=tk.X, pady=5)

        f_src_row1 = ttk.Frame(f_src)
        f_src_row1.pack(fill=tk.X)
        ttk.Checkbutton(
            f_src_row1, text="表示する", variable=self.show_crop_border
        ).pack(side=tk.LEFT)
        ttk.Label(f_src_row1, text="太さ (pt):").pack(side=tk.LEFT, padx=(20, 5))
        ttk.Entry(f_src_row1, textvariable=self.crop_border_w, width=5).pack(
            side=tk.LEFT
        )

        f_src_row2 = ttk.Frame(f_src)
        f_src_row2.pack(fill=tk.X, pady=5)
        ttk.Label(f_src_row2, text="形状:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_src_row2,
            text="角 (Rectangle)",
            variable=self.crop_border_shape,
            value="rectangle",
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_src_row2,
            text="丸 (Rounded)",
            variable=self.crop_border_shape,
            value="rounded",
        ).pack(side=tk.LEFT, padx=5)

        # Cropped image border
        f_zoom = ttk.LabelFrame(
            f_style, text="クロップ画像の枠線 (Cropped Image)", padding=5
        )
        f_zoom.pack(fill=tk.X, pady=5)

        f_zoom_row1 = ttk.Frame(f_zoom)
        f_zoom_row1.pack(fill=tk.X)
        ttk.Checkbutton(
            f_zoom_row1, text="表示する", variable=self.show_zoom_border
        ).pack(side=tk.LEFT)
        ttk.Label(f_zoom_row1, text="太さ (pt):").pack(side=tk.LEFT, padx=(20, 5))
        ttk.Entry(f_zoom_row1, textvariable=self.zoom_border_w, width=5).pack(
            side=tk.LEFT
        )

        f_zoom_row2 = ttk.Frame(f_zoom)
        f_zoom_row2.pack(fill=tk.X, pady=5)
        ttk.Label(f_zoom_row2, text="形状:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_zoom_row2,
            text="角 (Rectangle)",
            variable=self.zoom_border_shape,
            value="rectangle",
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_zoom_row2,
            text="丸 (Rounded)",
            variable=self.zoom_border_shape,
            value="rounded",
        ).pack(side=tk.LEFT, padx=5)

    def _add_preview_tracers(self):
        """Add trace callbacks for automatic preview updates."""
        vars_to_trace = [
            self.rows,
            self.cols,
            self.arrangement,
            self.layout_mode,
            self.flow_align,
            self.flow_vertical_align,
            self.slide_w,
            self.slide_h,
            self.margin_l,
            self.margin_t,
            self.margin_r,
            self.margin_b,
            self.gap_h_val,
            self.gap_h_mode,
            self.gap_v_val,
            self.gap_v_mode,
            self.image_size_mode,
            self.image_fit_mode,
            self.image_w,
            self.image_h,
            self.crop_pos,
            self.gap_mc_val,
            self.gap_mc_mode,
            self.gap_cc_val,
            self.gap_cc_mode,
            self.gap_cb_val,
            self.gap_cb_mode,
            self.crop_size_mode,
            self.crop_size_val,
            self.crop_scale_val,
            self.crop_rows_filter,
            self.crop_cols_filter,
            self.dummy_ratio_w,
            self.dummy_ratio_h,
            self.r_name,
            self.r_x,
            self.r_y,
            self.r_w,
            self.r_h,
            self.r_align,
            self.r_offset,
            self.r_gap,
            self.zoom_border_shape,
            self.show_zoom_border,
            self.zoom_border_w,
            self.crop_border_shape,
            self.show_crop_border,
            self.crop_border_w,
        ]
        for v in vars_to_trace:
            v.trace_add("write", lambda *args: self._schedule_preview())

    def _schedule_preview(self):
        """Schedule a preview update with debouncing."""
        if hasattr(self, "_after_id"):
            self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(100, self._update_preview)

    # -------------------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------------------

    def _get_safe_int(self, var, default=0):
        try:
            return var.get()
        except tk.TclError:
            return default

    def _get_safe_double(self, var, default=0.0):
        try:
            return var.get()
        except tk.TclError:
            return default

    def _get_current_config(self) -> GridConfig:
        """Build GridConfig from current GUI state."""
        c = GridConfig()

        c.layout_mode = self.layout_mode.get()
        c.flow_align = self.flow_align.get()
        c.flow_vertical_align = self.flow_vertical_align.get()
        c.slide_width = self._get_safe_double(self.slide_w, 33.867)
        c.slide_height = self._get_safe_double(self.slide_h, 19.05)
        c.rows = max(1, self._get_safe_int(self.rows, 2))
        c.cols = max(1, self._get_safe_int(self.cols, 3))
        c.arrangement = self.arrangement.get()

        c.margin_left = self._get_safe_double(self.margin_l, 1.0)
        c.margin_top = self._get_safe_double(self.margin_t, 1.0)
        c.margin_right = self._get_safe_double(self.margin_r, 1.0)
        c.margin_bottom = self._get_safe_double(self.margin_b, 1.0)

        c.gap_h = GapConfig(
            self._get_safe_double(self.gap_h_val, 0.5), self.gap_h_mode.get()
        )
        c.gap_v = GapConfig(
            self._get_safe_double(self.gap_v_val, 0.5), self.gap_v_mode.get()
        )

        c.size_mode = self.image_size_mode.get()
        c.fit_mode = self.image_fit_mode.get()
        c.image_width = self._get_safe_double(self.image_w, 10.0)
        c.image_height = self._get_safe_double(self.image_h, 7.5)

        # Inputs
        if self.input_mode.get() == "images":
            c.images = self.images[:]
            c.folders = []
        else:
            c.images = None
            c.folders = (
                self.folders if self.folders else ["dummy"] * max(1, c.rows * c.cols)
            )

        c.crop_regions = self.crop_regions
        c.output = self.output_path.get()

        # Crop filters
        rs = self.crop_rows_filter.get().strip()
        if rs:
            c.crop_rows = [int(x.strip()) for x in rs.split(",") if x.strip()]
        cs = self.crop_cols_filter.get().strip()
        if cs:
            c.crop_cols = [int(x.strip()) for x in cs.split(",") if x.strip()]

        # Crop display
        c.crop_display.position = self.crop_pos.get()
        c.crop_display.main_crop_gap = GapConfig(
            self._get_safe_double(self.gap_mc_val, 0.15), self.gap_mc_mode.get()
        )
        c.crop_display.crop_crop_gap = GapConfig(
            self._get_safe_double(self.gap_cc_val, 0.15), self.gap_cc_mode.get()
        )
        c.crop_display.crop_bottom_gap = GapConfig(
            self._get_safe_double(self.gap_cb_val, 0.0), self.gap_cb_mode.get()
        )

        if self.crop_size_mode.get() == "size":
            val = self._get_safe_double(self.crop_size_val, 0.0)
            if val > 0:
                c.crop_display.size = val
        else:
            val = self._get_safe_double(self.crop_scale_val, 0.4)
            if val > 0:
                c.crop_display.scale = val

        # Border settings
        c.zoom_border_shape = self.zoom_border_shape.get()
        c.crop_border_shape = self.crop_border_shape.get()
        c.show_crop_border = self.show_crop_border.get()
        c.crop_border_width = self._get_safe_double(self.crop_border_w, 1.5)
        c.show_zoom_border = self.show_zoom_border.get()
        c.zoom_border_width = self._get_safe_double(self.zoom_border_w, 1.5)

        return c

    def _apply_config_to_gui(self, c: GridConfig):
        """Apply a GridConfig to the GUI state."""
        self.layout_mode.set(c.layout_mode)
        self.flow_align.set(c.flow_align)
        self.flow_vertical_align.set(c.flow_vertical_align)
        self.rows.set(c.rows)
        self.cols.set(c.cols)
        self.arrangement.set(c.arrangement)
        self.slide_w.set(c.slide_width)
        self.slide_h.set(c.slide_height)

        self.margin_l.set(c.margin_left)
        self.margin_t.set(c.margin_top)
        self.margin_r.set(c.margin_right)
        self.margin_b.set(c.margin_bottom)

        self.gap_h_val.set(c.gap_h.value)
        self.gap_h_mode.set(c.gap_h.mode)
        self.gap_v_val.set(c.gap_v.value)
        self.gap_v_mode.set(c.gap_v.mode)

        self.image_size_mode.set(c.size_mode)
        self.image_fit_mode.set(c.fit_mode)
        if c.image_width:
            self.image_w.set(c.image_width)
        if c.image_height:
            self.image_h.set(c.image_height)

        # Inputs
        if c.images:
            self.input_mode.set("images")
            self.images = c.images
            self.folders = []
        else:
            self.input_mode.set("folders")
            self.folders = c.folders
            self.images = []

        self._refresh_input_listbox()

        self.crop_regions = c.crop_regions
        self._update_region_list()

        self.crop_rows_filter.set(
            ",".join(map(str, c.crop_rows)) if c.crop_rows else ""
        )
        self.crop_cols_filter.set(
            ",".join(map(str, c.crop_cols)) if c.crop_cols else ""
        )

        self.crop_pos.set(c.crop_display.position)
        self.gap_mc_val.set(c.crop_display.main_crop_gap.value)
        self.gap_mc_mode.set(c.crop_display.main_crop_gap.mode)
        self.gap_cc_val.set(c.crop_display.crop_crop_gap.value)
        self.gap_cc_mode.set(c.crop_display.crop_crop_gap.mode)
        self.gap_cb_val.set(c.crop_display.crop_bottom_gap.value)
        self.gap_cb_mode.set(c.crop_display.crop_bottom_gap.mode)

        if c.crop_display.size:
            self.crop_size_mode.set("size")
            self.crop_size_val.set(c.crop_display.size)
        elif c.crop_display.scale:
            self.crop_size_mode.set("scale")
            self.crop_scale_val.set(c.crop_display.scale)

        self.zoom_border_shape.set(c.zoom_border_shape)
        self.crop_border_shape.set(c.crop_border_shape)
        self.show_crop_border.set(c.show_crop_border)
        self.crop_border_w.set(c.crop_border_width)
        self.show_zoom_border.set(c.show_zoom_border)
        self.zoom_border_w.set(c.zoom_border_width)

    # -------------------------------------------------------------------------
    # Event Handlers
    # -------------------------------------------------------------------------

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PPTX", "*.pptx")]
        )
        if path:
            self.output_path.set(path)

    def _refresh_input_listbox(self):
        self.folder_listbox.delete(0, tk.END)
        if self.input_mode.get() == "images":
            for p in self.images:
                self.folder_listbox.insert(tk.END, p)
        else:
            for p in self.folders:
                self.folder_listbox.insert(tk.END, p)
        self._schedule_preview()

    def _add_input(self):
        if self.input_mode.get() == "images":
            paths = filedialog.askopenfilenames(
                filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.tiff;*.gif")]
            )
            if not paths:
                return
            self.images.extend(list(paths))
            self._refresh_input_listbox()

            # Auto-detect aspect ratio from first added image
            try:
                with Image.open(self.images[0]) as img:
                    w, h = img.size
                    self.dummy_ratio_w.set(1.0)
                    self.dummy_ratio_h.set(h / w)
            except Exception:
                pass
        else:
            p = filedialog.askdirectory()
            if not p:
                return
            self.folders.append(p)
            self._refresh_input_listbox()

            # Auto-detect aspect ratio from first image in folder
            try:
                images = get_sorted_images(p)
                if images:
                    with Image.open(images[0]) as img:
                        w, h = img.size
                        self.dummy_ratio_w.set(1.0)
                        self.dummy_ratio_h.set(h / w)
            except Exception:
                pass

    def _remove_input(self):
        sel = list(self.folder_listbox.curselection())
        if not sel:
            return
        for i in reversed(sel):
            if self.input_mode.get() == "images":
                if 0 <= i < len(self.images):
                    self.images.pop(i)
            else:
                if 0 <= i < len(self.folders):
                    self.folders.pop(i)
        self._refresh_input_listbox()

    def _clear_inputs(self):
        self.folders = []
        self.images = []
        self._refresh_input_listbox()

    def _open_crop_editor(self):
        target_img = None

        # NEW: prioritize actual cell images
        if self.input_mode.get() == "images":
            if self.images:
                target_img = self.images[0]
        else:
            for f in self.folders:
                imgs = get_sorted_images(f)
                if imgs:
                    target_img = imgs[0]
                    break

        if not target_img:
            target_img = filedialog.askopenfilename(
                filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.webp")]
            )

        if target_img:
            CropEditor(self.root, target_img, self._add_region_from_editor)

    def _get_image_path_for_cell(self, row: int, col: int) -> Optional[str]:
        config = self._get_current_config()
        if config.images:
            idx = row * config.cols + col
            if 0 <= idx < len(config.images):
                return config.images[idx]
            return None

        folders = config.folders or []
        if config.arrangement == "row":
            if row >= len(folders):
                return None
            imgs = get_sorted_images(folders[row])
            return imgs[col] if col < len(imgs) else None
        else:
            if col >= len(folders):
                return None
            imgs = get_sorted_images(folders[col])
            return imgs[row] if row < len(imgs) else None

    def _on_preview_cell_click(self, row: int, col: int):
        p = self._get_image_path_for_cell(row, col)
        if not p:
            p = filedialog.askopenfilename(
                filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.webp")]
            )
        if p:
            CropEditor(self.root, p, self._add_region_from_editor)

    def _update_preview(self):
        """Render the preview canvas."""
        config = self._get_current_config()
        canvas = self.preview_canvas
        canvas.delete("all")

        c_w, c_h = canvas.winfo_width(), canvas.winfo_height()
        if c_w < 10:
            return

        scale = min((c_w - 20) / config.slide_width, (c_h - 20) / config.slide_height)
        sx = (c_w - config.slide_width * scale) / 2
        sy = (c_h - config.slide_height * scale) / 2

        def tx(v):
            return sx + v * scale

        def ty(v):
            return sy + v * scale

        # Draw slide background
        canvas.create_rectangle(
            tx(0),
            ty(0),
            tx(config.slide_width),
            ty(config.slide_height),
            fill="white",
            outline="black",
        )

        metrics = calculate_grid_metrics(config)

        # Get dummy image ratio
        d_rw = max(0.01, self._get_safe_double(self.dummy_ratio_w, 1.0))
        d_rh = max(0.01, self._get_safe_double(self.dummy_ratio_h, 1.0))
        dummy_w_px = 400
        dummy_h_px = 400 * (d_rh / d_rw)

        border_offset_cm = (
            pt_to_cm(config.zoom_border_width) if config.show_zoom_border else 0.0
        )
        half_border = border_offset_cm / 2.0 if config.show_zoom_border else 0.0

        # Pre-calculate flow layout
        total_content_height = 0.0
        row_heights_flow = []

        if config.layout_mode == "flow":
            for r in range(config.rows):
                sim_row_h = 0.0
                for c in range(config.cols):
                    img_w_cm, img_h_cm = calculate_size_fit_static(
                        dummy_w_px,
                        dummy_h_px,
                        metrics.main_width,
                        metrics.main_height,
                        config.fit_mode,
                    )
                    override_sz = (img_w_cm, img_h_cm)
                    min_x, min_y, max_x, max_y = calculate_item_bounds(
                        config,
                        metrics,
                        "dummy",
                        r,
                        c,
                        border_offset_cm,
                        override_size=override_sz,
                    )
                    item_h = max_y - min_y
                    sim_row_h = max(sim_row_h, item_h)

                row_heights_flow.append(
                    sim_row_h if sim_row_h > 0 else metrics.main_height
                )
                total_content_height += row_heights_flow[-1]
                if r < config.rows - 1:
                    total_content_height += config.gap_v.to_cm(metrics.main_height)

        # Calculate starting Y position
        cy = config.margin_top
        if config.layout_mode == "flow":
            avail_h = config.slide_height - config.margin_top - config.margin_bottom
            if config.flow_vertical_align == "center":
                cy = config.margin_top + (avail_h - total_content_height) / 2
            elif config.flow_vertical_align == "bottom":
                cy = (config.margin_top + avail_h) - total_content_height

        # Draw grid
        for r in range(config.rows):
            current_row_h = (
                row_heights_flow[r]
                if config.layout_mode == "flow" and r < len(row_heights_flow)
                else (
                    metrics.row_heights[r]
                    if r < len(metrics.row_heights)
                    else metrics.main_height
                )
            )

            cx = config.margin_left

            # Calculate row content width for flow alignment
            if config.layout_mode == "flow":
                row_content_width = 0.0
                valid_items = 0

                for c in range(config.cols):
                    img_w_cm, img_h_cm = calculate_size_fit_static(
                        dummy_w_px,
                        dummy_h_px,
                        metrics.main_width,
                        metrics.main_height,
                        config.fit_mode,
                    )
                    min_x, min_y, max_x, max_y = calculate_item_bounds(
                        config,
                        metrics,
                        "dummy",
                        r,
                        c,
                        border_offset_cm,
                        override_size=(img_w_cm, img_h_cm),
                    )
                    row_content_width += max_x - min_x
                    valid_items += 1

                if valid_items > 1:
                    row_content_width += (valid_items - 1) * config.gap_h.to_cm(
                        metrics.main_width
                    )

                avail_w = config.slide_width - config.margin_left - config.margin_right
                if config.flow_align == "center":
                    cx = config.margin_left + (avail_w - row_content_width) / 2
                elif config.flow_align == "right":
                    cx = config.margin_left + (avail_w - row_content_width)

            # Draw cells
            for c in range(config.cols):
                img_w_cm, img_h_cm = calculate_size_fit_static(
                    dummy_w_px,
                    dummy_h_px,
                    metrics.main_width,
                    metrics.main_height,
                    config.fit_mode,
                )
                this_gap_h = config.gap_h.to_cm(img_w_cm)

                min_x, min_y, max_x, max_y = calculate_item_bounds(
                    config,
                    metrics,
                    "dummy",
                    r,
                    c,
                    border_offset_cm,
                    override_size=(img_w_cm, img_h_cm),
                )
                item_w = max_x - min_x
                item_h = max_y - min_y

                if config.layout_mode == "flow":
                    main_l = cx
                    main_t = cy + (current_row_h - item_h) / 2
                else:
                    main_l = cx + (metrics.main_width - img_w_cm) / 2
                    main_t = cy + (metrics.main_height - img_h_cm) / 2

                has_crops = should_apply_crop(r, c, config)

                # Draw main image placeholder (clickable)
                tag = f"cell_{r}_{c}"
                canvas.create_rectangle(
                    tx(main_l),
                    ty(main_t),
                    tx(main_l + img_w_cm),
                    ty(main_t + img_h_cm),
                    fill="#ddf",
                    outline="blue",
                    tags=(tag,),
                )
                canvas.tag_bind(
                    tag,
                    "<Button-1>",
                    lambda e, rr=r, cc=c: self._on_preview_cell_click(rr, cc),
                )

                # Draw crop placeholders
                if has_crops and config.crop_regions:
                    self._draw_crop_previews(
                        canvas,
                        config,
                        metrics,
                        tx,
                        ty,
                        main_l,
                        main_t,
                        img_w_cm,
                        img_h_cm,
                        border_offset_cm,
                        half_border,
                    )

                # Advance X position
                if config.layout_mode == "flow":
                    cx += item_w + this_gap_h
                else:
                    cx += (
                        metrics.col_widths[c] + this_gap_h
                        if c < len(metrics.col_widths)
                        else metrics.main_width + this_gap_h
                    )

            # Advance Y position
            cy += current_row_h + config.gap_v.to_cm(metrics.main_height)

    def _draw_crop_previews(
        self,
        canvas,
        config,
        metrics,
        tx,
        ty,
        main_l,
        main_t,
        img_w_cm,
        img_h_cm,
        border_offset_cm,
        half_border,
    ):
        """Draw crop region previews on the canvas."""
        disp = config.crop_display
        num_crops = len(config.crop_regions)

        actual_gap_mc = disp.main_crop_gap.to_cm(
            img_w_cm if disp.position == "right" else img_h_cm
        )
        actual_gap_cc = disp.crop_crop_gap.to_cm(
            img_w_cm if disp.position == "right" else img_h_cm
        )

        if config.show_zoom_border:
            actual_gap_mc += border_offset_cm
            actual_gap_cc += border_offset_cm

        for crop_idx, region in enumerate(config.crop_regions):
            # Calculate crop size
            if disp.scale is not None or disp.size is not None:
                if disp.size:
                    c_w = c_h = disp.size
                else:
                    c_w = (
                        img_w_cm * disp.scale
                        if disp.position == "right"
                        else img_h_cm * disp.scale
                    )
                    c_h = c_w
            else:
                if disp.position == "right":
                    single_h = (img_h_cm - actual_gap_cc * (num_crops - 1)) / num_crops
                    c_w, c_h = calculate_size_fit_static(
                        100, 100, metrics.crop_size, single_h, "fit"
                    )
                else:
                    single_w = (img_w_cm - actual_gap_cc * (num_crops - 1)) / num_crops
                    c_w, c_h = calculate_size_fit_static(
                        100, 100, single_w, metrics.crop_size, "fit"
                    )

            # Calculate position
            this_gap_mc = region.gap if region.gap is not None else actual_gap_mc
            if region.gap is not None and config.show_zoom_border:
                this_gap_mc += border_offset_cm

            if disp.position == "right":
                c_l = main_l + img_w_cm + this_gap_mc

                if region.align == "start":
                    c_t = main_t + region.offset + half_border
                elif region.align == "center":
                    c_t = main_t + (img_h_cm - c_h) / 2 + region.offset
                elif region.align == "end":
                    c_t = main_t + img_h_cm - c_h + region.offset - half_border
                else:  # auto
                    if disp.scale is not None or disp.size is not None:
                        c_t = main_t + crop_idx * (c_h + actual_gap_cc)
                    else:
                        if crop_idx == 0:
                            c_t = main_t
                        elif num_crops > 1 and crop_idx == num_crops - 1:
                            c_t = (main_t + img_h_cm) - c_h
                        else:
                            single_h = (
                                img_h_cm - actual_gap_cc * (num_crops - 1)
                            ) / num_crops
                            slot_top = main_t + crop_idx * (single_h + actual_gap_cc)
                            c_t = slot_top + (single_h - c_h) / 2
            else:  # bottom
                c_t = main_t + img_h_cm + this_gap_mc

                if region.align == "start":
                    c_l = main_l + region.offset + half_border
                elif region.align == "center":
                    c_l = main_l + (img_w_cm - c_w) / 2 + region.offset
                elif region.align == "end":
                    c_l = main_l + img_w_cm - c_w + region.offset - half_border
                else:  # auto
                    if disp.scale is not None or disp.size is not None:
                        c_l = main_l + crop_idx * (c_w + actual_gap_cc)
                    else:
                        if crop_idx == 0:
                            c_l = main_l
                        elif num_crops > 1 and crop_idx == num_crops - 1:
                            c_l = (main_l + img_w_cm) - c_w
                        else:
                            single_w = (
                                img_w_cm - actual_gap_cc * (num_crops - 1)
                            ) / num_crops
                            slot_left = main_l + crop_idx * (single_w + actual_gap_cc)
                            c_l = slot_left + (single_w - c_w) / 2

            # Draw crop placeholder
            canvas.create_rectangle(
                tx(c_l),
                ty(c_t),
                tx(c_l + c_w),
                ty(c_t + c_h),
                fill="#fdd",
                outline="red",
            )


# =============================================================================
# Main Entry Point
# =============================================================================


def main():
    """Main entry point for GUI application."""
    root = tk.Tk()

    # Check for initial config argument
    initial_config = None
    if len(sys.argv) > 1:
        initial_config = sys.argv[1]

    app = ImageGridApp(root, initial_config)
    root.mainloop()


if __name__ == "__main__":
    main()
