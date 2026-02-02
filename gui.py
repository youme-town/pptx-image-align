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
import copy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from typing import Optional, List, Dict, Tuple
from dataclasses import dataclass, field

from PIL import Image, ImageTk

from core import (
    GridConfig,
    GapConfig,
    CropRegion,
    CropPreset,
    LabelConfig,
    ConnectorConfig,
    load_config,
    save_config,
    get_sorted_images,
    create_grid_presentation,
    calculate_grid_metrics,
    calculate_item_bounds,
    calculate_size_fit_static,
    normalize_slide_size,
    should_apply_crop,
    pt_to_cm,
    load_crop_presets,
    save_crop_preset,
)


# =============================================================================
# Undo/Redo Support
# =============================================================================


@dataclass
class AppState:
    """アプリケーション状態のスナップショット"""

    folders: List[str] = field(default_factory=list)
    images: List[str] = field(default_factory=list)
    crop_regions: List[CropRegion] = field(default_factory=list)
    input_mode: str = "folders"
    rows: int = 3
    cols: int = 3


class UndoRedoManager:
    """Undo/Redo履歴管理"""

    def __init__(self, max_history: int = 50):
        self.history: List[AppState] = []
        self.redo_stack: List[AppState] = []
        self.max_history = max_history
        self.is_restoring = False

    def push(self, state: AppState) -> None:
        """状態を履歴に追加"""
        if self.is_restoring:
            return
        self.history.append(state)
        if len(self.history) > self.max_history:
            self.history.pop(0)
        self.redo_stack.clear()

    def undo(self) -> Optional[AppState]:
        """1つ前の状態を返す"""
        if len(self.history) < 2:
            return None
        current = self.history.pop()
        self.redo_stack.append(current)
        return self.history[-1] if self.history else None

    def redo(self) -> Optional[AppState]:
        """やり直し状態を返す"""
        if not self.redo_stack:
            return None
        state = self.redo_stack.pop()
        self.history.append(state)
        return state

    def can_undo(self) -> bool:
        return len(self.history) >= 2

    def can_redo(self) -> bool:
        return len(self.redo_stack) > 0


# =============================================================================
# Crop Editor Window
# =============================================================================


class CropEditor(tk.Toplevel):
    """Window for visually selecting crop regions on an image."""

    def __init__(self, parent, image_path: str, callback, existing_regions=None):
        super().__init__(parent)
        self.title("Crop Editor")
        self.geometry("1100x700")
        self.callback = callback
        self.existing_regions = existing_regions or []

        self.image_path = image_path
        self.orig_img = Image.open(image_path)
        self.orig_w, self.orig_h = self.orig_img.size

        # UI Components - Main container with left (image) and right (preview)
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Left: Image canvas
        self.canvas_frame = ttk.Frame(self.main_frame)
        self.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Right: Preview panel
        self.preview_frame = ttk.LabelFrame(
            self.main_frame, text="クロップ領域プレビュー", padding=5
        )
        self.preview_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

        self.preview_canvas = tk.Canvas(
            self.preview_frame, bg="white", width=200, height=200
        )
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)

        self.preview_label = ttk.Label(
            self.preview_frame, text="領域を選択してください", wraplength=190
        )
        self.preview_label.pack(pady=5)

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

        # Preview image reference
        self.preview_tk_img = None

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

        # Draw existing crop regions (from the region list)
        for region in self.existing_regions:
            # Calculate region position based on mode
            if region.mode == "ratio":
                rx = (region.x_ratio or 0) * self.orig_w
                ry = (region.y_ratio or 0) * self.orig_h
                rw = (region.width_ratio or 0) * self.orig_w
                rh = (region.height_ratio or 0) * self.orig_h
            else:
                # px mode - treat x/y/width/height as actual pixels (core.resolve_crop_box と同じ解釈)
                rx = region.x
                ry = region.y
                rw = region.width
                rh = region.height

            # Convert to display coordinates
            sx = rx * self.display_scale + self.off_x
            sy = ry * self.display_scale + self.off_y
            ex = (rx + rw) * self.display_scale + self.off_x
            ey = (ry + rh) * self.display_scale + self.off_y

            # Draw with region color
            color = "#%02x%02x%02x" % region.color
            self.canvas.create_rectangle(sx, sy, ex, ey, outline=color, width=1)

            # Draw region name
            if region.name:
                self.canvas.create_text(
                    (sx + ex) / 2,
                    (sy + ey) / 2,
                    text=region.name,
                    fill=color,
                    font=("Arial", 9, "bold"),
                )

        # Redraw current selection rect if present
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
        if self.rect_id is None or self.start_x is None or self.start_y is None:
            return
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

        # リアルタイムでプレビューを更新（負荷軽減のため一時的な矩形を計算）
        x1 = min(self.start_x, event.x) - self.off_x
        y1 = min(self.start_y, event.y) - self.off_y
        x2 = max(self.start_x, event.x) - self.off_x
        y2 = max(self.start_y, event.y) - self.off_y

        ox1 = max(0, int(x1 / self.display_scale))
        oy1 = max(0, int(y1 / self.display_scale))
        ox2 = min(self.orig_w, int(x2 / self.display_scale))
        oy2 = min(self.orig_h, int(y2 / self.display_scale))

        w = ox2 - ox1
        h = oy2 - oy1

        if w > 5 and h > 5:
            self.cur_rect = (ox1, oy1, w, h)
            self._update_crop_preview()

    def on_release(self, event):
        if self.start_x is None or self.start_y is None:
            return
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
            self._update_crop_preview()

    def _update_crop_preview(self):
        """選択したクロップ領域のプレビューを更新"""
        if not self.cur_rect:
            return

        x, y, w, h = self.cur_rect

        # クロップ領域を切り出し
        cropped = self.orig_img.crop((x, y, x + w, y + h))

        # プレビューキャンバスのサイズに合わせてリサイズ
        preview_w = self.preview_canvas.winfo_width()
        preview_h = self.preview_canvas.winfo_height()

        if preview_w < 10 or preview_h < 10:
            preview_w, preview_h = 200, 200

        # アスペクト比を維持してリサイズ
        scale = min(preview_w / w, preview_h / h) * 0.95
        new_w = max(1, int(w * scale))
        new_h = max(1, int(h * scale))

        resized = cropped.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self.preview_tk_img = ImageTk.PhotoImage(resized)

        # プレビューを描画
        self.preview_canvas.delete("all")
        cx = preview_w // 2
        cy = preview_h // 2
        self.preview_canvas.create_image(
            cx, cy, anchor=tk.CENTER, image=self.preview_tk_img
        )

        # 情報ラベルを更新
        self.preview_label.config(
            text=f"サイズ: {w} x {h} px\n"
            f"位置: ({x}, {y})\n"
            f"比率: {w / self.orig_w:.1%} x {h / self.orig_h:.1%}"
        )

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

        # Undo/Redo initialization
        self.undo_manager = UndoRedoManager()
        self._save_state()  # 初期状態を保存
        self._bind_keyboard_shortcuts()

    def _bind_keyboard_shortcuts(self):
        """キーボードショートカットをバインド"""
        self.root.bind("<Control-z>", lambda e: self._undo())
        self.root.bind("<Control-y>", lambda e: self._redo())
        self.root.bind("<Control-Z>", lambda e: self._redo())  # Ctrl+Shift+Z

    def _get_current_state(self) -> AppState:
        """現在の状態をスナップショットとして取得"""
        return AppState(
            folders=copy.deepcopy(self.folders),
            images=copy.deepcopy(self.images),
            crop_regions=copy.deepcopy(self.crop_regions),
            input_mode=self.input_mode.get(),
            rows=self.rows.get(),
            cols=self.cols.get(),
        )

    def _restore_state(self, state: AppState) -> None:
        """状態を復元"""
        self.undo_manager.is_restoring = True
        try:
            self.folders = copy.deepcopy(state.folders)
            self.images = copy.deepcopy(state.images)
            self.crop_regions = copy.deepcopy(state.crop_regions)
            self.input_mode.set(state.input_mode)
            self.rows.set(state.rows)
            self.cols.set(state.cols)
            self._refresh_input_listbox()
            self._update_region_list()
            self._schedule_preview()
        finally:
            self.undo_manager.is_restoring = False

    def _save_state(self):
        """現在の状態を履歴に保存"""
        self.undo_manager.push(self._get_current_state())

    def _undo(self):
        """Undo操作"""
        state = self.undo_manager.undo()
        if state:
            self._restore_state(state)

    def _redo(self):
        """Redo操作"""
        state = self.undo_manager.redo()
        if state:
            self._restore_state(state)

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
        self.flow_axis = tk.StringVar(value="both")

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
        self.crop_border_dash = tk.StringVar(value="solid")
        self.show_zoom_border = tk.BooleanVar(value=True)
        self.zoom_border_w = tk.DoubleVar(value=1.5)
        self.zoom_border_shape = tk.StringVar(value="rectangle")
        self.zoom_border_dash = tk.StringVar(value="solid")

        # Label settings
        self.label_enabled = tk.BooleanVar(value=False)
        self.label_mode = tk.StringVar(value="filename")  # filename, number, custom
        self.label_position = tk.StringVar(value="bottom")  # top, bottom
        self.label_font_name = tk.StringVar(value="Arial")
        self.label_font_size = tk.DoubleVar(value=10.0)
        self.label_font_color = (0, 0, 0)
        self.label_font_bold = tk.BooleanVar(value=False)
        self.label_number_format = tk.StringVar(value="({n})")
        self.label_custom_texts = tk.StringVar(value="")  # comma-separated
        self.label_gap = tk.DoubleVar(value=0.1)

        # Template settings
        self.template_path = tk.StringVar(value="")
        self.slide_layout_index = tk.IntVar(value=6)

        # Connector settings
        self.connector_show = tk.BooleanVar(value=False)
        self.connector_width = tk.DoubleVar(value=1.0)
        self.connector_color = None  # None means use crop region color
        self.connector_style = tk.StringVar(value="straight")
        self.connector_dash_style = tk.StringVar(value="solid")

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
        self._loading_region = False  # Flag to prevent auto-update during load
        self.r_name = tk.StringVar()
        self.r_x = tk.IntVar()
        self.r_y = tk.IntVar()
        self.r_w = tk.IntVar()
        self.r_h = tk.IntVar()
        self.r_color = (255, 0, 0)
        self.r_align = tk.StringVar(value="auto")
        self.r_offset = tk.DoubleVar(value=0.0)
        self.r_gap = tk.StringVar(value="")
        self.r_show_zoomed = tk.BooleanVar(value=True)

        # 領域ごとの枠線設定 (空=グローバル設定を使用)
        self.r_show_crop_border = tk.StringVar(value="")  # "", "true", "false"
        self.r_crop_border_width = tk.StringVar(value="")
        self.r_crop_border_shape = tk.StringVar(value="")  # "", "rectangle", "rounded"
        self.r_crop_border_dash = tk.StringVar(value="")  # "", "solid", "dash", "dot", "dash_dot"
        self.r_show_zoom_border = tk.StringVar(value="")
        self.r_zoom_border_width = tk.StringVar(value="")
        self.r_zoom_border_shape = tk.StringVar(value="")
        self.r_zoom_border_dash = tk.StringVar(value="")

        # Cache for image dimensions used by preview rendering
        self._image_dim_cache: Dict[str, Tuple[int, int]] = {}

        # Add trace callbacks for auto-update
        for var in [
            self.r_name,
            self.r_x,
            self.r_y,
            self.r_w,
            self.r_h,
            self.r_align,
            self.r_offset,
            self.r_gap,
            self.r_show_zoomed,
            # 領域ごとの枠線設定
            self.r_show_crop_border,
            self.r_crop_border_width,
            self.r_crop_border_shape,
            self.r_crop_border_dash,
            self.r_show_zoom_border,
            self.r_zoom_border_width,
            self.r_zoom_border_shape,
            self.r_zoom_border_dash,
        ]:
            var.trace_add("write", self._on_region_detail_change)

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
            text="プレビュー",
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
        # 並び替えボタン
        ttk.Button(btn_frame, text="↑上へ", command=self._move_input_up).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="↓下へ", command=self._move_input_down).pack(
            side=tk.LEFT, padx=2
        )
        # 追加機能ボタン
        ttk.Button(btn_frame, text="複製", command=self._duplicate_input).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="空セル", command=self._add_placeholder).pack(
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

        ttk.Label(f_mode, text="| Axis:").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Combobox(
            f_mode,
            textvariable=self.flow_axis,
            values=["both", "horizontal", "vertical"],
            width=10,
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

        # Template settings
        f_tpl = ttk.LabelFrame(
            self.tab_layout, text="テンプレートPPTX (オプション)", padding=5
        )
        f_tpl.pack(fill=tk.X, pady=5)

        f_tpl_row1 = ttk.Frame(f_tpl)
        f_tpl_row1.pack(fill=tk.X)
        ttk.Label(f_tpl_row1, text="テンプレート:").pack(side=tk.LEFT)
        ttk.Entry(f_tpl_row1, textvariable=self.template_path, width=30).pack(
            side=tk.LEFT, padx=5, fill=tk.X, expand=True
        )
        ttk.Button(f_tpl_row1, text="参照", command=self._browse_template).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            f_tpl_row1, text="クリア", command=lambda: self.template_path.set("")
        ).pack(side=tk.LEFT, padx=2)

        f_tpl_row2 = ttk.Frame(f_tpl)
        f_tpl_row2.pack(fill=tk.X, pady=2)
        ttk.Label(f_tpl_row2, text="レイアウト番号:").pack(side=tk.LEFT)
        ttk.Entry(f_tpl_row2, textvariable=self.slide_layout_index, width=5).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Label(f_tpl_row2, text="(0=タイトル, 6=白紙)").pack(side=tk.LEFT)

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
        # Preset section
        f_preset = ttk.LabelFrame(self.tab_crop, text="クロッププリセット", padding=5)
        f_preset.pack(fill=tk.X, pady=5)

        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(
            f_preset, textvariable=self.preset_var, state="readonly", width=20
        )
        self.preset_combo.pack(side=tk.LEFT, padx=5)

        ttk.Button(f_preset, text="読込", command=self._load_preset).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            f_preset, text="現在の設定を保存", command=self._save_preset_dialog
        ).pack(side=tk.LEFT, padx=2)

        self._refresh_preset_list()

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
        ttk.Checkbutton(f_r3, text="拡大表示", variable=self.r_show_zoomed).pack(
            side=tk.LEFT, padx=10
        )

        # 領域ごとの枠線設定 (個別上書き)
        f_border = ttk.LabelFrame(f_detail, text="枠線設定 (空=グローバル)", padding=3)
        f_border.pack(fill=tk.X, pady=3)

        # クロップ枠線（元画像上）
        f_cb = ttk.Frame(f_border)
        f_cb.pack(fill=tk.X)
        ttk.Label(f_cb, text="元画像:", width=6).pack(side=tk.LEFT)
        ttk.Label(f_cb, text="表示").pack(side=tk.LEFT)
        cb_show = ttk.Combobox(
            f_cb, textvariable=self.r_show_crop_border, values=["", "true", "false"], width=5
        )
        cb_show.pack(side=tk.LEFT, padx=2)
        ttk.Label(f_cb, text="太さ").pack(side=tk.LEFT)
        ttk.Entry(f_cb, textvariable=self.r_crop_border_width, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(f_cb, text="形状").pack(side=tk.LEFT)
        ttk.Combobox(
            f_cb, textvariable=self.r_crop_border_shape,
            values=["", "rectangle", "rounded"], width=8
        ).pack(side=tk.LEFT, padx=2)
        ttk.Label(f_cb, text="線種").pack(side=tk.LEFT)
        ttk.Combobox(
            f_cb, textvariable=self.r_crop_border_dash,
            values=["", "solid", "dash", "dot", "dash_dot"], width=7
        ).pack(side=tk.LEFT, padx=2)

        # ズーム枠線（拡大画像）
        f_zb = ttk.Frame(f_border)
        f_zb.pack(fill=tk.X, pady=2)
        ttk.Label(f_zb, text="拡大:", width=6).pack(side=tk.LEFT)
        ttk.Label(f_zb, text="表示").pack(side=tk.LEFT)
        ttk.Combobox(
            f_zb, textvariable=self.r_show_zoom_border, values=["", "true", "false"], width=5
        ).pack(side=tk.LEFT, padx=2)
        ttk.Label(f_zb, text="太さ").pack(side=tk.LEFT)
        ttk.Entry(f_zb, textvariable=self.r_zoom_border_width, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(f_zb, text="形状").pack(side=tk.LEFT)
        ttk.Combobox(
            f_zb, textvariable=self.r_zoom_border_shape,
            values=["", "rectangle", "rounded"], width=8
        ).pack(side=tk.LEFT, padx=2)
        ttk.Label(f_zb, text="線種").pack(side=tk.LEFT)
        ttk.Combobox(
            f_zb, textvariable=self.r_zoom_border_dash,
            values=["", "solid", "dash", "dot", "dash_dot"], width=7
        ).pack(side=tk.LEFT, padx=2)

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

        f_src_row3 = ttk.Frame(f_src)
        f_src_row3.pack(fill=tk.X, pady=2)
        ttk.Label(f_src_row3, text="線種:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_src_row3, text="実線", variable=self.crop_border_dash, value="solid"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_src_row3, text="破線", variable=self.crop_border_dash, value="dash"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_src_row3, text="点線", variable=self.crop_border_dash, value="dot"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_src_row3, text="一点鎖線", variable=self.crop_border_dash, value="dash_dot"
        ).pack(side=tk.LEFT, padx=3)

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

        f_zoom_row3 = ttk.Frame(f_zoom)
        f_zoom_row3.pack(fill=tk.X, pady=2)
        ttk.Label(f_zoom_row3, text="線種:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_zoom_row3, text="実線", variable=self.zoom_border_dash, value="solid"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_zoom_row3, text="破線", variable=self.zoom_border_dash, value="dash"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_zoom_row3, text="点線", variable=self.zoom_border_dash, value="dot"
        ).pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(
            f_zoom_row3, text="一点鎖線", variable=self.zoom_border_dash, value="dash_dot"
        ).pack(side=tk.LEFT, padx=3)

        # Label settings
        f_label = ttk.LabelFrame(
            self.tab_style, text="テキストラベル/キャプション", padding=5
        )
        f_label.pack(fill=tk.X, pady=5)

        # Enable/disable
        f_lbl_row1 = ttk.Frame(f_label)
        f_lbl_row1.pack(fill=tk.X)
        ttk.Checkbutton(
            f_lbl_row1, text="ラベルを表示", variable=self.label_enabled
        ).pack(side=tk.LEFT)

        # Mode selection
        f_lbl_row2 = ttk.Frame(f_label)
        f_lbl_row2.pack(fill=tk.X, pady=2)
        ttk.Label(f_lbl_row2, text="モード:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_lbl_row2, text="ファイル名", variable=self.label_mode, value="filename"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_lbl_row2, text="連番", variable=self.label_mode, value="number"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_lbl_row2, text="カスタム", variable=self.label_mode, value="custom"
        ).pack(side=tk.LEFT, padx=5)

        # Position
        f_lbl_row3 = ttk.Frame(f_label)
        f_lbl_row3.pack(fill=tk.X, pady=2)
        ttk.Label(f_lbl_row3, text="位置:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_lbl_row3, text="上", variable=self.label_position, value="top"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_lbl_row3, text="下", variable=self.label_position, value="bottom"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Label(f_lbl_row3, text="間隔(cm):").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(f_lbl_row3, textvariable=self.label_gap, width=5).pack(side=tk.LEFT)

        # Font settings
        f_lbl_row4 = ttk.Frame(f_label)
        f_lbl_row4.pack(fill=tk.X, pady=2)
        ttk.Label(f_lbl_row4, text="フォント:").pack(side=tk.LEFT)
        ttk.Entry(f_lbl_row4, textvariable=self.label_font_name, width=12).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Label(f_lbl_row4, text="サイズ(pt):").pack(side=tk.LEFT, padx=(5, 2))
        ttk.Entry(f_lbl_row4, textvariable=self.label_font_size, width=5).pack(
            side=tk.LEFT
        )
        ttk.Checkbutton(f_lbl_row4, text="太字", variable=self.label_font_bold).pack(
            side=tk.LEFT, padx=5
        )
        self.btn_label_color = tk.Button(
            f_lbl_row4, text="色", width=4, command=self._pick_label_color
        )
        self.btn_label_color.pack(side=tk.LEFT, padx=5)

        # Number format
        f_lbl_row5 = ttk.Frame(f_label)
        f_lbl_row5.pack(fill=tk.X, pady=2)
        ttk.Label(f_lbl_row5, text="連番フォーマット:").pack(side=tk.LEFT)
        ttk.Entry(f_lbl_row5, textvariable=self.label_number_format, width=10).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Label(f_lbl_row5, text="({n}で番号)").pack(side=tk.LEFT)

        # Custom texts
        f_lbl_row6 = ttk.Frame(f_label)
        f_lbl_row6.pack(fill=tk.X, pady=2)
        ttk.Label(f_lbl_row6, text="カスタムテキスト (カンマ区切り):").pack(
            side=tk.LEFT
        )
        ttk.Entry(f_lbl_row6, textvariable=self.label_custom_texts, width=30).pack(
            side=tk.LEFT, padx=2, fill=tk.X, expand=True
        )

        # Connector settings
        f_conn = ttk.LabelFrame(self.tab_style, text="クロップ連結線", padding=5)
        f_conn.pack(fill=tk.X, pady=5)

        # Enable/disable
        f_conn_row1 = ttk.Frame(f_conn)
        f_conn_row1.pack(fill=tk.X)
        ttk.Checkbutton(
            f_conn_row1, text="連結線を表示", variable=self.connector_show
        ).pack(side=tk.LEFT)
        ttk.Label(f_conn_row1, text="太さ(pt):").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(f_conn_row1, textvariable=self.connector_width, width=5).pack(
            side=tk.LEFT
        )

        # Style options
        f_conn_row2 = ttk.Frame(f_conn)
        f_conn_row2.pack(fill=tk.X, pady=2)
        ttk.Label(f_conn_row2, text="線種:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_conn_row2, text="実線", variable=self.connector_dash_style, value="solid"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_conn_row2, text="破線", variable=self.connector_dash_style, value="dash"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_conn_row2, text="点線", variable=self.connector_dash_style, value="dot"
        ).pack(side=tk.LEFT, padx=5)

    def _pick_label_color(self):
        """ラベルの色を選択"""
        init_color = f"#{self.label_font_color[0]:02x}{self.label_font_color[1]:02x}{self.label_font_color[2]:02x}"
        color = colorchooser.askcolor(initialcolor=init_color)
        if color[0]:
            self.label_font_color = (
                int(color[0][0]),
                int(color[0][1]),
                int(color[0][2]),
            )
            self._schedule_preview()

    def _browse_template(self):
        """テンプレートPPTXファイルを選択"""
        path = filedialog.askopenfilename(
            filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")]
        )
        if path:
            self.template_path.set(path)

    def _add_preview_tracers(self):
        """Add trace callbacks for automatic preview updates."""
        vars_to_trace = [
            self.rows,
            self.cols,
            self.arrangement,
            self.layout_mode,
            self.flow_align,
            self.flow_vertical_align,
            self.flow_axis,
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
            self.zoom_border_dash,
            self.show_zoom_border,
            self.zoom_border_w,
            self.crop_border_shape,
            self.crop_border_dash,
            self.show_crop_border,
            self.crop_border_w,
            # Label settings
            self.label_enabled,
            self.label_mode,
            self.label_position,
            self.label_font_name,
            self.label_font_size,
            self.label_font_bold,
            self.label_number_format,
            self.label_custom_texts,
            self.label_gap,
            # Connector settings
            self.connector_show,
            self.connector_width,
            self.connector_style,
            self.connector_dash_style,
        ]
        for v in vars_to_trace:
            v.trace_add("write", lambda *args: self._schedule_preview())

    def _update_region_list(self):
        """region_tree を更新（選択状態を保持）"""
        # Remember current selection
        current_selection = self.region_tree.selection()

        for item in self.region_tree.get_children():
            self.region_tree.delete(item)
        for i, r in enumerate(self.crop_regions):
            # 座標表示
            coord = f"({r.x}, {r.y}, {r.width}, {r.height})"
            if r.mode == "ratio":
                coord += " [比率]"
            if not r.show_zoomed:
                coord += " [枠のみ]"
            self.region_tree.insert(
                "", tk.END, iid=str(i), values=(r.name, coord, r.align)
            )

        # Restore selection if still valid
        for sel_id in current_selection:
            if sel_id in self.region_tree.get_children():
                self.region_tree.selection_add(sel_id)

    def _add_region_dialog(self):
        """数値入力でクロップ領域を追加"""
        from tkinter import simpledialog

        name = simpledialog.askstring(
            "領域追加", "領域名:", initialvalue=f"Region_{len(self.crop_regions) + 1}"
        )
        if not name:
            return
        region = CropRegion(
            x=0, y=0, width=100, height=100, color=(255, 0, 0), name=name
        )
        self.crop_regions.append(region)
        self._update_region_list()
        self._schedule_preview()
        self._save_state()

    def _remove_region(self):
        """選択したクロップ領域を削除"""
        sel = self.region_tree.selection()
        if not sel:
            return
        indices = sorted([int(s) for s in sel], reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.crop_regions):
                self.crop_regions.pop(idx)
        self._update_region_list()
        self._schedule_preview()
        self._save_state()

    def _on_region_select(self, event):
        """領域選択時に詳細を表示"""
        sel = self.region_tree.selection()
        if not sel:
            self.sel_idx = None
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.crop_regions):
            self.sel_idx = idx
            r = self.crop_regions[idx]
            # Prevent auto-update while loading values
            self._loading_region = True
            try:
                self.r_name.set(r.name)
                self.r_x.set(r.x)
                self.r_y.set(r.y)
                self.r_w.set(r.width)
                self.r_h.set(r.height)
                self.r_align.set(r.align)
                self.r_offset.set(r.offset)
                self.r_gap.set(str(r.gap) if r.gap is not None else "")
                self.r_color = r.color
                self.r_show_zoomed.set(r.show_zoomed)

                # 領域ごとの枠線設定を読み込み
                self.r_show_crop_border.set(
                    "true" if r.show_crop_border is True else "false" if r.show_crop_border is False else ""
                )
                self.r_crop_border_width.set(str(r.crop_border_width) if r.crop_border_width is not None else "")
                self.r_crop_border_shape.set(r.crop_border_shape or "")
                self.r_crop_border_dash.set(r.crop_border_dash or "")
                self.r_show_zoom_border.set(
                    "true" if r.show_zoom_border is True else "false" if r.show_zoom_border is False else ""
                )
                self.r_zoom_border_width.set(str(r.zoom_border_width) if r.zoom_border_width is not None else "")
                self.r_zoom_border_shape.set(r.zoom_border_shape or "")
                self.r_zoom_border_dash.set(r.zoom_border_dash or "")
            finally:
                self._loading_region = False

    def _pick_region_color(self):
        """領域の色を選択"""
        # 現在の色を16進数文字列に変換
        init_color = f"#{self.r_color[0]:02x}{self.r_color[1]:02x}{self.r_color[2]:02x}"
        color = colorchooser.askcolor(initialcolor=init_color)
        if color[0]:
            self.r_color = (int(color[0][0]), int(color[0][1]), int(color[0][2]))
            if self.sel_idx is not None and 0 <= self.sel_idx < len(self.crop_regions):
                self.crop_regions[self.sel_idx].color = self.r_color
                self._schedule_preview()

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

    def _get_dash_pattern(self, dash_style: str):
        """Convert dash style string to Tkinter dash pattern."""
        if dash_style == "dash":
            return (6, 3)
        elif dash_style == "dot":
            return (2, 3)
        elif dash_style == "dash_dot":
            return (6, 3, 2, 3)
        elif dash_style == "long_dash":
            return (10, 5)
        elif dash_style == "long_dash_dot":
            return (10, 5, 2, 5)
        else:
            return ()  # solid

    def _get_current_config(self) -> GridConfig:
        """Build GridConfig from current GUI state."""
        c = GridConfig()

        c.layout_mode = self.layout_mode.get()
        c.flow_align = self.flow_align.get()
        c.flow_vertical_align = self.flow_vertical_align.get()
        c.flow_axis = self.flow_axis.get()
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
        c.crop_border_dash = self.crop_border_dash.get()
        c.zoom_border_dash = self.zoom_border_dash.get()
        c.show_crop_border = self.show_crop_border.get()
        c.crop_border_width = self._get_safe_double(self.crop_border_w, 1.5)
        c.show_zoom_border = self.show_zoom_border.get()
        c.zoom_border_width = self._get_safe_double(self.zoom_border_w, 1.5)

        # Label settings
        c.label_config = LabelConfig(
            enabled=self.label_enabled.get(),
            mode=self.label_mode.get(),
            position=self.label_position.get(),
            font_name=self.label_font_name.get(),
            font_size=self._get_safe_double(self.label_font_size, 10.0),
            font_color=self.label_font_color,
            font_bold=self.label_font_bold.get(),
            number_format=self.label_number_format.get(),
            custom_texts=[
                t.strip() for t in self.label_custom_texts.get().split(",") if t.strip()
            ],
            gap=self._get_safe_double(self.label_gap, 0.1),
        )

        # Template settings
        tpl_path = self.template_path.get().strip()
        c.template_path = tpl_path if tpl_path else None
        c.slide_layout_index = self._get_safe_int(self.slide_layout_index, 6)

        # Connector settings
        c.connector = ConnectorConfig(
            show=self.connector_show.get(),
            width=self._get_safe_double(self.connector_width, 1.0),
            color=self.connector_color,
            style=self.connector_style.get(),
            dash_style=self.connector_dash_style.get(),
        )

        return c

    def _apply_config_to_gui(self, c: GridConfig):
        """Apply a GridConfig to the GUI state."""
        self.layout_mode.set(c.layout_mode)
        self.flow_align.set(c.flow_align)
        self.flow_vertical_align.set(c.flow_vertical_align)
        self.flow_axis.set(c.flow_axis)
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
        self.crop_border_dash.set(c.crop_border_dash)
        self.zoom_border_dash.set(c.zoom_border_dash)
        self.show_crop_border.set(c.show_crop_border)
        self.crop_border_w.set(c.crop_border_width)
        self.show_zoom_border.set(c.show_zoom_border)
        self.zoom_border_w.set(c.zoom_border_width)

        # Label settings
        self.label_enabled.set(c.label_config.enabled)
        self.label_mode.set(c.label_config.mode)
        self.label_position.set(c.label_config.position)
        self.label_font_name.set(c.label_config.font_name)
        self.label_font_size.set(c.label_config.font_size)
        self.label_font_color = c.label_config.font_color
        self.label_font_bold.set(c.label_config.font_bold)
        self.label_number_format.set(c.label_config.number_format)
        self.label_custom_texts.set(",".join(c.label_config.custom_texts))
        self.label_gap.set(c.label_config.gap)

        # Template settings
        self.template_path.set(c.template_path or "")
        self.slide_layout_index.set(c.slide_layout_index)

        # Connector settings
        self.connector_show.set(c.connector.show)
        self.connector_width.set(c.connector.width)
        self.connector_color = c.connector.color
        self.connector_style.set(c.connector.style)
        self.connector_dash_style.set(c.connector.dash_style)

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
            for i, p in enumerate(self.images):
                if p == "__PLACEHOLDER__":
                    self.folder_listbox.insert(tk.END, f"[{i + 1}] (空セル)")
                else:
                    self.folder_listbox.insert(tk.END, f"[{i + 1}] {p}")
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
            self._save_state()  # Undo/Redo

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
            self._save_state()  # Undo/Redo

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
        self._save_state()  # Undo/Redo

    def _clear_inputs(self):
        self.folders = []
        self.images = []
        self._refresh_input_listbox()
        self._save_state()  # Undo/Redo

    def _move_input_up(self):
        """選択項目を1つ上に移動"""
        sel = self.folder_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        items = self.folders if self.input_mode.get() == "folders" else self.images
        if idx > 0 and idx < len(items):
            items[idx], items[idx - 1] = items[idx - 1], items[idx]
            self._refresh_input_listbox()
            self.folder_listbox.selection_set(idx - 1)
            self._schedule_preview()
            self._save_state()  # Undo/Redo

    def _move_input_down(self):
        """選択項目を1つ下に移動"""
        sel = self.folder_listbox.curselection()
        items = self.folders if self.input_mode.get() == "folders" else self.images
        if not sel or sel[0] >= len(items) - 1:
            return
        idx = sel[0]
        if idx >= 0 and idx < len(items) - 1:
            items[idx], items[idx + 1] = items[idx + 1], items[idx]
            self._refresh_input_listbox()
            self.folder_listbox.selection_set(idx + 1)
            self._schedule_preview()
            self._save_state()  # Undo/Redo

    def _duplicate_input(self):
        """選択した画像を複製"""
        if self.input_mode.get() != "images":
            messagebox.showinfo("情報", "画像リストモードでのみ使用できます")
            return
        sel = self.folder_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "複製する画像を選択してください")
            return
        idx = sel[0]
        if 0 <= idx < len(self.images):
            # 選択した画像を直後に挿入
            self.images.insert(idx + 1, self.images[idx])
            self._refresh_input_listbox()
            self.folder_listbox.selection_set(idx + 1)
            self._schedule_preview()
            self._save_state()

    def _add_placeholder(self):
        """空セル（プレースホルダー）を追加"""
        if self.input_mode.get() != "images":
            messagebox.showinfo("情報", "画像リストモードでのみ使用できます")
            return
        # 選択位置に挿入、なければ末尾に追加
        sel = self.folder_listbox.curselection()
        if sel:
            idx = sel[0] + 1
            self.images.insert(idx, "__PLACEHOLDER__")
            self._refresh_input_listbox()
            self.folder_listbox.selection_set(idx)
        else:
            self.images.append("__PLACEHOLDER__")
            self._refresh_input_listbox()
            self.folder_listbox.selection_set(len(self.images) - 1)
        self._schedule_preview()
        self._save_state()

    def _open_crop_editor(self):
        target_img = None

        # NEW: prioritize actual cell images (skip placeholders)
        if self.input_mode.get() == "images":
            for img in self.images:
                if img and img != "__PLACEHOLDER__":
                    target_img = img
                    break
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
            CropEditor(
                self.root, target_img, self._add_region_from_editor, self.crop_regions
            )

    def _get_image_path_for_cell(self, row: int, col: int) -> Optional[str]:
        config = self._get_current_config()
        if config.images:
            idx = row * config.cols + col
            if 0 <= idx < len(config.images):
                path = config.images[idx]
                # プレースホルダーは画像として扱わない
                if path == "__PLACEHOLDER__":
                    return None
                return path
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
            CropEditor(self.root, p, self._add_region_from_editor, self.crop_regions)

    def _add_region_from_editor(self, x: int, y: int, w: int, h: int, name: str):
        """CropEditorからのコールバック - 新しいクロップ領域を追加"""
        region = CropRegion(
            x=x,
            y=y,
            width=w,
            height=h,
            color=(255, 0, 0),
            name=name or f"Region_{len(self.crop_regions) + 1}",
        )
        self.crop_regions.append(region)
        self._update_region_list()
        self._schedule_preview()
        self._save_state()  # Undo/Redo

    def _update_preview(self):
        """Render the preview canvas."""
        config = self._get_current_config()
        normalize_slide_size(config)
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
        flow_mode = config.layout_mode == "flow"
        flow_h = flow_mode and config.flow_axis in ("both", "horizontal")
        flow_v = flow_mode and config.flow_axis in ("both", "vertical")

        def _cell_input_state(row: int, col: int) -> tuple[Optional[str], bool]:
            """Return (image_path, is_placeholder).

            image_path is None for empty cells and placeholders.
            """
            if config.images:
                idx = row * config.cols + col
                if 0 <= idx < len(config.images):
                    raw = config.images[idx]
                    if raw == "__PLACEHOLDER__":
                        return None, True
                    return raw, False
                return None, False

            # folder mode
            folders = config.folders or []
            if config.arrangement == "row":
                if row >= len(folders):
                    return None, False
                imgs = get_sorted_images(folders[row])
                return (imgs[col], False) if col < len(imgs) else (None, False)
            else:
                if col >= len(folders):
                    return None, False
                imgs = get_sorted_images(folders[col])
                return (imgs[row], False) if row < len(imgs) else (None, False)

        # Get dummy image ratio
        d_rw = max(0.01, self._get_safe_double(self.dummy_ratio_w, 1.0))
        d_rh = max(0.01, self._get_safe_double(self.dummy_ratio_h, 1.0))
        dummy_w_px = 400
        dummy_h_px = int(400 * (d_rh / d_rw))

        border_offset_cm = (
            pt_to_cm(config.zoom_border_width) if config.show_zoom_border else 0.0
        )
        half_border = border_offset_cm / 2.0 if config.show_zoom_border else 0.0

        # Pre-calculate flow layout
        total_content_height = 0.0
        total_content_width = 0.0
        row_heights_flow = []
        flow_col_widths = []

        if flow_mode:
            if flow_v:
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
            else:
                row_heights_flow = metrics.row_heights[:]

            if flow_h:
                for c in range(config.cols):
                    sim_col_w = 0.0
                    for r in range(config.rows):
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
                        item_w = max_x - min_x
                        sim_col_w = max(sim_col_w, item_w)

                    flow_col_widths.append(
                        sim_col_w if sim_col_w > 0 else metrics.main_width
                    )

                total_content_width = sum(flow_col_widths)
                if config.cols > 1:
                    total_content_width += (config.cols - 1) * config.gap_h.to_cm(
                        metrics.main_width
                    )
            else:
                flow_col_widths = metrics.col_widths[:]

        # Calculate starting positions
        cy = config.margin_top
        flow_start_x = config.margin_left
        if flow_mode:
            if flow_v:
                avail_h = config.slide_height - config.margin_top - config.margin_bottom
                if config.flow_vertical_align == "center":
                    cy = config.margin_top + (avail_h - total_content_height) / 2
                elif config.flow_vertical_align == "bottom":
                    cy = (config.margin_top + avail_h) - total_content_height

            if flow_h:
                avail_w = config.slide_width - config.margin_left - config.margin_right
                if config.flow_align == "center":
                    flow_start_x = (
                        config.margin_left + (avail_w - total_content_width) / 2
                    )
                elif config.flow_align == "right":
                    flow_start_x = config.margin_left + (avail_w - total_content_width)

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

            cx = flow_start_x if flow_mode else config.margin_left

            # Draw cells
            for c in range(config.cols):
                img_w_cm, img_h_cm = calculate_size_fit_static(
                    dummy_w_px,
                    dummy_h_px,
                    metrics.main_width,
                    metrics.main_height,
                    config.fit_mode,
                )
                this_gap_h = config.gap_h.to_cm(metrics.main_width)

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

                if flow_mode and flow_h:
                    cell_w = (
                        flow_col_widths[c]
                        if c < len(flow_col_widths)
                        else metrics.main_width
                    )
                    item_draw_left = cx + (cell_w - item_w) / 2
                    main_l = item_draw_left - min_x
                else:
                    # Non-flow horizontal axis keeps grid-aligned main images.
                    main_l = cx + (metrics.main_width - img_w_cm) / 2

                if flow_mode and flow_v:
                    item_draw_top = cy + (current_row_h - item_h) / 2
                    main_t = item_draw_top - min_y
                else:
                    # Non-flow vertical axis keeps grid-aligned main images.
                    main_t = cy + (metrics.main_height - img_h_cm) / 2

                has_crops = should_apply_crop(r, c, config)

                cell_img_path, is_placeholder = _cell_input_state(r, c)
                is_empty = cell_img_path is None

                # Draw main image placeholder (clickable)
                tag = f"cell_{r}_{c}"
                if is_empty:
                    canvas.create_rectangle(
                        tx(main_l),
                        ty(main_t),
                        tx(main_l + img_w_cm),
                        ty(main_t + img_h_cm),
                        fill="#f0f0f0",
                        outline="#888",
                        dash=(4, 2),
                        tags=(tag,),
                    )
                else:
                    canvas.create_rectangle(
                        tx(main_l),
                        ty(main_t),
                        tx(main_l + img_w_cm),
                        ty(main_t + img_h_cm),
                        fill="#ddf",
                        outline="blue",
                        tags=(tag,),
                    )

                if is_empty:
                    label = "EMPTY"
                    rect_px_w = img_w_cm * scale
                    rect_px_h = img_h_cm * scale
                    font_size = int(max(8, min(14, min(rect_px_w, rect_px_h) / 8)))
                    canvas.create_text(
                        tx(main_l + img_w_cm / 2),
                        ty(main_t + img_h_cm / 2),
                        text=label,
                        fill="#555",
                        font=("Segoe UI", font_size, "bold"),
                        tags=(tag,),
                    )
                canvas.tag_bind(
                    tag,
                    "<Button-1>",
                    lambda e, rr=r, cc=c: self._on_preview_cell_click(rr, cc),
                )

                # Draw crop placeholders
                if (not is_empty) and has_crops and config.crop_regions:
                    self._draw_crop_previews(
                        canvas,
                        config,
                        metrics,
                        tx,
                        ty,
                        r,
                        c,
                        main_l,
                        main_t,
                        img_w_cm,
                        img_h_cm,
                        border_offset_cm,
                        half_border,
                        scale,
                    )

                # Advance X position
                if flow_mode:
                    w = (
                        flow_col_widths[c]
                        if c < len(flow_col_widths)
                        else metrics.main_width
                    )
                    cx += w + this_gap_h
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
        row,
        col,
        main_l,
        main_t,
        img_w_cm,
        img_h_cm,
        border_offset_cm,
        half_border,
        scale,
    ):
        """Draw crop region previews on the canvas."""
        # ガード: img_w_cmやimg_h_cmが0の場合は描画をスキップ
        if img_w_cm <= 0 or img_h_cm <= 0 or scale <= 0:
            return
        disp = config.crop_display
        # 空クロップ（show_zoomed=False）はautoの位置計算から除外
        num_crops = sum(1 for r in config.crop_regions if r.show_zoomed)

        # If we have a real image for this cell, use its pixel size to convert px-mode regions -> ratio.
        img_w_px = None
        img_h_px = None
        image_path = self._get_image_path_for_cell(row, col)
        if image_path:
            try:
                if image_path in self._image_dim_cache:
                    img_w_px, img_h_px = self._image_dim_cache[image_path]
                else:
                    with Image.open(image_path) as _img:
                        img_w_px, img_h_px = _img.size
                    self._image_dim_cache[image_path] = (img_w_px, img_h_px)
            except Exception:
                img_w_px = img_h_px = None

        # ----------------------
        # Step 1: Draw crop source regions on main image
        # ----------------------
        for region in config.crop_regions:
            # 領域ごとの枠線設定（Noneの場合はグローバル設定を使用）
            show_cb = region.show_crop_border if region.show_crop_border is not None else config.show_crop_border
            if not show_cb:
                continue

            # Calculate crop region position (relative to main image)
            if region.mode == "ratio":
                # Use ratio-based positioning
                rx = region.x_ratio or 0.0
                ry = region.y_ratio or 0.0
                rw = region.width_ratio or 0.0
                rh = region.height_ratio or 0.0
            else:
                # px-mode: prefer converting using actual image pixel dimensions when available.
                if img_w_px and img_h_px and img_w_px > 0 and img_h_px > 0:
                    rx = region.x / float(img_w_px)
                    ry = region.y / float(img_h_px)
                    rw = region.width / float(img_w_px)
                    rh = region.height / float(img_h_px)
                else:
                    # Fallback (legacy preview behavior): treat values as 1000-based normalized coords.
                    rx = region.x / 1000.0
                    ry = region.y / 1000.0
                    rw = region.width / 1000.0
                    rh = region.height / 1000.0

            # Convert to canvas coordinates
            crop_l = main_l + rx * img_w_cm
            crop_t = main_t + ry * img_h_cm
            crop_w = rw * img_w_cm
            crop_h = rh * img_h_cm

            # 領域ごとの枠線設定
            dash_style = region.crop_border_dash if region.crop_border_dash else config.crop_border_dash
            dash_pattern = self._get_dash_pattern(dash_style)
            border_width = region.crop_border_width if region.crop_border_width is not None else config.crop_border_width
            # プレビュー用の太さ（破線の場合は最低2pxで見やすく）
            min_width = 2 if dash_pattern else 1
            preview_width = max(min_width, min(3, int(border_width + 0.5)))

            # Draw crop region rectangle on main image
            color = "#%02x%02x%02x" % region.color
            canvas.create_rectangle(
                tx(crop_l),
                ty(crop_t),
                tx(crop_l + crop_w),
                ty(crop_t + crop_h),
                fill="",
                outline=color,
                width=preview_width,
                dash=dash_pattern,
            )

            # Draw region name label
            if region.name:
                canvas.create_text(
                    tx(crop_l + crop_w / 2),
                    ty(crop_t + crop_h / 2),
                    text=region.name,
                    fill=color,
                    font=("Arial", 8, "bold"),
                )

        # ----------------------
        # Step 2: Draw cropped (zoomed) previews at destination position
        # ----------------------

        actual_gap_mc = disp.main_crop_gap.to_cm(
            img_w_cm if disp.position == "right" else img_h_cm
        )
        actual_gap_cc = disp.crop_crop_gap.to_cm(
            img_w_cm if disp.position == "right" else img_h_cm
        )

        if config.show_zoom_border:
            actual_gap_mc += border_offset_cm
            actual_gap_cc += border_offset_cm

        visible_crop_idx = 0
        for crop_idx, region in enumerate(config.crop_regions):
            # 空クロップの場合、拡大表示はスキップ
            if not region.show_zoomed:
                continue

            # Get crop region aspect ratio
            if region.mode == "ratio":
                crop_rw = region.width_ratio or 0.1
                crop_rh = region.height_ratio or 0.1
            else:
                crop_rw = region.width / 1000.0 if region.width > 0 else 0.1
                crop_rh = region.height / 1000.0 if region.height > 0 else 0.1
            # クロップ領域の実際のピクセル比率（画像のアスペクト比を考慮）
            crop_aspect = (
                (crop_rw * img_w_cm) / (crop_rh * img_h_cm) if crop_rh > 0 else 1.0
            )

            # Calculate crop size (maintaining aspect ratio)
            if disp.scale is not None or disp.size is not None:
                if disp.size:
                    # サイズ指定時はアスペクト比を維持
                    if crop_aspect >= 1.0:
                        c_w = disp.size
                        c_h = disp.size / crop_aspect
                    else:
                        c_h = disp.size
                        c_w = disp.size * crop_aspect
                else:
                    base = (
                        img_w_cm * disp.scale
                        if disp.position == "right"
                        else img_h_cm * disp.scale
                    )
                    if crop_aspect >= 1.0:
                        c_w = base
                        c_h = base / crop_aspect
                    else:
                        c_h = base
                        c_w = base * crop_aspect
            else:
                if disp.position == "right":
                    single_h = (img_h_cm - actual_gap_cc * (num_crops - 1)) / num_crops
                    c_w, c_h = calculate_size_fit_static(
                        int(crop_aspect * 100), 100, metrics.crop_size, single_h, "fit"
                    )
                else:
                    single_w = (img_w_cm - actual_gap_cc * (num_crops - 1)) / num_crops
                    c_w, c_h = calculate_size_fit_static(
                        int(crop_aspect * 100), 100, single_w, metrics.crop_size, "fit"
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
                    # 最初と最後のクロップは端に揃える（pin ends）
                    if visible_crop_idx == 0:
                        c_t = main_t
                    elif num_crops > 1 and visible_crop_idx == num_crops - 1:
                        c_t = (main_t + img_h_cm) - c_h
                    else:
                        # 中間のクロップは均等配置
                        if disp.scale is not None or disp.size is not None:
                            total_crop_h = (
                                num_crops * c_h + (num_crops - 1) * actual_gap_cc
                            )
                            start_y = main_t + (img_h_cm - total_crop_h) / 2
                            c_t = start_y + visible_crop_idx * (c_h + actual_gap_cc)
                        else:
                            single_h = (
                                img_h_cm - actual_gap_cc * (num_crops - 1)
                            ) / num_crops
                            slot_top = main_t + visible_crop_idx * (
                                single_h + actual_gap_cc
                            )
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
                    # 最初と最後のクロップは端に揃える（pin ends）
                    if visible_crop_idx == 0:
                        c_l = main_l
                    elif num_crops > 1 and visible_crop_idx == num_crops - 1:
                        c_l = (main_l + img_w_cm) - c_w
                    else:
                        # 中間のクロップは均等配置
                        if disp.scale is not None or disp.size is not None:
                            total_crop_w = (
                                num_crops * c_w + (num_crops - 1) * actual_gap_cc
                            )
                            start_x = main_l + (img_w_cm - total_crop_w) / 2
                            c_l = start_x + visible_crop_idx * (c_w + actual_gap_cc)
                        else:
                            single_w = (
                                img_w_cm - actual_gap_cc * (num_crops - 1)
                            ) / num_crops
                            slot_left = main_l + visible_crop_idx * (
                                single_w + actual_gap_cc
                            )
                            c_l = slot_left + (single_w - c_w) / 2

            # Draw crop placeholder (with matching color)
            color = "#%02x%02x%02x" % region.color
            # Lighten the fill color
            fill_r = min(255, region.color[0] + 200)
            fill_g = min(255, region.color[1] + 200)
            fill_b = min(255, region.color[2] + 200)
            fill_color = "#%02x%02x%02x" % (fill_r, fill_g, fill_b)

            # 領域ごとの枠線設定（Noneの場合はグローバル設定を使用）
            show_zb = region.show_zoom_border if region.show_zoom_border is not None else config.show_zoom_border
            dash_style = region.zoom_border_dash if region.zoom_border_dash else config.zoom_border_dash
            dash_pattern = self._get_dash_pattern(dash_style)
            zb_width = region.zoom_border_width if region.zoom_border_width is not None else config.zoom_border_width
            # プレビュー用の太さ（破線の場合は最低2pxで見やすく）
            min_width = 2 if dash_pattern else 1
            preview_zb_width = max(min_width, min(3, int(zb_width + 0.5)))

            canvas.create_rectangle(
                tx(c_l),
                ty(c_t),
                tx(c_l + c_w),
                ty(c_t + c_h),
                fill=fill_color,
                outline=color if show_zb else "",
                width=preview_zb_width if show_zb else 1,
                dash=dash_pattern if show_zb else (),
            )

            # Draw region name in zoomed preview
            if region.name:
                canvas.create_text(
                    tx(c_l + c_w / 2),
                    ty(c_t + c_h / 2),
                    text=region.name,
                    fill=color,
                    font=("Arial", 8, "bold"),
                )

            visible_crop_idx += 1

    # -------------------------------------------------------------------------
    # Preset Methods
    # -------------------------------------------------------------------------

    def _refresh_preset_list(self):
        """プリセットリストを更新"""
        presets = load_crop_presets()
        preset_names = [p.name for p in presets]
        self.preset_combo["values"] = preset_names
        if preset_names:
            self.preset_combo.current(0)

    def _load_preset(self):
        """選択したプリセットを読み込む"""
        preset_name = self.preset_var.get()
        if not preset_name:
            messagebox.showwarning("警告", "プリセットを選択してください")
            return

        presets = load_crop_presets()
        for preset in presets:
            if preset.name == preset_name:
                self.crop_regions = copy.deepcopy(preset.regions)
                # ratio モードのリージョンをピクセルモードに変換
                # 基準サイズは1000x1000を使用（編集しやすい値）
                ref_size = 1000
                for region in self.crop_regions:
                    if region.mode == "ratio":
                        region.x = int((region.x_ratio or 0) * ref_size)
                        region.y = int((region.y_ratio or 0) * ref_size)
                        region.width = int((region.width_ratio or 0) * ref_size)
                        region.height = int((region.height_ratio or 0) * ref_size)
                        # ratio値もpx値と同期させる（将来の互換性のため）
                        region.x_ratio = region.x / ref_size
                        region.y_ratio = region.y / ref_size
                        region.width_ratio = region.width / ref_size
                        region.height_ratio = region.height / ref_size
                        region.mode = (
                            "ratio"  # ratioモードを維持（異なる画像サイズに対応）
                        )
                # 配置方向も適用
                self.crop_pos.set(preset.display_position)
                self._update_region_list()
                self._schedule_preview()
                self._save_state()
                messagebox.showinfo(
                    "成功", f"プリセット '{preset_name}' を読み込みました"
                )
                return

        messagebox.showerror("エラー", f"プリセット '{preset_name}' が見つかりません")

    def _save_preset_dialog(self):
        """現在のクロップ設定をプリセットとして保存"""
        from tkinter import simpledialog

        if not self.crop_regions:
            messagebox.showwarning("警告", "保存するクロップ領域がありません")
            return

        name = simpledialog.askstring("プリセット保存", "プリセット名:")
        if not name:
            return

        description = simpledialog.askstring(
            "プリセット保存", "説明 (オプション):", initialvalue=""
        )

        preset = CropPreset(
            name=name,
            regions=copy.deepcopy(self.crop_regions),
            description=description or "",
            display_position=self.crop_pos.get(),  # 現在の配置方向も保存
        )

        try:
            save_crop_preset(preset)
            self._refresh_preset_list()
            messagebox.showinfo("成功", f"プリセット '{name}' を保存しました")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    # -------------------------------------------------------------------------
    # Config & Generation Methods
    # -------------------------------------------------------------------------

    def _load_config_gui(self):
        """設定ファイルを読み込む"""
        path = filedialog.askopenfilename(
            filetypes=[("YAML files", "*.yaml;*.yml"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            config = load_config(path)
            self._apply_config_to_gui(config)
            self._save_state()
            messagebox.showinfo("成功", f"設定を読み込みました: {path}")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込みに失敗しました: {e}")

    def _save_config(self):
        """設定ファイルを保存する"""
        path = filedialog.asksaveasfilename(
            defaultextension=".yaml",
            filetypes=[("YAML files", "*.yaml;*.yml"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            config = self._get_current_config()
            save_config(config, path)
            messagebox.showinfo("成功", f"設定を保存しました: {path}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    def _generate(self):
        """PPTXファイルを生成する"""
        config = self._get_current_config()


        # Validate inputs
        if self.input_mode.get() == "images":
            if not self.images:
                messagebox.showerror("エラー", "画像が追加されていません")
                return
        else:
            if not self.folders:
                messagebox.showerror("エラー", "フォルダが追加されていません")
                return

        try:
            create_grid_presentation(config)
            messagebox.showinfo("成功", f"PPTXを生成しました: {config.output}")
        except Exception as e:
            messagebox.showerror("エラー", f"生成に失敗しました: {e}")

    def _on_region_detail_change(self, *args):
        """Auto-update when region detail fields change."""
        if self._loading_region:
            return
        self._update_region_detail()

    def _update_region_detail(self):
        """選択した領域の詳細を更新"""
        if self.sel_idx is None or self.sel_idx >= len(self.crop_regions):
            return

        region = self.crop_regions[self.sel_idx]
        region.name = self.r_name.get()
        region.x = self._get_safe_int(self.r_x, region.x)
        region.y = self._get_safe_int(self.r_y, region.y)
        region.width = self._get_safe_int(self.r_w, region.width)
        region.height = self._get_safe_int(self.r_h, region.height)
        region.align = self.r_align.get()
        region.offset = self._get_safe_double(self.r_offset, 0.0)
        region.color = self.r_color

        # ratio モードの場合、ratio値も更新（基準サイズ1000で計算）
        if region.mode == "ratio":
            ref_size = 1000
            region.x_ratio = region.x / ref_size
            region.y_ratio = region.y / ref_size
            region.width_ratio = region.width / ref_size
            region.height_ratio = region.height / ref_size

        gap_str = self.r_gap.get().strip()
        if gap_str:
            try:
                region.gap = float(gap_str)
            except ValueError:
                region.gap = None
        else:
            region.gap = None

        region.show_zoomed = self.r_show_zoomed.get()

        # 領域ごとの枠線設定を保存
        scb = self.r_show_crop_border.get()
        region.show_crop_border = True if scb == "true" else False if scb == "false" else None
        cbw = self.r_crop_border_width.get().strip()
        region.crop_border_width = float(cbw) if cbw else None
        region.crop_border_shape = self.r_crop_border_shape.get() or None
        region.crop_border_dash = self.r_crop_border_dash.get() or None

        szb = self.r_show_zoom_border.get()
        region.show_zoom_border = True if szb == "true" else False if szb == "false" else None
        zbw = self.r_zoom_border_width.get().strip()
        region.zoom_border_width = float(zbw) if zbw else None
        region.zoom_border_shape = self.r_zoom_border_shape.get() or None
        region.zoom_border_dash = self.r_zoom_border_dash.get() or None

        self._update_region_list()
        self._schedule_preview()
        self._save_state()


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

    ImageGridApp(root, initial_config)
    root.mainloop()


if __name__ == "__main__":
    main()
