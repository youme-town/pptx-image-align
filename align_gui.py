"""
PowerPoint Image Grid Generator GUI

This application provides a graphical interface for creating PowerPoint presentations
with images arranged in a grid layout. It uses `grid_logic` for backend calculations.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from typing import Optional, List, Tuple

import yaml
from PIL import Image, ImageTk

# Import shared backend logic
import grid_logic
from grid_logic import (
    GridConfig,
    GapConfig,
    CropRegion,
    calculate_grid_metrics,
    calculate_item_bounds,
    calculate_size_fit_static,
    pt_to_cm,
    should_apply_crop,
)


class CropEditor(tk.Toplevel):
    def __init__(self, parent, image_path, callback):
        super().__init__(parent)
        self.title("Crop Editor")
        self.geometry("900x700")
        self.callback = callback

        self.image_path = image_path
        self.orig_img = Image.open(image_path)
        self.orig_w, self.orig_h = self.orig_img.size

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

        ttk.Button(self.bottom_frame, text="Save", command=self.on_save).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(self.bottom_frame, text="Cancel", command=self.destroy).pack(
            side=tk.RIGHT
        )

        self.rect_id = None
        self.start_x = None
        self.start_y = None
        self.cur_rect = None
        self.display_scale = 1.0
        self.tk_img = None
        self.off_x = 0
        self.off_y = 0

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
        x1 = min(self.start_x, end_x) - self.off_x
        y1 = min(self.start_y, end_y) - self.off_y
        x2 = max(self.start_x, end_x) - self.off_x
        y2 = max(self.start_y, end_y) - self.off_y
        ox1 = max(0, int(x1 / self.display_scale))
        oy1 = max(0, int(y1 / self.display_scale))
        ox2 = min(self.orig_w, int(x2 / self.display_scale))
        oy2 = min(self.orig_h, int(y2 / self.display_scale))
        w, h = ox2 - ox1, oy2 - oy1
        if w > 0 and h > 0:
            self.cur_rect = (ox1, oy1, w, h)

    def on_save(self):
        if self.cur_rect:
            self.callback(*self.cur_rect, self.var_name.get())
        self.destroy()


class ImageGridApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Grid Generator GUI")
        self.root.geometry("1400x950")

        self.output_path = tk.StringVar(value="output.pptx")
        self.rows = tk.IntVar(value=3)
        self.cols = tk.IntVar(value=3)
        self.arrangement = tk.StringVar(value="row")
        self.layout_mode = tk.StringVar(value="flow")
        self.flow_align = tk.StringVar(value="left")
        self.flow_vertical_align = tk.StringVar(value="center")

        self.slide_w = tk.DoubleVar(value=33.867)
        self.slide_h = tk.DoubleVar(value=19.05)
        self.margin_l = tk.DoubleVar(value=1.0)
        self.margin_t = tk.DoubleVar(value=1.0)
        self.margin_r = tk.DoubleVar(value=1.0)
        self.margin_b = tk.DoubleVar(value=1.0)

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

        self.image_size_mode = tk.StringVar(value="fit")
        self.image_fit_mode = tk.StringVar(value="fit")
        self.image_w = tk.DoubleVar(value=10.0)
        self.image_h = tk.DoubleVar(value=7.5)
        self.crop_pos = tk.StringVar(value="right")

        self.crop_size_mode = tk.StringVar(value="scale")
        self.crop_size_val = tk.DoubleVar(value=0.0)
        self.crop_scale_val = tk.DoubleVar(value=0.4)

        self.show_crop_border = tk.BooleanVar(value=True)
        self.crop_border_w = tk.DoubleVar(value=1.5)
        self.zoom_border_shape = tk.StringVar(value="rectangle")
        self.show_zoom_border = tk.BooleanVar(value=True)
        self.zoom_border_w = tk.DoubleVar(value=1.5)

        self.dummy_ratio_w = tk.DoubleVar(value=1.0)
        self.dummy_ratio_h = tk.DoubleVar(value=1.0)

        self.folders = []
        self.crop_regions = []
        self.crop_rows_filter = tk.StringVar(value="")
        self.crop_cols_filter = tk.StringVar(value="")

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

        self.create_widgets()
        self.add_preview_tracers()
        self.root.after(500, self.update_preview)

    def create_widgets(self):
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)
        self.left_frame = ttk.Frame(self.paned)
        self.paned.add(self.left_frame, weight=1)
        self.right_frame = ttk.Frame(self.paned)
        self.paned.add(self.right_frame, weight=2)

        top_frame = ttk.Frame(self.left_frame, padding=5)
        top_frame.pack(fill=tk.X)
        ttk.Button(top_frame, text="Load Config", command=self.load_config_gui).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(top_frame, text="Save Config", command=self.save_config).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(top_frame, text="Generate PPTX", command=self.generate).pack(
            side=tk.RIGHT, padx=2
        )

        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.setup_tab_basic(ttk.Frame(self.notebook, padding=10))
        self.setup_tab_layout(ttk.Frame(self.notebook, padding=10))
        self.setup_tab_crop(ttk.Frame(self.notebook, padding=10))
        self.setup_tab_style(ttk.Frame(self.notebook, padding=10))

        f_prev_set = ttk.Frame(self.right_frame, padding=5)
        f_prev_set.pack(fill=tk.X)
        ttk.Label(f_prev_set, text="Preview Dummy Ratio (W:H)").pack(
            side=tk.LEFT, padx=5
        )
        ttk.Entry(f_prev_set, textvariable=self.dummy_ratio_w, width=5).pack(
            side=tk.LEFT
        )
        ttk.Label(f_prev_set, text=":").pack(side=tk.LEFT)
        ttk.Entry(f_prev_set, textvariable=self.dummy_ratio_h, width=5).pack(
            side=tk.LEFT
        )

        canvas_frame = ttk.Frame(self.right_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.preview_canvas = tk.Canvas(canvas_frame, bg="#e0e0e0")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self.schedule_preview())

    def setup_tab_basic(self, frame):
        self.notebook.add(frame, text="Basic")
        f_out = ttk.LabelFrame(frame, text="Output File", padding=5)
        f_out.pack(fill=tk.X, pady=5)
        ttk.Entry(f_out, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5
        )
        ttk.Button(f_out, text="Browse", command=self.browse_output).pack(side=tk.RIGHT)

        f_folders = ttk.LabelFrame(frame, text="Image Folders", padding=5)
        f_folders.pack(fill=tk.BOTH, expand=True, pady=5)
        btn_frame = ttk.Frame(f_folders)
        btn_frame.pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="Add", command=self.add_folder).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="Remove", command=self.remove_folder).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_frame, text="Clear", command=self.clear_folders).pack(
            side=tk.LEFT, padx=2
        )
        self.folder_listbox = tk.Listbox(f_folders, selectmode=tk.EXTENDED)
        self.folder_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        f_grid = ttk.LabelFrame(frame, text="Grid Structure", padding=5)
        f_grid.pack(fill=tk.X, pady=5)
        ttk.Label(f_grid, text="Rows:").grid(row=0, column=0, padx=5)
        ttk.Entry(f_grid, textvariable=self.rows, width=5).grid(row=0, column=1)
        ttk.Label(f_grid, text="Cols:").grid(row=0, column=2, padx=5)
        ttk.Entry(f_grid, textvariable=self.cols, width=5).grid(row=0, column=3)
        ttk.Label(f_grid, text="Order:").grid(row=1, column=0, padx=5)
        ttk.Radiobutton(
            f_grid, text="Row", variable=self.arrangement, value="row"
        ).grid(row=1, column=1)
        ttk.Radiobutton(
            f_grid, text="Col", variable=self.arrangement, value="col"
        ).grid(row=1, column=2)

    def setup_tab_layout(self, frame):
        self.notebook.add(frame, text="Layout")
        f_mode = ttk.LabelFrame(frame, text="Layout Mode", padding=5)
        f_mode.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(
            f_mode, text="Flow", variable=self.layout_mode, value="flow"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_mode, text="Grid", variable=self.layout_mode, value="grid"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Label(f_mode, text="| Align:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(
            f_mode,
            textvariable=self.flow_align,
            values=["left", "center", "right"],
            width=7,
        ).pack(side=tk.LEFT)
        ttk.Label(f_mode, text="| V-Align:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(
            f_mode,
            textvariable=self.flow_vertical_align,
            values=["top", "center", "bottom"],
            width=7,
        ).pack(side=tk.LEFT)

        f_slide = ttk.LabelFrame(frame, text="Slide (cm)", padding=5)
        f_slide.pack(fill=tk.X, pady=5)
        ttk.Label(f_slide, text="W:").pack(side=tk.LEFT)
        ttk.Entry(f_slide, textvariable=self.slide_w, width=6).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Label(f_slide, text="H:").pack(side=tk.LEFT)
        ttk.Entry(f_slide, textvariable=self.slide_h, width=6).pack(
            side=tk.LEFT, padx=5
        )

        f_mg = ttk.LabelFrame(frame, text="Margins (cm)", padding=5)
        f_mg.pack(fill=tk.X, pady=5)
        for val, label in zip(
            [self.margin_l, self.margin_t, self.margin_r, self.margin_b], "LTRB"
        ):
            ttk.Label(f_mg, text=f"{label}:").pack(side=tk.LEFT)
            ttk.Entry(f_mg, textvariable=val, width=4).pack(side=tk.LEFT)

        f_gap = ttk.LabelFrame(frame, text="Gaps", padding=5)
        f_gap.pack(fill=tk.X, pady=5)
        for i, (l, v, m) in enumerate(
            [
                ("H:", self.gap_h_val, self.gap_h_mode),
                ("V:", self.gap_v_val, self.gap_v_mode),
            ]
        ):
            ttk.Label(f_gap, text=l).grid(row=i, column=0)
            ttk.Entry(f_gap, textvariable=v, width=5).grid(row=i, column=1)
            ttk.Radiobutton(f_gap, text="cm", variable=m, value="cm").grid(
                row=i, column=2
            )
            ttk.Radiobutton(f_gap, text="Scale", variable=m, value="scale").grid(
                row=i, column=3
            )

        f_img = ttk.LabelFrame(frame, text="Image Sizing", padding=5)
        f_img.pack(fill=tk.X, pady=5)
        ttk.Label(f_img, text="Fit Mode:").pack(side=tk.LEFT)
        ttk.Combobox(
            f_img,
            textvariable=self.image_fit_mode,
            values=["fit", "width", "height"],
            width=7,
        ).pack(side=tk.LEFT)
        ttk.Label(f_img, text=" | Size Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_img, text="Fit", variable=self.image_size_mode, value="fit"
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_img, text="Fixed", variable=self.image_size_mode, value="fixed"
        ).pack(side=tk.LEFT)

        f_fix = ttk.Frame(f_img)
        f_fix.pack(fill=tk.X, pady=2)
        ttk.Label(f_fix, text="Fixed W:").pack(side=tk.LEFT)
        ttk.Entry(f_fix, textvariable=self.image_w, width=5).pack(side=tk.LEFT)
        ttk.Label(f_fix, text="Fixed H:").pack(side=tk.LEFT)
        ttk.Entry(f_fix, textvariable=self.image_h, width=5).pack(side=tk.LEFT)

    def setup_tab_crop(self, frame):
        self.notebook.add(frame, text="Crops")
        f_reg = ttk.LabelFrame(frame, text="Regions", padding=5)
        f_reg.pack(fill=tk.X, pady=5)

        btn_row = ttk.Frame(f_reg)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="Open Editor", command=self.open_crop_editor).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_row, text="Add Manual", command=self.add_region_dialog).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_row, text="Remove", command=self.remove_region).pack(
            side=tk.LEFT, padx=2
        )

        self.region_tree = ttk.Treeview(
            f_reg, columns=("name", "xywh", "align"), show="headings", height=4
        )
        for col in ["name", "xywh", "align"]:
            self.region_tree.heading(col, text=col.title())
        self.region_tree.pack(fill=tk.BOTH, expand=True)
        self.region_tree.bind("<<TreeviewSelect>>", self.on_region_select)

        f_det = ttk.LabelFrame(frame, text="Selected Region Details", padding=5)
        f_det.pack(fill=tk.X, pady=5)
        r1 = ttk.Frame(f_det)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="Name:").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.r_name, width=10).pack(side=tk.LEFT)
        self.btn_r_color = tk.Button(
            r1, text="Color", width=5, command=self.pick_region_color
        )
        self.btn_r_color.pack(side=tk.LEFT, padx=5)
        r2 = ttk.Frame(f_det)
        r2.pack(fill=tk.X)
        for l, v in zip("XYWH", [self.r_x, self.r_y, self.r_w, self.r_h]):
            ttk.Label(r2, text=f"{l}:").pack(side=tk.LEFT)
            ttk.Entry(r2, textvariable=v, width=5).pack(side=tk.LEFT)
        r3 = ttk.Frame(f_det)
        r3.pack(fill=tk.X)
        ttk.Label(r3, text="Align:").pack(side=tk.LEFT)
        ttk.Combobox(
            r3,
            textvariable=self.r_align,
            values=["auto", "start", "center", "end"],
            width=7,
        ).pack(side=tk.LEFT)
        ttk.Label(r3, text="Off(cm):").pack(side=tk.LEFT)
        ttk.Entry(r3, textvariable=self.r_offset, width=5).pack(side=tk.LEFT)
        ttk.Label(r3, text="Gap:").pack(side=tk.LEFT)
        ttk.Entry(r3, textvariable=self.r_gap, width=5).pack(side=tk.LEFT)
        ttk.Button(f_det, text="Update", command=self.update_region_detail).pack(
            anchor=tk.E
        )

        f_glob = ttk.LabelFrame(frame, text="Crop Layout", padding=5)
        f_glob.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(
            f_glob, text="Right", variable=self.crop_pos, value="right"
        ).grid(row=0, column=0)
        ttk.Radiobutton(
            f_glob, text="Bottom", variable=self.crop_pos, value="bottom"
        ).grid(row=0, column=1)

        for i, (txt, v, m) in enumerate(
            [
                ("Main-Crop", self.gap_mc_val, self.gap_mc_mode),
                ("Crop-Crop", self.gap_cc_val, self.gap_cc_mode),
                ("Crop-Btm", self.gap_cb_val, self.gap_cb_mode),
            ],
            start=1,
        ):
            ttk.Label(f_glob, text=txt).grid(row=i, column=0)
            ttk.Entry(f_glob, textvariable=v, width=5).grid(row=i, column=1)
            ttk.Radiobutton(f_glob, text="cm", variable=m, value="cm").grid(
                row=i, column=2
            )
            ttk.Radiobutton(f_glob, text="Sc", variable=m, value="scale").grid(
                row=i, column=3
            )

        f_sz = ttk.Frame(f_glob)
        f_sz.grid(row=4, column=0, columnspan=4, sticky="w")
        ttk.Radiobutton(
            f_sz, text="Scale", variable=self.crop_size_mode, value="scale"
        ).pack(side=tk.LEFT)
        ttk.Entry(f_sz, textvariable=self.crop_scale_val, width=5).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_sz, text="Fix(cm)", variable=self.crop_size_mode, value="size"
        ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Entry(f_sz, textvariable=self.crop_size_val, width=5).pack(side=tk.LEFT)

        f_filter = ttk.LabelFrame(frame, text="Filters (Rows/Cols)", padding=5)
        f_filter.pack(fill=tk.X, pady=5)
        ttk.Label(f_filter, text="Rows:").pack(side=tk.LEFT)
        ttk.Entry(f_filter, textvariable=self.crop_rows_filter, width=10).pack(
            side=tk.LEFT
        )
        ttk.Label(f_filter, text="Cols:").pack(side=tk.LEFT)
        ttk.Entry(f_filter, textvariable=self.crop_cols_filter, width=10).pack(
            side=tk.LEFT
        )

    def setup_tab_style(self, frame):
        self.notebook.add(frame, text="Style")
        f_style = ttk.LabelFrame(frame, text="Borders", padding=5)
        f_style.pack(fill=tk.X, pady=5)
        f_src = ttk.LabelFrame(f_style, text="Source Image Border", padding=5)
        f_src.pack(fill=tk.X)
        ttk.Checkbutton(f_src, text="Show", variable=self.show_crop_border).pack(
            side=tk.LEFT
        )
        ttk.Label(f_src, text="Width(pt):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(f_src, textvariable=self.crop_border_w, width=5).pack(side=tk.LEFT)
        f_zoom = ttk.LabelFrame(f_style, text="Cropped Image Border", padding=5)
        f_zoom.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(f_zoom, text="Show", variable=self.show_zoom_border).pack(
            side=tk.LEFT
        )
        ttk.Label(f_zoom, text="Width(pt):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(f_zoom, textvariable=self.zoom_border_w, width=5).pack(side=tk.LEFT)
        ttk.Radiobutton(
            f_zoom, text="Rect", variable=self.zoom_border_shape, value="rectangle"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            f_zoom, text="Rounded", variable=self.zoom_border_shape, value="rounded"
        ).pack(side=tk.LEFT)

    def add_preview_tracers(self):
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
            self.show_crop_border,
            self.crop_border_w,
            self.show_zoom_border,
            self.zoom_border_w,
            self.zoom_border_shape,
            self.r_name,
            self.r_x,
            self.r_y,
            self.r_w,
            self.r_h,
            self.r_align,
            self.r_offset,
            self.r_gap,
        ]
        for v in vars_to_trace:
            v.trace_add("write", lambda *args: self.schedule_preview())

    def schedule_preview(self):
        if hasattr(self, "_after_id"):
            self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(100, self.update_preview)

    def get_safe(self, var, default):
        try:
            return var.get()
        except tk.TclError:
            return default

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PPTX", "*.pptx")]
        )
        if path:
            self.output_path.set(path)

    def add_folder(self):
        p = filedialog.askdirectory()
        if p:
            self.folders.append(p)
            self.folder_listbox.insert(tk.END, p)
            try:
                images = grid_logic.get_sorted_images(p)
                if images:
                    with Image.open(images[0]) as img:
                        w, h = img.size
                        self.dummy_ratio_w.set(1.0)
                        self.dummy_ratio_h.set(h / w)
            except Exception:
                pass
            self.schedule_preview()

    def remove_folder(self):
        s = self.folder_listbox.curselection()
        if s:
            self.folders.pop(s[0])
            self.folder_listbox.delete(s[0])
            self.schedule_preview()

    def clear_folders(self):
        self.folders = []
        self.folder_listbox.delete(0, tk.END)
        self.schedule_preview()

    def open_crop_editor(self):
        target_img = None
        for f in self.folders:
            imgs = grid_logic.get_sorted_images(f)
            if imgs:
                target_img = imgs[0]
                break
        if not target_img:
            target_img = filedialog.askopenfilename(
                filetypes=[("Images", "*.png;*.jpg;*.jpeg")]
            )
        if target_img:
            CropEditor(self.root, target_img, self.add_region_from_editor)

    def add_region_from_editor(self, x, y, w, h, name):
        self.crop_regions.append(CropRegion(x, y, w, h, (255, 0, 0), name))
        self.update_region_list()
        self.schedule_preview()

    def add_region_dialog(self):
        self.add_region_from_editor(50, 50, 100, 100, f"R{len(self.crop_regions) + 1}")

    def remove_region(self):
        s = self.region_tree.selection()
        if s:
            idx = self.region_tree.index(s[0])
            self.crop_regions.pop(idx)
            self.update_region_list()
            self.schedule_preview()

    def update_region_list(self):
        self.region_tree.delete(*self.region_tree.get_children())
        for r in self.crop_regions:
            self.region_tree.insert(
                "",
                tk.END,
                values=(r.name, f"{r.x},{r.y},{r.width},{r.height}", r.align),
            )

    def on_region_select(self, event):
        sel = self.region_tree.selection()
        if not sel:
            return
        self.sel_idx = self.region_tree.index(sel[0])
        r = self.crop_regions[self.sel_idx]
        self.r_name.set(r.name)
        self.r_x.set(r.x)
        self.r_y.set(r.y)
        self.r_w.set(r.width)
        self.r_h.set(r.height)
        self.r_color = r.color
        self.btn_r_color.config(bg=f"#{r.color[0]:02x}{r.color[1]:02x}{r.color[2]:02x}")
        self.r_align.set(r.align)
        self.r_offset.set(r.offset)
        self.r_gap.set(str(r.gap) if r.gap is not None else "")

    def update_region_detail(self):
        if self.sel_idx is None or self.sel_idx >= len(self.crop_regions):
            return
        r = self.crop_regions[self.sel_idx]
        r.name = self.r_name.get()
        r.x = self.get_safe(self.r_x, 0)
        r.y = self.get_safe(self.r_y, 0)
        r.width = self.get_safe(self.r_w, 100)
        r.height = self.get_safe(self.r_h, 100)
        r.color = self.r_color
        r.align = self.r_align.get()
        r.offset = self.get_safe(self.r_offset, 0.0)
        g = self.r_gap.get().strip()
        r.gap = float(g) if g else None
        self.update_region_list()
        self.schedule_preview()

    def pick_region_color(self):
        c = colorchooser.askcolor(color=self.r_color)
        if c[0]:
            self.r_color = tuple(map(int, c[0]))
            self.btn_r_color.config(bg=c[1])

    def get_current_config(self) -> GridConfig:
        config = GridConfig()
        config.layout_mode = self.layout_mode.get()
        config.flow_align = self.flow_align.get()
        config.flow_vertical_align = self.flow_vertical_align.get()
        config.slide_width = self.get_safe(self.slide_w, 33.867)
        config.slide_height = self.get_safe(self.slide_h, 19.05)
        config.rows = self.get_safe(self.rows, 2)
        config.cols = self.get_safe(self.cols, 3)
        config.arrangement = self.arrangement.get()
        config.margin_left = self.get_safe(self.margin_l, 1.0)
        config.margin_top = self.get_safe(self.margin_t, 1.0)
        config.margin_right = self.get_safe(self.margin_r, 1.0)
        config.margin_bottom = self.get_safe(self.margin_b, 1.0)

        config.gap_h = GapConfig(
            self.get_safe(self.gap_h_val, 0.5), self.gap_h_mode.get()
        )
        config.gap_v = GapConfig(
            self.get_safe(self.gap_v_val, 0.5), self.gap_v_mode.get()
        )

        config.size_mode = self.image_size_mode.get()
        config.fit_mode = self.image_fit_mode.get()
        config.image_width = self.get_safe(self.image_w, 10.0)
        config.image_height = self.get_safe(self.image_h, 7.5)

        config.folders = self.folders if self.folders else []
        config.crop_regions = self.crop_regions

        cr = self.crop_rows_filter.get().strip()
        if cr:
            config.crop_rows = [int(x) for x in cr.split(",") if x.isdigit()]
        cc = self.crop_cols_filter.get().strip()
        if cc:
            config.crop_cols = [int(x) for x in cc.split(",") if x.isdigit()]

        config.crop_display.position = self.crop_pos.get()
        config.crop_display.main_crop_gap = GapConfig(
            self.get_safe(self.gap_mc_val, 0.15), self.gap_mc_mode.get()
        )
        config.crop_display.crop_crop_gap = GapConfig(
            self.get_safe(self.gap_cc_val, 0.15), self.gap_cc_mode.get()
        )
        config.crop_display.crop_bottom_gap = GapConfig(
            self.get_safe(self.gap_cb_val, 0.0), self.gap_cb_mode.get()
        )

        if self.crop_size_mode.get() == "size":
            config.crop_display.size = self.get_safe(self.crop_size_val, 0.0)
        else:
            config.crop_display.scale = self.get_safe(self.crop_scale_val, 0.4)

        config.show_crop_border = self.show_crop_border.get()
        config.crop_border_width = self.get_safe(self.crop_border_w, 1.5)
        config.show_zoom_border = self.show_zoom_border.get()
        config.zoom_border_width = self.get_safe(self.zoom_border_w, 1.5)
        config.zoom_border_shape = self.zoom_border_shape.get()
        config.output = self.output_path.get()

        if config.rows < 1:
            config.rows = 1
        if config.cols < 1:
            config.cols = 1
        return config

    def update_preview(self):
        config = self.get_current_config()
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

        canvas.create_rectangle(
            tx(0),
            ty(0),
            tx(config.slide_width),
            ty(config.slide_height),
            fill="white",
            outline="black",
        )

        metrics = calculate_grid_metrics(config)

        d_rw = self.get_safe(self.dummy_ratio_w, 1.0)
        d_rh = self.get_safe(self.dummy_ratio_h, 1.0)
        if d_rw <= 0:
            d_rw = 1.0
        if d_rh <= 0:
            d_rh = 1.0
        dummy_w_px = 400
        dummy_h_px = 400 * (d_rh / d_rw)

        dummy_sz_cm = calculate_size_fit_static(
            dummy_w_px,
            dummy_h_px,
            metrics.main_width,
            metrics.main_height,
            config.fit_mode,
        )
        border_offset_cm = 0.0
        if config.show_zoom_border:
            border_offset_cm = pt_to_cm(config.zoom_border_width)
        half_border = border_offset_cm / 2.0 if config.show_zoom_border else 0.0

        # Set line join style for preview
        # Note: create_rectangle doesn't support joinstyle, use create_polygon
        join_style = "miter" if config.zoom_border_shape == "rectangle" else "round"

        cy = config.margin_top
        flow_row_heights = []

        if config.layout_mode == "flow":
            total_content_h = 0.0
            for r in range(config.rows):
                row_h = 0.0
                for c in range(config.cols):
                    _, min_y, _, max_y = calculate_item_bounds(
                        config,
                        metrics,
                        "dummy",
                        r,
                        c,
                        border_offset_cm,
                        override_size=dummy_sz_cm,
                    )
                    row_h = max(row_h, max_y - min_y)
                if row_h == 0:
                    row_h = metrics.main_height
                flow_row_heights.append(row_h)
                total_content_h += row_h
            avail_h = config.slide_height - config.margin_top - config.margin_bottom
            total_content_h += (config.rows - 1) * config.gap_v.to_cm(
                metrics.main_height
            )
            if config.flow_vertical_align == "center":
                cy = config.margin_top + (avail_h - total_content_h) / 2
            elif config.flow_vertical_align == "bottom":
                cy = (config.margin_top + avail_h) - total_content_h

        num_crops = len(config.crop_regions)
        disp = config.crop_display
        ref_size = dummy_sz_cm[0] if disp.position == "right" else dummy_sz_cm[1]
        gap_mc = disp.main_crop_gap.to_cm(ref_size)
        gap_cc = disp.crop_crop_gap.to_cm(ref_size)
        if config.show_zoom_border:
            gap_mc += border_offset_cm
            gap_cc += border_offset_cm

        for r in range(config.rows):
            if config.layout_mode == "flow":
                cur_rh = (
                    flow_row_heights[r]
                    if r < len(flow_row_heights)
                    else metrics.main_height
                )
            else:
                cur_rh = (
                    metrics.row_heights[r]
                    if r < len(metrics.row_heights)
                    else metrics.main_height
                )

            cx = config.margin_left
            if config.layout_mode == "flow":
                row_w = 0.0
                count = 0
                for c in range(config.cols):
                    min_x, _, max_x, _ = calculate_item_bounds(
                        config,
                        metrics,
                        "dummy",
                        r,
                        c,
                        border_offset_cm,
                        override_size=dummy_sz_cm,
                    )
                    row_w += max_x - min_x
                    count += 1
                if count > 1:
                    row_w += (count - 1) * config.gap_h.to_cm(metrics.main_width)
                avail_w = config.slide_width - config.margin_left - config.margin_right
                if config.flow_align == "center":
                    cx = config.margin_left + (avail_w - row_w) / 2
                elif config.flow_align == "right":
                    cx = config.margin_left + (avail_w - row_w)

            for c in range(config.cols):
                min_x, min_y, max_x, max_y = calculate_item_bounds(
                    config,
                    metrics,
                    "dummy",
                    r,
                    c,
                    border_offset_cm,
                    override_size=dummy_sz_cm,
                )
                item_w = max_x - min_x
                item_h = max_y - min_y

                if config.layout_mode == "flow":
                    draw_x = cx
                    draw_y = cy + (cur_rh - item_h) / 2
                else:
                    cell_w = (
                        metrics.col_widths[c]
                        if c < len(metrics.col_widths)
                        else metrics.main_width
                    )
                    draw_x = cx + (cell_w - item_w) / 2
                    draw_y = cy + (cur_rh - item_h) / 2

                main_l = draw_x - min_x
                main_t = draw_y - min_y

                canvas.create_rectangle(
                    tx(main_l),
                    ty(main_t),
                    tx(main_l + dummy_sz_cm[0]),
                    ty(main_t + dummy_sz_cm[1]),
                    fill="#ddf",
                    outline="blue",
                )

                if should_apply_crop(r, c, config):
                    # Item Bounding Box
                    canvas.create_rectangle(
                        tx(draw_x),
                        ty(draw_y),
                        tx(draw_x + item_w),
                        ty(draw_y + item_h),
                        outline="red",
                        dash=(2, 2),
                    )

                    if num_crops > 0:
                        for crop_idx, region in enumerate(config.crop_regions):
                            cw_px = region.width if region.width > 0 else 100
                            ch_px = region.height if region.height > 0 else 100
                            cw, ch = 0, 0

                            if disp.size is not None:
                                if disp.position == "right":
                                    cw, ch = calculate_size_fit_static(
                                        cw_px, ch_px, disp.size, 9999, "width"
                                    )
                                else:
                                    cw, ch = calculate_size_fit_static(
                                        cw_px, ch_px, 9999, disp.size, "height"
                                    )
                            elif disp.scale is not None:
                                if disp.position == "right":
                                    cw, ch = calculate_size_fit_static(
                                        cw_px,
                                        ch_px,
                                        dummy_sz_cm[0] * disp.scale,
                                        9999,
                                        "width",
                                    )
                                else:
                                    cw, ch = calculate_size_fit_static(
                                        cw_px,
                                        ch_px,
                                        9999,
                                        dummy_sz_cm[1] * disp.scale,
                                        "height",
                                    )
                            else:
                                if disp.position == "right":
                                    sh = (
                                        dummy_sz_cm[1] - gap_cc * (num_crops - 1)
                                    ) / num_crops
                                    cw, ch = calculate_size_fit_static(
                                        cw_px, ch_px, metrics.crop_size, sh, "fit"
                                    )
                                else:
                                    sw = (
                                        dummy_sz_cm[0] - gap_cc * (num_crops - 1)
                                    ) / num_crops
                                    cw, ch = calculate_size_fit_static(
                                        cw_px, ch_px, sw, metrics.crop_size, "fit"
                                    )

                            reg_gap = region.gap if region.gap is not None else gap_mc
                            if region.gap is not None and config.show_zoom_border:
                                reg_gap += border_offset_cm

                            if disp.position == "right":
                                c_left = main_l + dummy_sz_cm[0] + reg_gap
                                if region.align == "start":
                                    c_top = main_t + region.offset + half_border
                                elif region.align == "center":
                                    c_top = (
                                        main_t
                                        + (dummy_sz_cm[1] - ch) / 2
                                        + region.offset
                                    )
                                elif region.align == "end":
                                    c_top = (
                                        main_t
                                        + dummy_sz_cm[1]
                                        - ch
                                        + region.offset
                                        - half_border
                                    )
                                else:  # auto
                                    if disp.scale or disp.size:
                                        c_top = main_t + crop_idx * (ch + gap_cc)
                                    else:
                                        sh = (
                                            dummy_sz_cm[1] - gap_cc * (num_crops - 1)
                                        ) / num_crops
                                        slot_top = main_t + crop_idx * (sh + gap_cc)
                                        c_top = slot_top + (sh - ch) / 2
                            else:
                                c_top = main_t + dummy_sz_cm[1] + reg_gap
                                if region.align == "start":
                                    c_left = main_l + region.offset + half_border
                                elif region.align == "center":
                                    c_left = (
                                        main_l
                                        + (dummy_sz_cm[0] - cw) / 2
                                        + region.offset
                                    )
                                elif region.align == "end":
                                    c_left = (
                                        main_l
                                        + dummy_sz_cm[0]
                                        - cw
                                        + region.offset
                                        - half_border
                                    )
                                else:
                                    if disp.scale or disp.size:
                                        c_left = main_l + crop_idx * (cw + gap_cc)
                                    else:
                                        sw = (
                                            dummy_sz_cm[0] - gap_cc * (num_crops - 1)
                                        ) / num_crops
                                        slot_left = main_l + crop_idx * (sw + gap_cc)
                                        c_left = slot_left + (sw - cw) / 2

                            col_hex = (
                                f"#{region.color[0]:02x}"
                                f"{region.color[1]:02x}"
                                f"{region.color[2]:02x}"
                            )
                            # Use create_polygon to support joinstyle
                            x1, y1 = tx(c_left), ty(c_top)
                            x2, y2 = tx(c_left + cw), ty(c_top + ch)

                            canvas.create_polygon(
                                x1,
                                y1,
                                x2,
                                y1,
                                x2,
                                y2,
                                x1,
                                y2,
                                fill="#fee",
                                outline=col_hex,
                                width=2,
                                joinstyle=join_style,
                            )

                if config.layout_mode == "flow":
                    cx += item_w + config.gap_h.to_cm(metrics.main_width)
                else:
                    w = (
                        metrics.col_widths[c]
                        if c < len(metrics.col_widths)
                        else metrics.main_width
                    )
                    cx += w + config.gap_h.to_cm(metrics.main_width)
            cy += cur_rh + config.gap_v.to_cm(metrics.main_height)

    def save_config(self):
        f = filedialog.asksaveasfilename(
            defaultextension=".yaml", filetypes=[("YAML", "*.yaml")]
        )
        if not f:
            return
        try:
            config = self.get_current_config()
            data = {
                "slide": {
                    "width": config.slide_width,
                    "height": config.slide_height,
                },
                "grid": {
                    "rows": config.rows,
                    "cols": config.cols,
                    "arrangement": config.arrangement,
                    "layout_mode": config.layout_mode,
                    "flow_align": config.flow_align,
                    "flow_vertical_align": config.flow_vertical_align,
                },
                "margin": {
                    "left": config.margin_left,
                    "top": config.margin_top,
                    "right": config.margin_right,
                    "bottom": config.margin_bottom,
                },
                "gap": {
                    "horizontal": {
                        "value": config.gap_h.value,
                        "mode": config.gap_h.mode,
                    },
                    "vertical": {
                        "value": config.gap_v.value,
                        "mode": config.gap_v.mode,
                    },
                },
                "image": {
                    "size_mode": config.size_mode,
                    "fit_mode": config.fit_mode,
                    "width": config.image_width,
                    "height": config.image_height,
                },
                "crop": {
                    "regions": [
                        {
                            "name": r.name,
                            "x": r.x,
                            "y": r.y,
                            "width": r.width,
                            "height": r.height,
                            "color": list(r.color),
                            "align": r.align,
                            "offset": r.offset,
                            "gap": r.gap,
                        }
                        for r in config.crop_regions
                    ],
                    "rows": config.crop_rows,
                    "cols": config.crop_cols,
                    "display": {
                        "position": config.crop_display.position,
                        "size": config.crop_display.size,
                        "scale": config.crop_display.scale,
                        "main_crop_gap": {
                            "value": config.crop_display.main_crop_gap.value,
                            "mode": config.crop_display.main_crop_gap.mode,
                        },
                        "crop_crop_gap": {
                            "value": config.crop_display.crop_crop_gap.value,
                            "mode": config.crop_display.crop_crop_gap.mode,
                        },
                        "crop_bottom_gap": {
                            "value": config.crop_display.crop_bottom_gap.value,
                            "mode": config.crop_display.crop_bottom_gap.mode,
                        },
                    },
                },
                "border": {
                    "crop": {
                        "show": config.show_crop_border,
                        "width": config.crop_border_width,
                    },
                    "zoom": {
                        "show": config.show_zoom_border,
                        "width": config.zoom_border_width,
                        "shape": config.zoom_border_shape,
                    },
                },
                "folders": config.folders,
                "output": config.output,
            }
            with open(f, "w", encoding="utf-8") as yf:
                yaml.dump(data, yf, allow_unicode=True)
            messagebox.showinfo("Saved", "Config saved.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_config_gui(self):
        f = filedialog.askopenfilename(filetypes=[("YAML", "*.yaml")])
        if f:
            try:
                self.apply_config_to_gui(grid_logic.load_config(f))
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def apply_config_to_gui(self, config: GridConfig):
        self.rows.set(config.rows)
        self.cols.set(config.cols)
        self.arrangement.set(config.arrangement)
        self.layout_mode.set(config.layout_mode)
        self.flow_align.set(config.flow_align)
        self.flow_vertical_align.set(config.flow_vertical_align)
        self.slide_w.set(config.slide_width)
        self.slide_h.set(config.slide_height)
        self.margin_l.set(config.margin_left)
        self.margin_t.set(config.margin_top)
        self.margin_r.set(config.margin_right)
        self.margin_b.set(config.margin_bottom)
        self.gap_h_val.set(config.gap_h.value)
        self.gap_h_mode.set(config.gap_h.mode)
        self.gap_v_val.set(config.gap_v.value)
        self.gap_v_mode.set(config.gap_v.mode)
        self.image_size_mode.set(config.size_mode)
        self.image_fit_mode.set(config.fit_mode)
        if config.image_width:
            self.image_w.set(config.image_width)
        if config.image_height:
            self.image_h.set(config.image_height)
        self.folder_listbox.delete(0, tk.END)
        self.folders = config.folders
        for p in config.folders:
            self.folder_listbox.insert(tk.END, p)
        self.crop_regions = config.crop_regions
        self.update_region_list()
        if config.crop_rows:
            self.crop_rows_filter.set(",".join(map(str, config.crop_rows)))
        else:
            self.crop_rows_filter.set("")
        if config.crop_cols:
            self.crop_cols_filter.set(",".join(map(str, config.crop_cols)))
        else:
            self.crop_cols_filter.set("")
        self.crop_pos.set(config.crop_display.position)
        self.gap_mc_val.set(config.crop_display.main_crop_gap.value)
        self.gap_mc_mode.set(config.crop_display.main_crop_gap.mode)
        self.gap_cc_val.set(config.crop_display.crop_crop_gap.value)
        self.gap_cc_mode.set(config.crop_display.crop_crop_gap.mode)
        self.gap_cb_val.set(config.crop_display.crop_bottom_gap.value)
        self.gap_cb_mode.set(config.crop_display.crop_bottom_gap.mode)
        if config.crop_display.size:
            self.crop_size_mode.set("size")
            self.crop_size_val.set(config.crop_display.size)
        else:
            self.crop_size_mode.set("scale")
            self.crop_scale_val.set(
                config.crop_display.scale if config.crop_display.scale else 0.4
            )
        self.show_crop_border.set(config.show_crop_border)
        self.crop_border_w.set(config.crop_border_width)
        self.show_zoom_border.set(config.show_zoom_border)
        self.zoom_border_w.set(config.zoom_border_width)
        self.zoom_border_shape.set(config.zoom_border_shape)
        self.output_path.set(config.output)
        self.schedule_preview()

    def generate(self):
        try:
            grid_logic.create_grid_presentation(self.get_current_config())
            messagebox.showinfo("Success", "PPTX Generated!")
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageGridApp(root)
    root.mainloop()
