"""
PowerPoint Image Grid Generator - Core Module

This module contains shared logic for creating PowerPoint presentations
with images arranged in a grid layout. It is used by both GUI and CLI interfaces.

Features:
- Data classes for configuration
- Layout calculation algorithms
- PPTX generation functions
- Image processing utilities
"""

import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional, List, Tuple
from dataclasses import dataclass, field

import yaml
from PIL import Image
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# =============================================================================
# Constants
# =============================================================================

CM_TO_EMU = 360000
PT_TO_EMU = 12700


# =============================================================================
# Unit Conversion Functions
# =============================================================================


def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMU (English Metric Units)."""
    return int(cm * CM_TO_EMU)


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return int(pt * PT_TO_EMU)


def pt_to_cm(pt: float) -> float:
    """Convert points to centimeters."""
    # 1 inch = 2.54 cm = 72 points
    return pt * (2.54 / 72.0)


# =============================================================================
# Data Classes
# =============================================================================


@dataclass
class CropRegion:
    """Defines a rectangular region to crop from source images."""

    x: int
    y: int
    width: int
    height: int
    color: Tuple[int, int, int] = (255, 0, 0)
    name: str = ""
    align: str = "auto"  # 'auto', 'start', 'center', 'end'
    offset: float = 0.0  # cm, offset from alignment anchor
    gap: Optional[float] = None  # cm, overrides global main-crop gap if set


@dataclass
class GapConfig:
    """Configuration for gap/spacing with support for absolute or relative values."""

    value: float = 0.5
    mode: str = "cm"  # 'cm' or 'scale'

    def to_cm(self, ref_size: float) -> float:
        """Convert gap value to centimeters based on mode."""
        if self.mode == "scale":
            return ref_size * self.value
        return self.value


@dataclass
class CropDisplayConfig:
    """Configuration for how cropped regions are displayed."""

    position: str = "right"  # 'right' or 'bottom'
    main_crop_gap: GapConfig = field(default_factory=lambda: GapConfig(0.15, "cm"))
    crop_crop_gap: GapConfig = field(default_factory=lambda: GapConfig(0.15, "cm"))
    crop_bottom_gap: GapConfig = field(default_factory=lambda: GapConfig(0.0, "cm"))
    size: Optional[float] = None  # Absolute size in cm
    scale: Optional[float] = None  # Scale relative to main image


@dataclass
class GridConfig:
    """Main configuration for grid-based image layout."""

    # Slide dimensions
    slide_width: float = 33.867
    slide_height: float = 19.05

    # Grid structure
    rows: int = 2
    cols: int = 3
    arrangement: str = "row"  # 'row' or 'col'

    # Margins
    margin_left: float = 1.0
    margin_top: float = 1.0
    margin_right: float = 1.0
    margin_bottom: float = 1.0

    # Gap settings
    gap_h: GapConfig = field(default_factory=lambda: GapConfig(0.5, "cm"))
    gap_v: GapConfig = field(default_factory=lambda: GapConfig(0.5, "cm"))

    # Layout mode
    layout_mode: str = "grid"  # 'grid' (aligned) or 'flow' (compact)
    flow_align: str = "left"  # 'left', 'center', 'right'
    flow_vertical_align: str = "center"  # 'top', 'center', 'bottom'

    # Image sizing
    size_mode: str = "fit"  # 'fit' or 'fixed'
    fit_mode: str = "fit"  # 'fit', 'width', 'height'
    image_width: Optional[float] = None
    image_height: Optional[float] = None
    image_scale: float = 1.0

    # Crop settings
    crop_regions: List[CropRegion] = field(default_factory=list)
    crop_rows: Optional[List[int]] = None
    crop_cols: Optional[List[int]] = None
    crop_display: CropDisplayConfig = field(default_factory=CropDisplayConfig)

    # Border settings
    show_crop_border: bool = True
    crop_border_width: float = 1.5
    crop_border_shape: str = "rectangle"  # 'rectangle' or 'rounded'
    show_zoom_border: bool = True
    zoom_border_width: float = 1.5
    zoom_border_shape: str = "rectangle"  # 'rectangle' or 'rounded'

    # Input/Output
    folders: List[str] = field(default_factory=list)
    output: str = "output.pptx"


@dataclass
class LayoutMetrics:
    """Calculated layout metrics for grid positioning."""

    main_width: float
    main_height: float
    col_widths: List[float]
    row_heights: List[float]
    crop_size: float
    crop_main_gap: float
    crop_crop_gap: float
    crop_bottom_gap: float


# =============================================================================
# Configuration Parsing
# =============================================================================


def parse_color(color_value) -> Tuple[int, int, int]:
    """Parse color from various formats (list, tuple, hex string, comma-separated)."""
    if isinstance(color_value, (list, tuple)):
        return tuple(color_value[:3])
    elif isinstance(color_value, str):
        if color_value.startswith("#"):
            hex_color = color_value.lstrip("#")
            return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
        else:
            parts = color_value.split(",")
            return tuple(int(p.strip()) for p in parts)
    return (255, 0, 0)


def parse_gap(data) -> GapConfig:
    """Parse gap configuration from YAML data."""
    if isinstance(data, (int, float)):
        return GapConfig(float(data), "cm")
    elif isinstance(data, dict):
        return GapConfig(float(data.get("value", 0.5)), data.get("mode", "cm"))
    return GapConfig(0.5, "cm")


def load_config(config_path: str) -> GridConfig:
    """Load configuration from a YAML file."""
    with open(config_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    config = GridConfig()

    # Slide settings
    if "slide" in data:
        slide = data["slide"]
        config.slide_width = slide.get("width", config.slide_width)
        config.slide_height = slide.get("height", config.slide_height)

    # Grid settings
    if "grid" in data:
        grid = data["grid"]
        config.rows = grid.get("rows", config.rows)
        config.cols = grid.get("cols", config.cols)
        config.arrangement = grid.get("arrangement", config.arrangement)
        config.layout_mode = grid.get("layout_mode", config.layout_mode)
        config.flow_align = grid.get("flow_align", config.flow_align)
        config.flow_vertical_align = grid.get(
            "flow_vertical_align", config.flow_vertical_align
        )

    # Margin settings
    if "margin" in data:
        margin = data["margin"]
        if isinstance(margin, (int, float)):
            config.margin_left = config.margin_top = margin
            config.margin_right = config.margin_bottom = margin
        else:
            config.margin_left = margin.get("left", config.margin_left)
            config.margin_top = margin.get("top", config.margin_top)
            config.margin_right = margin.get("right", config.margin_right)
            config.margin_bottom = margin.get("bottom", config.margin_bottom)

    # Gap settings
    if "gap" in data:
        gap = data["gap"]
        if isinstance(gap, (int, float)):
            config.gap_h = GapConfig(float(gap), "cm")
            config.gap_v = GapConfig(float(gap), "cm")
        else:
            if "horizontal" in gap:
                config.gap_h = parse_gap(gap["horizontal"])
            if "vertical" in gap:
                config.gap_v = parse_gap(gap["vertical"])

    # Image settings
    if "image" in data:
        img = data["image"]
        config.size_mode = img.get("size_mode", config.size_mode)
        config.fit_mode = img.get("fit_mode", config.fit_mode)
        config.image_scale = img.get("scale", config.image_scale)
        config.image_width = img.get("width")
        config.image_height = img.get("height")

    # Crop settings
    if "crop" in data:
        crop = data["crop"]

        # Parse crop regions
        if "regions" in crop:
            for i, r in enumerate(crop["regions"]):
                region = CropRegion(
                    x=r.get("x", 0),
                    y=r.get("y", 0),
                    width=r.get("width", 100),
                    height=r.get("height", 100),
                    color=parse_color(r.get("color", "#FF0000")),
                    name=r.get("name", f"crop_{i + 1}"),
                    align=r.get("align", "auto"),
                    offset=r.get("offset", 0.0),
                    gap=r.get("gap", None),
                )
                config.crop_regions.append(region)
        elif "region" in crop:
            r = crop["region"]
            region = CropRegion(
                x=r.get("x", 0),
                y=r.get("y", 0),
                width=r.get("width", 100),
                height=r.get("height", 100),
                color=parse_color(r.get("color", "#FF0000")),
                name="crop_1",
                align=r.get("align", "auto"),
                offset=r.get("offset", 0.0),
                gap=r.get("gap", None),
            )
            config.crop_regions.append(region)

        config.crop_rows = crop.get("rows")
        config.crop_cols = crop.get("cols")

        # Crop display settings
        if "display" in crop:
            disp = crop["display"]
            config.crop_display.position = disp.get("position", "right")
            config.crop_display.size = disp.get("size")
            config.crop_display.scale = disp.get("scale")

            # Legacy gap support
            legacy_gap = disp.get("gap")
            if legacy_gap is not None:
                if isinstance(legacy_gap, (int, float)):
                    config.crop_display.main_crop_gap = GapConfig(
                        float(legacy_gap), "cm"
                    )
                    config.crop_display.crop_crop_gap = GapConfig(
                        float(legacy_gap), "cm"
                    )

            if "main_crop_gap" in disp:
                config.crop_display.main_crop_gap = parse_gap(disp["main_crop_gap"])
            if "crop_crop_gap" in disp:
                config.crop_display.crop_crop_gap = parse_gap(disp["crop_crop_gap"])
            if "crop_bottom_gap" in disp:
                config.crop_display.crop_bottom_gap = parse_gap(disp["crop_bottom_gap"])

    # Border settings
    if "border" in data:
        border = data["border"]
        if "crop" in border:
            cb = border["crop"]
            config.show_crop_border = cb.get("show", config.show_crop_border)
            config.crop_border_width = cb.get("width", config.crop_border_width)
            config.crop_border_shape = cb.get("shape", config.crop_border_shape)
        if "zoom" in border:
            zb = border["zoom"]
            config.show_zoom_border = zb.get("show", config.show_zoom_border)
            config.zoom_border_width = zb.get("width", config.zoom_border_width)
            config.zoom_border_shape = zb.get("shape", config.zoom_border_shape)

    # Input/Output
    if "folders" in data:
        config.folders = data["folders"]
    config.output = data.get("output", config.output)

    return config


def save_config(config: GridConfig, output_path: str) -> None:
    """Save configuration to a YAML file."""
    data = {
        "slide": {"width": config.slide_width, "height": config.slide_height},
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
            "horizontal": {"value": config.gap_h.value, "mode": config.gap_h.mode},
            "vertical": {"value": config.gap_v.value, "mode": config.gap_v.mode},
        },
        "image": {
            "size_mode": config.size_mode,
            "fit_mode": config.fit_mode,
            "width": config.image_width,
            "height": config.image_height,
        },
        "folders": config.folders,
        "output": config.output,
        "border": {
            "crop": {
                "show": config.show_crop_border,
                "width": config.crop_border_width,
                "shape": config.crop_border_shape,
            },
            "zoom": {
                "show": config.show_zoom_border,
                "width": config.zoom_border_width,
                "shape": config.zoom_border_shape,
            },
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
    }
    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True)


# =============================================================================
# Image Utilities
# =============================================================================


def extract_number_from_filename(filename: str) -> int:
    """Extract the first number from a filename for sorting."""
    numbers = re.findall(r"\d+", filename)
    return int(numbers[0]) if numbers else 0


def get_sorted_images(folder_path: str) -> List[str]:
    """Get sorted list of image files from a folder."""
    supported_extensions = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp"}
    folder = Path(folder_path)
    if not folder.exists():
        return []

    image_files = [
        f
        for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in supported_extensions
    ]
    image_files.sort(key=lambda f: extract_number_from_filename(f.stem))
    return [str(f) for f in image_files]


def crop_image(image_path: str, region: CropRegion, output_path: str) -> str:
    """Crop an image according to a CropRegion and save to output_path."""
    with Image.open(image_path) as img:
        box = (region.x, region.y, region.x + region.width, region.y + region.height)
        cropped = img.crop(box)
        cropped.save(output_path)
    return output_path


def get_image_dimensions(image_path: str) -> Tuple[int, int]:
    """Get width and height of an image in pixels."""
    with Image.open(image_path) as img:
        return img.size


# =============================================================================
# Layout Calculation
# =============================================================================


def should_apply_crop(row: int, col: int, config: GridConfig) -> bool:
    """Determine if crop regions should be applied to a specific cell."""
    if not config.crop_regions:
        return False
    if config.crop_rows is not None and row not in config.crop_rows:
        return False
    if config.crop_cols is not None and col not in config.crop_cols:
        return False
    return True


def calculate_size_fit_static(
    img_w: int, img_h: int, max_width: float, max_height: float, fit_mode: str = "fit"
) -> Tuple[float, float]:
    """Calculate image size to fit within constraints (static version)."""
    if max_height <= 0 or max_width <= 0 or img_h == 0:
        return 0, 0

    aspect_ratio = img_w / img_h

    if fit_mode == "width":
        width = max_width
        height = width / aspect_ratio
    elif fit_mode == "height":
        height = max_height
        width = height * aspect_ratio
    else:  # 'fit'
        cell_aspect = max_width / max_height
        if aspect_ratio > cell_aspect:
            width = max_width
            height = width / aspect_ratio
        else:
            height = max_height
            width = height * aspect_ratio

    return width, height


def calculate_image_size_fit(
    image_path: str, max_width: float, max_height: float, fit_mode: str = "fit"
) -> Tuple[float, float]:
    """Calculate image size to fit within constraints (from file)."""
    try:
        img_width_px, img_height_px = get_image_dimensions(image_path)
    except Exception:
        return 0, 0
    return calculate_size_fit_static(
        img_width_px, img_height_px, max_width, max_height, fit_mode
    )


def calculate_grid_metrics(config: GridConfig) -> LayoutMetrics:
    """Calculate layout metrics for grid positioning."""
    total_grid_w = config.slide_width - config.margin_left - config.margin_right
    total_grid_h = config.slide_height - config.margin_top - config.margin_bottom

    est_cols = max(1, config.cols)
    est_rows = max(1, config.rows)
    est_main_w = total_grid_w / est_cols
    est_main_h = total_grid_h / est_rows

    gap_h_val = config.gap_h.to_cm(est_main_w)
    gap_v_val = config.gap_v.to_cm(est_main_h)

    avail_w_for_cells = total_grid_w - (gap_h_val * (config.cols - 1))
    avail_h_for_cells = total_grid_h - (gap_v_val * (config.rows - 1))

    # Determine which cells have crops
    has_crops = len(config.crop_regions) > 0
    expanded_cols = set()
    expanded_rows = set()
    if has_crops:
        expanded_cols = (
            set(range(config.cols))
            if config.crop_cols is None
            else set(config.crop_cols)
        )
        expanded_rows = (
            set(range(config.rows))
            if config.crop_rows is None
            else set(config.crop_rows)
        )
    num_exp_cols = len([c for c in range(config.cols) if c in expanded_cols])
    num_exp_rows = len([r for r in range(config.rows) if r in expanded_rows])

    disp = config.crop_display
    gap_mc = disp.main_crop_gap.to_cm(est_main_w)
    gap_cc = disp.crop_crop_gap.to_cm(est_main_w)
    gap_cb = disp.crop_bottom_gap.to_cm(
        est_main_w if disp.position == "right" else est_main_h
    )

    # Border offset compensation
    border_offset = 0.0
    if config.show_zoom_border:
        border_offset = pt_to_cm(config.zoom_border_width)
        gap_mc += border_offset
        gap_cc += border_offset

    main_w = 0.0
    main_h = 0.0
    crop_box_size = 0.0

    is_fixed_size = (
        config.size_mode == "fixed"
        and config.image_width
        and config.image_width > 0
        and config.image_height
        and config.image_height > 0
    )

    if is_fixed_size:
        main_w = config.image_width
        main_h = config.image_height

        if disp.size is not None:
            crop_box_size = disp.size
        elif disp.scale is not None:
            crop_box_size = (
                main_w * disp.scale if disp.position == "right" else main_h * disp.scale
            )
        else:
            crop_box_size = main_w * 0.33 if disp.position == "right" else main_h * 0.33
    else:
        if disp.position == "right":
            main_h = avail_h_for_cells / config.rows
            crop_k = disp.scale if disp.scale is not None else 0.33
            crop_c = disp.size if disp.size is not None else 0.0
            if disp.size is not None:
                crop_k = 0.0

            denom = config.cols + num_exp_cols * crop_k
            numer = avail_w_for_cells - num_exp_cols * (gap_mc + crop_c)
            main_w = numer / denom
            crop_box_size = main_w * crop_k + crop_c
        else:  # bottom
            main_w = avail_w_for_cells / config.cols
            crop_k = disp.scale if disp.scale is not None else 0.33
            crop_c = disp.size if disp.size is not None else 0.0
            if disp.size is not None:
                crop_k = 0.0

            denom = config.rows + num_exp_rows * crop_k
            numer = avail_h_for_cells - num_exp_rows * (gap_mc + crop_c)
            main_h = numer / denom
            crop_box_size = main_h * crop_k + crop_c

    # Calculate column widths and row heights
    col_widths = []
    row_heights = []

    if disp.position == "right":
        extra = crop_box_size + gap_mc
        for c in range(config.cols):
            col_widths.append(main_w + extra if c in expanded_cols else main_w)
        row_heights = [main_h] * config.rows
    else:  # bottom
        extra = crop_box_size + gap_mc + gap_cb
        col_widths = [main_w] * config.cols
        for r in range(config.rows):
            row_heights.append(main_h + extra if r in expanded_rows else main_h)

    return LayoutMetrics(
        main_w, main_h, col_widths, row_heights, crop_box_size, gap_mc, gap_cc, gap_cb
    )


def calculate_item_bounds(
    config: GridConfig,
    metrics: LayoutMetrics,
    image_path: str,
    row_idx: int,
    col_idx: int,
    border_offset_cm: float = 0.0,
    override_size: Optional[Tuple[float, float]] = None,
) -> Tuple[float, float, float, float]:
    """
    Calculate bounding box of an item (main image + crops) relative to main image top-left (0,0).
    Returns (min_x, min_y, max_x, max_y).
    """
    has_crops = should_apply_crop(row_idx, col_idx, config)

    if override_size:
        orig_w, orig_h = override_size
    elif image_path == "dummy":
        return 0, 0, 0, 0
    else:
        orig_w, orig_h = calculate_image_size_fit(
            image_path, metrics.main_width, metrics.main_height, config.fit_mode
        )

    min_x, min_y = 0.0, 0.0
    max_x, max_y = orig_w, orig_h
    half_border = border_offset_cm / 2.0 if config.show_zoom_border else 0.0

    if has_crops and config.crop_regions:
        num_crops = len(config.crop_regions)
        disp = config.crop_display
        actual_gap_mc = disp.main_crop_gap.to_cm(
            orig_w if disp.position == "right" else orig_h
        )
        actual_gap_cc = disp.crop_crop_gap.to_cm(
            orig_w if disp.position == "right" else orig_h
        )
        actual_gap_cb = disp.crop_bottom_gap.to_cm(
            orig_w if disp.position == "right" else orig_h
        )

        if config.show_zoom_border:
            actual_gap_mc += border_offset_cm
            actual_gap_cc += border_offset_cm

        for crop_idx, region in enumerate(config.crop_regions):
            # Calculate crop size
            cw, ch = _calculate_crop_size(
                config,
                metrics,
                disp,
                image_path,
                orig_w,
                orig_h,
                num_crops,
                actual_gap_cc,
                override_size,
            )

            # Calculate position
            this_gap_mc = region.gap if region.gap is not None else actual_gap_mc
            if region.gap is not None and config.show_zoom_border:
                this_gap_mc += border_offset_cm

            if disp.position == "right":
                c_left = orig_w + this_gap_mc
                c_top = _calculate_crop_vertical_position(
                    region,
                    crop_idx,
                    num_crops,
                    orig_h,
                    ch,
                    actual_gap_cc,
                    half_border,
                    disp,
                )
            else:  # bottom
                c_top = orig_h + this_gap_mc
                c_left = _calculate_crop_horizontal_position(
                    region,
                    crop_idx,
                    num_crops,
                    orig_w,
                    cw,
                    actual_gap_cc,
                    half_border,
                    disp,
                )

            # Update bounds
            min_x = min(min_x, c_left)
            min_y = min(min_y, c_top)
            max_x = max(max_x, c_left + cw)
            max_y = max(max_y, c_top + ch)

        if disp.position == "bottom":
            max_y += actual_gap_cb

    return min_x, min_y, max_x, max_y


def _calculate_crop_size(
    config: GridConfig,
    metrics: LayoutMetrics,
    disp: CropDisplayConfig,
    image_path: str,
    orig_w: float,
    orig_h: float,
    num_crops: int,
    actual_gap_cc: float,
    override_size: Optional[Tuple[float, float]],
) -> Tuple[float, float]:
    """Calculate the size of a crop region display."""
    if override_size:
        # Dummy/preview mode
        if disp.size is not None:
            if disp.position == "right":
                return calculate_size_fit_static(100, 100, disp.size, 9999, "width")
            else:
                return calculate_size_fit_static(100, 100, 9999, disp.size, "height")
        elif disp.scale is not None:
            if disp.position == "right":
                tw = orig_w * disp.scale
                return calculate_size_fit_static(100, 100, tw, 9999, "width")
            else:
                th = orig_h * disp.scale
                return calculate_size_fit_static(100, 100, 9999, th, "height")
        else:
            if disp.position == "right":
                single_h = (orig_h - actual_gap_cc * (num_crops - 1)) / num_crops
                return calculate_size_fit_static(
                    100, 100, metrics.crop_size, single_h, "fit"
                )
            else:
                sw = (orig_w - actual_gap_cc * (num_crops - 1)) / num_crops
                return calculate_size_fit_static(100, 100, sw, metrics.crop_size, "fit")
    else:
        # Real image mode
        if disp.size is not None:
            if disp.position == "right":
                return calculate_image_size_fit(image_path, disp.size, 9999, "width")
            else:
                return calculate_image_size_fit(image_path, 9999, disp.size, "height")
        elif disp.scale is not None:
            if disp.position == "right":
                tw = orig_w * disp.scale
                return calculate_image_size_fit(image_path, tw, 9999, "width")
            else:
                th = orig_h * disp.scale
                return calculate_image_size_fit(image_path, 9999, th, "height")
        else:
            if disp.position == "right":
                single_h = (orig_h - actual_gap_cc * (num_crops - 1)) / num_crops
                return calculate_image_size_fit(
                    image_path, metrics.crop_size, single_h, "fit"
                )
            else:
                sw = (orig_w - actual_gap_cc * (num_crops - 1)) / num_crops
                return calculate_image_size_fit(
                    image_path, sw, metrics.crop_size, "fit"
                )


def _calculate_crop_vertical_position(
    region: CropRegion,
    crop_idx: int,
    num_crops: int,
    orig_h: float,
    ch: float,
    actual_gap_cc: float,
    half_border: float,
    disp: CropDisplayConfig,
) -> float:
    """Calculate vertical position for a crop (used when position='right')."""
    if region.align == "start":
        return region.offset + half_border
    elif region.align == "center":
        return (orig_h - ch) / 2 + region.offset
    elif region.align == "end":
        return orig_h - ch + region.offset - half_border
    else:  # auto
        if disp.scale is not None or disp.size is not None:
            return crop_idx * (ch + actual_gap_cc)
        else:
            # Fit logic (pin ends)
            if crop_idx == 0:
                return 0.0
            elif num_crops > 1 and crop_idx == num_crops - 1:
                return orig_h - ch
            else:
                single_h = (orig_h - actual_gap_cc * (num_crops - 1)) / num_crops
                slot_top = crop_idx * (single_h + actual_gap_cc)
                return slot_top + (single_h - ch) / 2


def _calculate_crop_horizontal_position(
    region: CropRegion,
    crop_idx: int,
    num_crops: int,
    orig_w: float,
    cw: float,
    actual_gap_cc: float,
    half_border: float,
    disp: CropDisplayConfig,
) -> float:
    """Calculate horizontal position for a crop (used when position='bottom')."""
    if region.align == "start":
        return region.offset + half_border
    elif region.align == "center":
        return (orig_w - cw) / 2 + region.offset
    elif region.align == "end":
        return orig_w - cw + region.offset - half_border
    else:  # auto
        if disp.scale is not None or disp.size is not None:
            return crop_idx * (cw + actual_gap_cc)
        else:
            if crop_idx == 0:
                return 0.0
            elif num_crops > 1 and crop_idx == num_crops - 1:
                return orig_w - cw
            else:
                sw = (orig_w - actual_gap_cc * (num_crops - 1)) / num_crops
                slot_left = crop_idx * (sw + actual_gap_cc)
                return slot_left + (sw - cw) / 2


def calculate_flow_row_heights(
    config: GridConfig,
    metrics: LayoutMetrics,
    image_grid: List[List[Optional[str]]],
    border_offset_cm: float,
) -> List[float]:
    """Calculate maximum content height for each row in flow mode."""
    row_heights = []

    for row_idx, row_images in enumerate(image_grid):
        if row_idx >= config.rows:
            break

        row_max_h = 0.0
        for col_idx, image_path in enumerate(row_images):
            if col_idx >= config.cols or image_path is None:
                continue

            min_x, min_y, max_x, max_y = calculate_item_bounds(
                config, metrics, image_path, row_idx, col_idx, border_offset_cm
            )
            item_h = max_y - min_y
            row_max_h = max(row_max_h, item_h)

        row_heights.append(row_max_h if row_max_h > 0 else metrics.main_height)

    return row_heights


# =============================================================================
# PPTX Generation
# =============================================================================


def add_border_shape(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    border_color: Tuple[int, int, int],
    border_width: float,
    shape_type: str = "rectangle",
) -> None:
    """Add a border shape to a slide."""
    ms_shape_type = (
        MSO_SHAPE.ROUNDED_RECTANGLE if shape_type == "rounded" else MSO_SHAPE.RECTANGLE
    )

    shape = slide.shapes.add_shape(
        ms_shape_type,
        cm_to_emu(left),
        cm_to_emu(top),
        cm_to_emu(width),
        cm_to_emu(height),
    )
    shape.fill.background()
    shape.line.color.rgb = RGBColor(*border_color)
    shape.line.width = pt_to_emu(border_width)
    shape.shadow.inherit = False


def add_crop_borders_to_image(
    slide,
    image_left: float,
    image_top: float,
    image_width: float,
    image_height: float,
    original_image_path: str,
    crop_regions: List[CropRegion],
    border_width: float,
    border_shape: str = "rectangle",
) -> None:
    """Add crop region borders overlaid on the main image."""
    try:
        orig_width_px, orig_height_px = get_image_dimensions(original_image_path)
    except Exception:
        return

    scale_x = image_width / orig_width_px
    scale_y = image_height / orig_height_px

    for region in crop_regions:
        border_left = image_left + region.x * scale_x
        border_top = image_top + region.y * scale_y
        border_w = region.width * scale_x
        border_h = region.height * scale_y
        add_border_shape(
            slide,
            border_left,
            border_top,
            border_w,
            border_h,
            region.color,
            border_width,
            border_shape,
        )


def create_grid_presentation(config: GridConfig) -> str:
    """Create a PowerPoint presentation with images arranged in a grid layout."""
    prs = Presentation()
    prs.slide_width = cm_to_emu(config.slide_width)
    prs.slide_height = cm_to_emu(config.slide_height)
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Build image grid
    if config.arrangement == "row":
        image_grid = [get_sorted_images(folder) for folder in config.folders]
    else:
        columns = [get_sorted_images(folder) for folder in config.folders]
        max_len = max(len(col) for col in columns) if columns else 0
        image_grid = [
            [col[i] if i < len(col) else None for col in columns]
            for i in range(max_len)
        ]

    temp_dir = tempfile.mkdtemp()
    num_crops = len(config.crop_regions)
    metrics = calculate_grid_metrics(config)

    border_offset_cm = (
        pt_to_cm(config.crop_border_width) if config.show_zoom_border else 0.0
    )
    half_border = border_offset_cm / 2.0 if config.show_zoom_border else 0.0

    # Calculate row heights for flow mode
    flow_row_heights = []
    total_content_height = 0.0
    if config.layout_mode == "flow":
        flow_row_heights = calculate_flow_row_heights(
            config, metrics, image_grid, border_offset_cm
        )
        for i, rh in enumerate(flow_row_heights):
            total_content_height += rh
            if i < len(flow_row_heights) - 1:
                total_content_height += config.gap_v.to_cm(metrics.main_height)

    try:
        # Determine start Y based on flow vertical alignment
        current_y = config.margin_top
        if config.layout_mode == "flow":
            avail_h = config.slide_height - config.margin_top - config.margin_bottom
            if config.flow_vertical_align == "center":
                current_y = config.margin_top + (avail_h - total_content_height) / 2
            elif config.flow_vertical_align == "bottom":
                current_y = (config.margin_top + avail_h) - total_content_height

        for row_idx, row_images in enumerate(image_grid):
            if row_idx >= config.rows:
                break

            # Determine row height
            if config.layout_mode == "flow":
                current_row_height = (
                    flow_row_heights[row_idx]
                    if row_idx < len(flow_row_heights)
                    else metrics.main_height
                )
            else:
                current_row_height = (
                    metrics.row_heights[row_idx]
                    if row_idx < len(metrics.row_heights)
                    else metrics.main_height
                )

            this_gap_v = config.gap_v.to_cm(metrics.main_height)
            current_x = config.margin_left

            # Pre-calculate row width for flow alignment
            if config.layout_mode == "flow":
                row_total_content_width = 0.0
                valid_items = 0

                for col_idx, image_path in enumerate(row_images):
                    if col_idx >= config.cols or image_path is None:
                        continue
                    min_x, min_y, max_x, max_y = calculate_item_bounds(
                        config, metrics, image_path, row_idx, col_idx, border_offset_cm
                    )
                    row_total_content_width += max_x - min_x
                    valid_items += 1

                if valid_items > 1:
                    row_total_content_width += (valid_items - 1) * config.gap_h.to_cm(
                        metrics.main_width
                    )

                avail_w = config.slide_width - config.margin_left - config.margin_right
                if config.flow_align == "center":
                    current_x = (
                        config.margin_left + (avail_w - row_total_content_width) / 2
                    )
                elif config.flow_align == "right":
                    current_x = config.margin_left + (avail_w - row_total_content_width)

            # Process each cell
            for col_idx, image_path in enumerate(row_images):
                if col_idx >= config.cols:
                    break

                this_gap_h = config.gap_h.to_cm(metrics.main_width)

                if image_path is None:
                    if config.layout_mode == "grid":
                        w = (
                            metrics.col_widths[col_idx]
                            if col_idx < len(metrics.col_widths)
                            else metrics.main_width
                        )
                        current_x += w + this_gap_h
                    continue

                has_crops = should_apply_crop(row_idx, col_idx, config)
                orig_w, orig_h = calculate_image_size_fit(
                    image_path, metrics.main_width, metrics.main_height, config.fit_mode
                )

                # Calculate dynamic gaps
                global_gap_mc = config.crop_display.main_crop_gap.to_cm(
                    orig_w if config.crop_display.position == "right" else orig_h
                )
                global_gap_cc = config.crop_display.crop_crop_gap.to_cm(
                    orig_w if config.crop_display.position == "right" else orig_h
                )

                if config.show_zoom_border:
                    global_gap_mc += border_offset_cm
                    global_gap_cc += border_offset_cm

                # Calculate item bounds and position
                min_x, min_y, max_x, max_y = calculate_item_bounds(
                    config, metrics, image_path, row_idx, col_idx, border_offset_cm
                )
                item_w = max_x - min_x
                item_h = max_y - min_y

                if config.layout_mode == "flow":
                    item_draw_left = current_x
                    item_draw_top = current_y + (current_row_height - item_h) / 2
                else:
                    cell_w = (
                        metrics.col_widths[col_idx]
                        if col_idx < len(metrics.col_widths)
                        else metrics.main_width
                    )
                    item_draw_left = current_x + (cell_w - item_w) / 2
                    item_draw_top = current_y + (current_row_height - item_h) / 2

                final_main_left = item_draw_left - min_x
                final_main_top = item_draw_top - min_y

                # Add main image
                pic = slide.shapes.add_picture(
                    image_path,
                    cm_to_emu(final_main_left),
                    cm_to_emu(final_main_top),
                    cm_to_emu(orig_w),
                    cm_to_emu(orig_h),
                )
                pic.shadow.inherit = False

                if has_crops and config.show_crop_border:
                    add_crop_borders_to_image(
                        slide,
                        final_main_left,
                        final_main_top,
                        orig_w,
                        orig_h,
                        image_path,
                        config.crop_regions,
                        config.crop_border_width,
                        config.crop_border_shape,
                    )

                # Add cropped images
                if has_crops and num_crops > 0:
                    _add_crop_images(
                        slide,
                        config,
                        metrics,
                        image_path,
                        temp_dir,
                        row_idx,
                        col_idx,
                        final_main_left,
                        final_main_top,
                        orig_w,
                        orig_h,
                        global_gap_mc,
                        global_gap_cc,
                        border_offset_cm,
                        half_border,
                    )

                # Advance X position
                if config.layout_mode == "flow":
                    current_x += item_w + this_gap_h
                else:
                    w = (
                        metrics.col_widths[col_idx]
                        if col_idx < len(metrics.col_widths)
                        else metrics.main_width
                    )
                    current_x += w + this_gap_h

            current_y += current_row_height + this_gap_v

        prs.save(config.output)
    finally:
        shutil.rmtree(temp_dir)

    return config.output


def _add_crop_images(
    slide,
    config: GridConfig,
    metrics: LayoutMetrics,
    image_path: str,
    temp_dir: str,
    row_idx: int,
    col_idx: int,
    final_main_left: float,
    final_main_top: float,
    orig_w: float,
    orig_h: float,
    global_gap_mc: float,
    global_gap_cc: float,
    border_offset_cm: float,
    half_border: float,
) -> None:
    """Add cropped images next to the main image."""
    disp = config.crop_display
    num_crops = len(config.crop_regions)

    for crop_idx, region in enumerate(config.crop_regions):
        crop_filename = f"crop_{row_idx}_{col_idx}_{crop_idx}.png"
        crop_path = os.path.join(temp_dir, crop_filename)
        try:
            crop_image(image_path, region, crop_path)
        except Exception:
            continue

        # Calculate crop size
        cw, ch = _resolve_crop_size(
            config, metrics, disp, crop_path, orig_w, orig_h, num_crops, global_gap_cc
        )

        # Calculate gap
        this_gap_mc = region.gap if region.gap is not None else global_gap_mc
        if region.gap is not None and config.show_zoom_border:
            this_gap_mc += border_offset_cm

        # Calculate position
        if disp.position == "right":
            c_left = final_main_left + orig_w + this_gap_mc
            c_top = _calculate_crop_absolute_vertical_position(
                region,
                crop_idx,
                num_crops,
                final_main_top,
                orig_h,
                ch,
                global_gap_cc,
                half_border,
                disp,
            )
        else:  # bottom
            c_top = final_main_top + orig_h + this_gap_mc
            c_left = _calculate_crop_absolute_horizontal_position(
                region,
                crop_idx,
                num_crops,
                final_main_left,
                orig_w,
                cw,
                global_gap_cc,
                half_border,
                disp,
            )

        pic_crop = slide.shapes.add_picture(
            crop_path,
            cm_to_emu(c_left),
            cm_to_emu(c_top),
            cm_to_emu(cw),
            cm_to_emu(ch),
        )
        pic_crop.shadow.inherit = False

        if config.show_zoom_border:
            add_border_shape(
                slide,
                c_left,
                c_top,
                cw,
                ch,
                region.color,
                config.zoom_border_width,
                config.zoom_border_shape,
            )


def _resolve_crop_size(
    config: GridConfig,
    metrics: LayoutMetrics,
    disp: CropDisplayConfig,
    crop_path: str,
    orig_w: float,
    orig_h: float,
    num_crops: int,
    global_gap_cc: float,
) -> Tuple[float, float]:
    """Resolve the size of a crop image."""
    if disp.size is not None:
        if disp.position == "right":
            return calculate_image_size_fit(crop_path, disp.size, 9999, "width")
        else:
            return calculate_image_size_fit(crop_path, 9999, disp.size, "height")
    elif disp.scale is not None:
        if disp.position == "right":
            return calculate_image_size_fit(
                crop_path, orig_w * disp.scale, 9999, "width"
            )
        else:
            return calculate_image_size_fit(
                crop_path, 9999, orig_h * disp.scale, "height"
            )
    else:
        if disp.position == "right":
            single_h = (orig_h - global_gap_cc * (num_crops - 1)) / num_crops
            return calculate_image_size_fit(
                crop_path, metrics.crop_size, single_h, "fit"
            )
        else:
            single_w = (orig_w - global_gap_cc * (num_crops - 1)) / num_crops
            return calculate_image_size_fit(
                crop_path, single_w, metrics.crop_size, "fit"
            )


def _calculate_crop_absolute_vertical_position(
    region: CropRegion,
    crop_idx: int,
    num_crops: int,
    final_main_top: float,
    orig_h: float,
    ch: float,
    global_gap_cc: float,
    half_border: float,
    disp: CropDisplayConfig,
) -> float:
    """Calculate absolute vertical position for a crop."""
    if region.align == "start":
        return final_main_top + region.offset + half_border
    elif region.align == "center":
        return final_main_top + (orig_h - ch) / 2 + region.offset
    elif region.align == "end":
        return final_main_top + orig_h - ch + region.offset - half_border
    else:  # auto
        if disp.scale is not None or disp.size is not None:
            return final_main_top + crop_idx * (ch + global_gap_cc)
        else:
            if crop_idx == 0:
                return final_main_top
            elif num_crops > 1 and crop_idx == num_crops - 1:
                return (final_main_top + orig_h) - ch
            else:
                single_h = (orig_h - global_gap_cc * (num_crops - 1)) / num_crops
                slot_top = final_main_top + crop_idx * (single_h + global_gap_cc)
                return slot_top + (single_h - ch) / 2


def _calculate_crop_absolute_horizontal_position(
    region: CropRegion,
    crop_idx: int,
    num_crops: int,
    final_main_left: float,
    orig_w: float,
    cw: float,
    global_gap_cc: float,
    half_border: float,
    disp: CropDisplayConfig,
) -> float:
    """Calculate absolute horizontal position for a crop."""
    if region.align == "start":
        return final_main_left + region.offset + half_border
    elif region.align == "center":
        return final_main_left + (orig_w - cw) / 2 + region.offset
    elif region.align == "end":
        return final_main_left + orig_w - cw + region.offset - half_border
    else:  # auto
        if disp.scale is not None or disp.size is not None:
            return final_main_left + crop_idx * (cw + global_gap_cc)
        else:
            if crop_idx == 0:
                return final_main_left
            elif num_crops > 1 and crop_idx == num_crops - 1:
                return (final_main_left + orig_w) - cw
            else:
                single_w = (orig_w - global_gap_cc * (num_crops - 1)) / num_crops
                slot_left = final_main_left + crop_idx * (single_w + global_gap_cc)
                return slot_left + (single_w - cw) / 2


# =============================================================================
# Sample Configuration Generator
# =============================================================================

SAMPLE_CONFIG = """# PowerPoint Image Grid Generator - Configuration File
# Size values are in centimeters (cm) unless otherwise specified (e.g. pt)

# Slide size settings
slide:
  width: 33.867    # 16:9 aspect ratio
  height: 19.05

# Main grid layout
grid:
  rows: 2
  cols: 3
  arrangement: row  # 'row' or 'col'
  layout_mode: flow # 'grid' (aligned) or 'flow' (compact)
  flow_align: left # 'left', 'center', 'right' (only for flow)
  flow_vertical_align: center # 'top', 'center', 'bottom' (only for flow)

# Margins around the grid (cm)
margin:
  left: 1.0
  top: 1.0
  right: 1.0
  bottom: 1.0

# Gap between images in grid
gap:
  horizontal:
    value: 0.5
    mode: cm # 'cm' or 'scale'
  vertical:
    value: 0.5
    mode: cm

# Image sizing
image:
  size_mode: fit   # 'fit' (auto calc) or 'fixed' (user defined)
  fit_mode: fit    # 'fit', 'width', 'height'
  # width: 10.0    # used if size_mode is fixed
  # height: 7.5
  scale: 1.0

# Crop settings
crop:
  # Multiple crop regions
  regions:
    - name: "Region A"
      x: 50
      y: 50
      width: 100
      height: 100
      color: "#FF0000"
      align: auto   # auto, start, center, end
      offset: 0.0   # cm
      # gap: 0.5    # overrides main_crop_gap
    - name: "Region B"
      x: 200
      y: 100
      width: 120
      height: 80
      color: "#00FF00"
  
  # Which rows/cols to crop (0-indexed), null = all
  rows: [0]
  cols: null
  
  # How to display cropped images next to the original
  display:
    position: right   # 'right' or 'bottom'
    size: null        # Absolute size (cm). If set, takes priority over scale.
    scale: 0.4        # Display size ratio relative to main image
    main_crop_gap: {value: 0.15, mode: cm}
    crop_crop_gap: {value: 0.15, mode: cm}
    crop_bottom_gap: {value: 0.0, mode: cm}

# Border settings
border:
  crop:
    show: true
    width: 1.5  # Line width in points (pt)
    shape: rectangle # 'rectangle' or 'rounded'
  zoom:
    show: true
    width: 1.5  # Line width in points (pt)
    shape: rectangle # 'rectangle' or 'rounded'

# Input folders
folders:
  - "./images/row1"
  - "./images/row2"

# Output file
output: "output.pptx"
"""


def generate_sample_config(output_path: str) -> None:
    """Generate a sample configuration file."""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(SAMPLE_CONFIG)
