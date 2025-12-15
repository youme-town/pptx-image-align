#!/usr/bin/env python3
"""
PowerPoint Image Grid Generator - CLI Interface

This script provides a command-line interface for creating PowerPoint presentations
with images arranged in a grid layout.

Usage:
  python cli.py config.yaml                 - Generate presentation from config
  python cli.py --init [filename]           - Create sample config file
  python cli.py --help                      - Show help message

Features:
  - All layout features from the GUI version
  - YAML configuration support
  - Auto-detection of grid dimensions from folder structure
"""

import sys
from pathlib import Path

from core import (
    GridConfig,
    load_config,
    generate_sample_config,
    get_sorted_images,
    create_grid_presentation,
)


def print_help():
    """Print help message."""
    help_text = """
PowerPoint Image Grid Generator - CLI

Usage:
  python cli.py <config.yaml>        Generate presentation from config file
  python cli.py --init [filename]    Create sample config (default: config.yaml)
  python cli.py --help               Show this help message

Configuration File:
  The YAML config file supports all layout options including:
  - Slide dimensions and margins
  - Grid layout (rows, cols, arrangement)
  - Layout mode (grid/flow) with alignment options
  - Image sizing (fit/fixed) with fit modes
  - Crop regions with individual positioning
  - Border styles and widths

Example:
  python cli.py --init my_config.yaml    # Create sample config
  python cli.py my_config.yaml           # Generate presentation

For more information, see the generated sample config file.
"""
    print(help_text)


def auto_detect_grid_size(config: GridConfig) -> None:
    """Auto-detect grid dimensions from folder structure if not specified."""
    if not config.folders:
        return

    if config.arrangement == "row":
        if config.rows == 0:
            config.rows = len(config.folders)
        if config.cols == 0:
            try:
                max_cols = max(len(get_sorted_images(f)) for f in config.folders)
                config.cols = max_cols
            except Exception:
                config.cols = 3
    else:  # col arrangement
        if config.cols == 0:
            config.cols = len(config.folders)
        if config.rows == 0:
            try:
                max_rows = max(len(get_sorted_images(f)) for f in config.folders)
                config.rows = max_rows
            except Exception:
                config.rows = 3

    # Ensure minimum values
    if config.rows == 0:
        config.rows = 1
    if config.cols == 0:
        config.cols = 1


def main():
    """Main entry point for CLI."""
    if len(sys.argv) < 2:
        print_help()
        sys.exit(1)

    arg = sys.argv[1]

    # Help command
    if arg in ("--help", "-h"):
        print_help()
        sys.exit(0)

    # Init command
    if arg == "--init":
        output_name = sys.argv[2] if len(sys.argv) > 2 else "config.yaml"
        generate_sample_config(output_name)
        print(f"Sample config created: {output_name}")
        sys.exit(0)

    # Generate presentation from config
    config_path = arg

    if not Path(config_path).exists():
        print(f"Error: Config file not found: {config_path}")
        sys.exit(1)

    try:
        print(f"Loading config: {config_path}")
        config = load_config(config_path)

        # Auto-detect grid size
        auto_detect_grid_size(config)

        print(f"Grid: {config.rows} rows x {config.cols} cols")
        print(f"Layout: {config.layout_mode} mode")
        print(f"Folders: {len(config.folders)} input folder(s)")
        print(f"Crop regions: {len(config.crop_regions)}")
        print(f"Output: {config.output}")
        print()
        print("Generating presentation...")

        output = create_grid_presentation(config)
        print(f"Created: {output}")

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
