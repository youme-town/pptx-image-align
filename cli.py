#!/usr/bin/env python3
"""
PowerPoint Image Grid Generator - CLI Interface

This script provides a command-line interface for creating PowerPoint presentations
with images arranged in a grid layout.

Usage:
  python cli.py config.yaml                 - Generate presentation from config
  python cli.py --init [filename]           - Create sample config file
  python cli.py --batch config1.yaml ...    - Process multiple config files
  python cli.py --batch-dir <directory>     - Process all YAML files in directory
  python cli.py --help                      - Show help message

Features:
  - All layout features from the GUI version
  - YAML configuration support
  - Auto-detection of grid dimensions from folder structure
  - Batch processing of multiple config files
"""

import sys
from pathlib import Path
from typing import List, Optional

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
  python cli.py <config.yaml>              Generate presentation from config file
  python cli.py --init [filename]          Create sample config (default: config.yaml)
  python cli.py --batch <configs...>       Process multiple config files
  python cli.py --batch-dir <directory>    Process all YAML files in directory
  python cli.py --batch-output <dir> <configs...>  Process with custom output directory
  python cli.py --help                     Show this help message

Configuration File:
  The YAML config file supports all layout options including:
  - Slide dimensions and margins
  - Grid layout (rows, cols, arrangement)
  - Layout mode (grid/flow) with alignment options
  - Image sizing (fit/fixed) with fit modes
  - Crop regions with individual positioning
  - Border styles and widths
  - Text labels and captions
  - Template PPTX support

Example:
  python cli.py --init my_config.yaml         # Create sample config
  python cli.py my_config.yaml                # Generate presentation
  python cli.py --batch a.yaml b.yaml c.yaml  # Batch process multiple configs
  python cli.py --batch-dir ./configs/        # Process all YAML in directory
  python cli.py --batch-output ./out a.yaml   # Output to specific directory

For more information, see the generated sample config file.
"""
    print(help_text)


def batch_process(config_paths: List[str], output_dir: Optional[str] = None) -> int:
    """
    Process multiple config files in batch mode.

    Args:
        config_paths: List of config file paths to process
        output_dir: Optional output directory for all generated files

    Returns:
        Number of successfully processed files
    """
    success_count = 0
    total = len(config_paths)

    for i, config_path in enumerate(config_paths, 1):
        print(f"\n[{i}/{total}] Processing: {config_path}")

        if not Path(config_path).exists():
            print("  Error: File not found, skipping")
            continue

        try:
            config = load_config(config_path)
            auto_detect_grid_size(config)

            # Override output directory if specified
            if output_dir:
                output_name = Path(config.output).name
                config.output = str(Path(output_dir) / output_name)

            output = create_grid_presentation(config)
            print(f"  Created: {output}")
            success_count += 1

        except Exception as e:
            print(f"  Error: {e}")

    return success_count


def batch_process_directory(directory: str, output_dir: Optional[str] = None) -> int:
    """
    Process all YAML config files in a directory.

    Args:
        directory: Directory containing YAML config files
        output_dir: Optional output directory for all generated files

    Returns:
        Number of successfully processed files
    """
    dir_path = Path(directory)
    if not dir_path.is_dir():
        print(f"Error: Not a directory: {directory}")
        return 0

    # Find all YAML files
    yaml_files = list(dir_path.glob("*.yaml")) + list(dir_path.glob("*.yml"))

    if not yaml_files:
        print(f"No YAML files found in: {directory}")
        return 0

    print(f"Found {len(yaml_files)} YAML file(s) in {directory}")
    return batch_process([str(f) for f in yaml_files], output_dir)


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

    # Batch command - process multiple config files
    if arg == "--batch":
        if len(sys.argv) < 3:
            print("Error: --batch requires at least one config file")
            sys.exit(1)
        config_paths = sys.argv[2:]
        print(f"Batch processing {len(config_paths)} config file(s)...")
        success = batch_process(config_paths)
        print(f"\nCompleted: {success}/{len(config_paths)} successful")
        sys.exit(0 if success == len(config_paths) else 1)

    # Batch directory command - process all YAML files in a directory
    if arg == "--batch-dir":
        if len(sys.argv) < 3:
            print("Error: --batch-dir requires a directory path")
            sys.exit(1)
        directory = sys.argv[2]
        output_dir = sys.argv[3] if len(sys.argv) > 3 else None
        success = batch_process_directory(directory, output_dir)
        print(f"\nCompleted: {success} file(s) processed")
        sys.exit(0 if success > 0 else 1)

    # Batch output command - process with custom output directory
    if arg == "--batch-output":
        if len(sys.argv) < 4:
            print("Error: --batch-output requires output directory and at least one config file")
            sys.exit(1)
        output_dir = sys.argv[2]
        config_paths = sys.argv[3:]

        # Create output directory if it doesn't exist
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        print(f"Batch processing {len(config_paths)} config file(s) to {output_dir}...")
        success = batch_process(config_paths, output_dir)
        print(f"\nCompleted: {success}/{len(config_paths)} successful")
        sys.exit(0 if success == len(config_paths) else 1)

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
