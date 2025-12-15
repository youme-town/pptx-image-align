"""
PowerPoint Image Grid Generator CLI

This script creates a PowerPoint presentation with images arranged in a grid layout.
"""

import argparse
import sys
import grid_logic


def generate_sample_config(output_path: str):
    """Generate a sample configuration file."""
    sample = """# PowerPoint Image Grid Generator - Configuration File
slide:
  width: 33.867
  height: 19.05
grid:
  rows: 2
  cols: 3
  arrangement: row
  layout_mode: flow
margin: 1.0
gap: 0.5
image:
  size_mode: fit
  fit_mode: fit
folders: []
output: "output.pptx"
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(sample)
    print(f"Sample config created: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="PowerPoint Image Grid Generator CLI")
    parser.add_argument("config", nargs="?", help="Path to YAML configuration file")
    parser.add_argument(
        "--init", nargs="?", const="config.yaml", help="Create a sample config file"
    )

    args = parser.parse_args()

    if args.init:
        generate_sample_config(args.init)
        return

    if not args.config:
        parser.print_help()
        sys.exit(1)

    try:
        config = grid_logic.load_config(args.config)

        # Auto-determine grid size if not specified
        if config.arrangement == "row":
            if config.rows == 0 and config.folders:
                config.rows = len(config.folders)
            if config.cols == 0 and config.folders:
                try:
                    max_cols = max(
                        len(grid_logic.get_sorted_images(f)) for f in config.folders
                    )
                    config.cols = max(1, max_cols)
                except Exception:
                    config.cols = 3
        else:
            if config.cols == 0 and config.folders:
                config.cols = len(config.folders)
            if config.rows == 0 and config.folders:
                try:
                    max_rows = max(
                        len(grid_logic.get_sorted_images(f)) for f in config.folders
                    )
                    config.rows = max(1, max_rows)
                except Exception:
                    config.rows = 3

        # Ensure minimums
        config.rows = max(1, config.rows)
        config.cols = max(1, config.cols)

        output_file = grid_logic.create_grid_presentation(config)
        print(f"Success! Presentation saved to: {output_file}")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
