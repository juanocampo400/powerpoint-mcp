#!/usr/bin/env python3
"""
Download Phosphor Icons (fill variant) for the PowerPoint MCP server.

This script downloads the Phosphor Icons repository from GitHub and extracts
the fill variant SVGs (~1,200 icons) to the icons/phosphor/ directory.
"""

import io
import os
import shutil
import sys
import zipfile
from pathlib import Path
from urllib.request import urlopen


PHOSPHOR_REPO_ZIP = "https://github.com/phosphor-icons/core/archive/refs/heads/main.zip"
ICONS_SUBDIR = "core-main/assets/fill"  # Path within the ZIP file


def download_phosphor_icons(target_dir: Path) -> int:
    """Download and extract Phosphor fill icons.

    Args:
        target_dir: Directory to extract icons to (e.g., icons/phosphor/)

    Returns:
        Number of icons extracted
    """
    print(f"Downloading Phosphor Icons from {PHOSPHOR_REPO_ZIP}...")

    # Download the ZIP file
    with urlopen(PHOSPHOR_REPO_ZIP) as response:
        zip_data = response.read()

    print(f"Downloaded {len(zip_data) / 1024 / 1024:.1f} MB")

    # Create target directory
    target_dir.mkdir(parents=True, exist_ok=True)

    # Clear existing files in target directory
    for f in target_dir.glob("*.svg"):
        f.unlink()

    # Extract fill icons
    count = 0
    with zipfile.ZipFile(io.BytesIO(zip_data)) as zf:
        for name in zf.namelist():
            # Only extract from the fill directory
            if name.startswith(ICONS_SUBDIR) and name.endswith(".svg"):
                # Get just the filename
                filename = os.path.basename(name)
                if filename:
                    # Extract directly to target directory
                    target_path = target_dir / filename
                    with zf.open(name) as src, open(target_path, 'wb') as dst:
                        dst.write(src.read())
                    count += 1

    print(f"Extracted {count} fill icons to {target_dir}")
    return count


def main():
    # Determine project root (parent of scripts directory)
    script_dir = Path(__file__).parent
    project_root = script_dir.parent

    # Target directory for Phosphor icons
    target_dir = project_root / "icons" / "phosphor"

    print(f"Project root: {project_root}")
    print(f"Target directory: {target_dir}")

    try:
        count = download_phosphor_icons(target_dir)
        print(f"\nSuccess! {count} Phosphor fill icons are now available.")
        print(f"Location: {target_dir}")
        return 0
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
