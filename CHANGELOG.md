# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.1.0] - 2026-01-28

### Added
- `get_table_content` tool for reading full table data without evaluate_code
- `modify_table_cell` tool for updating individual cells with formatting preservation

## [1.0.1] - 2026-01-28

### Fixed
- `find_and_replace` now preserves formatting in table cells (operates at run level instead of cell level)

### Changed
- Updated `find_and_replace` docstring to reflect table cell formatting preservation

## [1.0.0] - 2026-01-25

### Initial Release

- Presentation management: create, open, save, save_as, close
- Slide management: add, delete, duplicate, move slides
- Content tools: add_textbox, add_image, add_shape, add_table, add_chart
- Modification tools: modify_shape, delete_shape, find_and_replace
- Icon support with Phosphor Icons (~1,500 icons)
- Slide snapshot for inspecting slide contents
- Code evaluation tool for advanced python-pptx automation
- macOS wrapper script (server.sh) for Homebrew cairo compatibility
