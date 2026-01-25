#!/bin/bash
# macOS wrapper script for PowerPoint MCP Server
#
# On macOS with Apple Silicon, Homebrew installs cairo to /opt/homebrew/lib/,
# but Python's ctypes doesn't search there by default. This wrapper ensures
# the library path is set before launching the Python server.
#
# Usage: ./server.sh (or add to Claude Code via `claude mcp add`)

export DYLD_FALLBACK_LIBRARY_PATH="/opt/homebrew/lib:${DYLD_FALLBACK_LIBRARY_PATH}"
exec python3 "$(dirname "$0")/server.py" "$@"
