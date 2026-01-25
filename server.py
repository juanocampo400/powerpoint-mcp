#!/usr/bin/env python
"""
PowerPoint MCP Server (Custom)
A Model Context Protocol server for creating and editing PowerPoint presentations.
Uses python-pptx for file-based PowerPoint manipulation (no PowerPoint installation required).
"""
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from mcp.server.fastmcp import FastMCP
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
from typing import Optional

# Initialize FastMCP server
mcp = FastMCP("powerpoint-mcp")

# Global state to track the current presentation
class PresentationState:
    def __init__(self):
        self.presentation: Optional[Presentation] = None
        self.file_path: Optional[str] = None
        self.is_modified: bool = False

    def reset(self):
        self.presentation = None
        self.file_path = None
        self.is_modified = False

state = PresentationState()

# Import tools from modules
from tools.presentation import register_presentation_tools
from tools.slides import register_slide_tools
from tools.content import register_content_tools
from tools.icons import register_icon_tools
from tools.modify import register_modify_tools
from tools.evaluate import register_evaluate_tools

# Register all tools
register_presentation_tools(mcp, state)
register_slide_tools(mcp, state)
register_content_tools(mcp, state)
register_icon_tools(mcp, state)
register_modify_tools(mcp, state)
register_evaluate_tools(mcp, state)

if __name__ == "__main__":
    mcp.run()
