# PowerPoint MCP Server

This is a Model Context Protocol (MCP) server for creating and editing PowerPoint presentations. Built with Python and python-pptx.

## Features

- **Presentation Management**: Create, open, save, and close presentations
- **Slide Operations**: Add, delete, duplicate, and reorder slides
- **Content Creation**: Add textboxes, images, shapes, tables, and charts
- **Icons**: Insert Phosphor SVG icons with custom colors
- **Modifications**: Modify shapes, delete elements, find and replace text
- **Escape Hatch**: Execute arbitrary python-pptx code for advanced operations

## Requirements

- Python 3.9 or higher
- No PowerPoint installation required – works directly with .pptx files

## Installation

1. Clone the repository:
```bash
git clone https://github.com/juanocampo400/powerpoint-mcp.git
cd powerpoint-mcp
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

2. (Optional) For icon support:
   - **Windows**: `pip install cairosvg` (works if pycairo is installed, which is common with graphics/PDF tools)
   - **macOS**: `brew install cairo pango && pip install cairosvg`

## Usage with Claude

Add to your Claude configuration using Claude Code CLI:

**macOS:**
```bash
cd powerpoint-mcp
chmod +x server.sh
claude mcp add powerpoint-mcp --scope user -- $PWD/server.sh
```

**Windows (Git Bash):**
```bash
cd powerpoint-mcp
claude mcp add powerpoint-mcp --scope user -- python $PWD/server.py
```

Note: `--scope user` makes the server available globally. Without it, the server only works when you're in the project directory.

> **Why `server.sh` on macOS?** The wrapper script sets `DYLD_FALLBACK_LIBRARY_PATH` so Python can find the Homebrew-installed cairo library (required for icon support). Without this, icons may not work even if cairo is installed correctly.

Or manually edit `~/.claude.json`:

**macOS:**
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "type": "stdio",
      "command": "/Users/yourname/powerpoint-mcp/server.sh",
      "args": []
    }
  }
}
```

**Windows:**
```json
{
  "mcpServers": {
    "powerpoint-mcp": {
      "type": "stdio",
      "command": "C:/Users/yourname/AppData/Local/Programs/Python/Python312/python.exe",
      "args": ["C:/Users/yourname/powerpoint-mcp/server.py"]
    }
  }
}
```

## Available Tools

### Core Management
| Tool | Description |
|------|-------------|
| `manage_presentation` | Open, create, save, save_as, close presentations |
| `get_presentation_info` | Get slide count, dimensions, overview |
| `manage_slide` | Add, delete, duplicate, move slides |
| `get_slide_snapshot` | Get detailed info about shapes on a slide |

### Content Creation
| Tool | Description |
|------|-------------|
| `add_textbox` | Add text with formatting (font, size, color, alignment) |
| `add_image` | Insert images (PNG, JPG, etc.) |
| `add_shape` | Add shapes (rectangle, oval, arrow, star, etc.) |
| `add_table` | Create tables with data |
| `add_chart` | Create charts (bar, column, line, pie, area) |
| `insert_icon` | Insert Phosphor SVG icons |
| `list_icons` | List available icons by category |

### Modifications
| Tool | Description |
|------|-------------|
| `modify_shape` | Change position, size, color, text of shapes |
| `delete_shape` | Remove shapes by ID or name |
| `find_and_replace` | Find and replace text across slides |

### Advanced
| Tool | Description |
|------|-------------|
| `evaluate_code` | Execute arbitrary python-pptx code |

## Example Workflow

```
User: Create a presentation about Q4 results

Claude:
1. manage_presentation(action="create", file_path="Q4_Results.pptx")
2. add_textbox(slide_number=1, text="Q4 2024 Results", font_size=44, ...)
3. manage_slide(action="add")
4. add_chart(slide_number=2, chart_type="bar", categories='["Q1","Q2","Q3","Q4"]', ...)
5. manage_presentation(action="save")
```

## Working with Templates

This MCP server works best when paired with a well-prepared PowerPoint template. Here's the recommended workflow based on testing:

### Quick Start (No Template Prep)

1. Download a template from PowerPoint's built-in library (e.g., "Architecture Pitch Deck")
2. Save it to your working folder
3. Ask Claude Code to create a presentation using the template

**Results:** Works fairly well out of the box – Claude Code can identify placeholders and populate content.

### Better Results (Templatized)

For improved results, ask Claude Code to prepare the template first:

1. Save a PowerPoint template to your working folder
2. Ask Claude Code to: *"Strip the existing text and replace it with generic placeholder text to make this a reusable template"*
3. Make your own personal tweaks to master layouts and example slides (e.g., icon placement, text box size)
4. Use the prepared template for your presentations

**Results:** More consistent placeholder detection, cleaner content replacement, and better overall output.

### Recommended Next Steps

Once you have a workflow that works well for your use case:

1. **Create a custom template** – Design or modify a template that matches your branding/style guidelines
2. **Package into a skill** – Create a Claude Code skill that encodes your storytelling and branding guidelines, so Claude follows them automatically when creating presentations
3. **Optional: Add a hook** – Configure a hook to load your skill before the MCP tools are invoked, ensuring consistent results every time

This approach gets closer to "one-shotting" powerpoint slides consistently.

## Notes

- All positions and sizes use **inches** as the unit
- Colors are specified as hex codes (e.g., "#FF0000" for red)
- Icons are bundled from Phosphor Icons (1,500+ icons available)
- The server keeps one presentation open in memory at a time
