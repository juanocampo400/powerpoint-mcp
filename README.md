# PowerPoint MCP Server

Create and edit PowerPoint presentations with Claude – locally, on any platform.

## Why This Exists

I wanted a PowerPoint MCP server that works on any machine, even with corporate IT restrictions.

Other servers I tried had:
- **External API calls** – Blocked by firewalls, require API keys, depend on third-party services
- **COM automation** – Windows-only, requires PowerPoint to be open, steals focus while working
- **Too many tools** – 30+ tools for animations and features I never use, slow to run

This server is:
- **100% local** – No external API calls, file operations happen locally through python-pptx
- **Cross-platform** – Works on Windows, macOS, and Linux
- **Non-intrusive** – Runs in background without launching windows that steal focus
- **Lightweight** – Lean toolset for 95% of use case

## Getting Started

### Prerequisites

**Python 3.8+** - Check with `python3 --version`

If not installed:
- macOS: `brew install python`
- Windows: install using `winget isntall Python.Python.3.13` or install from the Microsoft store (search "Python 3.13")
- Linux: use your package manager

### Install

#### 1. Clone the repo:
```bash
git clone https://github.com/juanocampo400/powerpoint-mcp.git
cd powerpoint-mcp
```

#### 2. Install dependencies:

macOS/Linux:
```bash
pip3 install -r requirements.txt
```
Windows:
```bash
pip install -r requirements.txt
```

#### 3. Add to Claude Code:

macOS:
```bash
chmod +x server.sh
claude mcp add powerpoint-mcp --scope user -- $PWD/server.sh
```

Windows (Git Bash):
```bash
claude mcp add powerpoint-mcp --scope user -- python $PWD/server.py
```

Linux:
```bash
claude mcp add powerpoint-mcp --scope user -- python3 "$PWD/server.py"
```
**Icon support (optional):**
- Windows: `pip install cairosvg` (works if pycairo installed, common with graphics/PDF tools)
- macOS: `brew install cairo pango && pip3 install cairosvg`
- Linux: Install Cairo for your distribution, then `pip3 install cairosvg`

<details>
<summary><strong>Manual configuration & platform/icons notes</strong></summary>

**Why `--scope user`?** Makes the server available globally. Without it, the server only works in the project directory.

**Why `server.sh` on macOS?** The wrapper script sets `DYLD_FALLBACK_LIBRARY_PATH` so Python can find the Homebrew-installed cairo library (required for icon support).

**Why Phosphor icons?** Fill-based SVGs (unlike Lucide's stroke-based SVGs) stay recolorable in PowerPoint. 1,000+ designs (vs Heroicons' ~300). MIT licensed.

**Manual config** – Edit `~/.claude.json`:

macOS:
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

Windows:
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
</details>

### Make Your First Deck

Open Claude Code and say:

> "Create a 5-slide presentation about [topic]"

Claude will create a .pptx file, feel free to specify file path too.

## Getting Better Results

This works out of the box works. But for consistent output, here's a suggested progression:

### Use a Template

Download an existing template or one from PowerPoint's library, save it to your working folder, and ask Claude to use it:

> "Create a presentation about [topic] using the template in my folder"

### Prepare the Template

For even better results, prep the template first or ask Claude to prep the template for you:

> "Strip the existing text from this template and replace it with generic placeholder text to make it reusable"

Then make your own tweaks to layouts, icon placement, text box sizes. Use this prepared template going forward.

### Create a Skill

Once you establish a workflow, package it into a Claude Code skill with your storytelling and branding guidelines. Claude will follow them automatically.

**One tester noted**: *"I made a skill for building pitch decks with the PowerPoint MCP and one of my analytics MCPs and they're working together well...I uploaded a template to the skill and it went wayyy faster. Under 1 minute."*


## Notes

- Positions and sizes use **inches**
- Colors are hex codes (e.g., `#FF0000` for red)
- 1,500+ Phosphor icons bundled
- One presentation open in memory at a time

## Available Tools

### Presentation & Slides
| Tool | Description |
|------|-------------|
| `manage_presentation` | Open, create, save, save_as, close presentations |
| `get_presentation_info` | Get slide count, dimensions, overview |
| `manage_slide` | Add, delete, duplicate, move slides |
| `get_slide_snapshot` | Get detailed info about shapes on a slide |

### Content Creation
| Tool | Description |
|------|-------------|
| `add_textbox` | Add text with formatting (font, size, color, alignment, bullets) |
| `add_image` | Insert images with fit modes (fill, fit, stretch) |
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
| `find_and_replace` | Find and replace text across slides (preserves formatting) |
| `get_table_content` | Get full table data (rows/columns) |
| `modify_table_cell` | Update individual table cells (preserves formatting) |

### Advanced
| Tool | Description |
|------|-------------|
| `evaluate_code` | Execute arbitrary python-pptx code for edge cases (aka the "escape hatch")|
