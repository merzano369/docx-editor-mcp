# docx-editor-mcp
A powerful MCP server for creating, editing, and extracting data from Microsoft Word documents.

# DOCX Editor MCP Server

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![MCP](https://img.shields.io/badge/MCP-Model%20Context%20Protocol-green)](https://modelcontextprotocol.io/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![python-docx](https://img.shields.io/badge/python--docx-latest-orange)](https://python-docx.readthedocs.io/)

A powerful **Model Context Protocol (MCP) server** for creating, editing, and extracting data from Microsoft Word documents (`.docx`). Designed for seamless integration with AI assistants like Claude Desktop, Kilocode, and other MCP-compatible clients.

## 📋 Table of Contents

- [Features](#-features)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Available Tools](#-available-tools)
- [Usage Examples](#-usage-examples)
- [Technology Stack](#-technology-stack)
- [Contributing](#-contributing)
- [License](#-license)

## ✨ Features

### Document Creation & Management
- **Create new documents** with pre-configured professional styles (Times New Roman, 14pt, justified alignment, 1.15 line spacing)
- **Load existing templates** for modification
- **Save documents** to any specified path

### Content Editing
- **Add headings** with automatic styling (centered, Times New Roman, 16pt)
- **Add paragraphs** with customizable alignment (LEFT, CENTER, RIGHT, JUSTIFY)
- **Insert formatted text** with bold, italic, custom font sizes, and language attributes
- **Create bullet and numbered lists** with automatic formatting

### Document Analysis & Extraction
- **Extract all document parameters** as structured JSON including:
  - Core properties (author, title, subject, keywords, created/modified dates)
  - Custom properties (user-defined metadata)
  - Document variables (for mail merge and automation)
  - Section properties (margins, page size, orientation)
  - Style definitions (fonts, colors, spacing, indentation)
  - Numbering and list definitions
  - Headers and footers content
  - Table structures

### Template Generation
- **Apply extracted parameters** to create new documents with identical formatting
- **Set core properties** programmatically (author, title, subject, etc.)
- **Set custom properties** for document metadata

### Document Structure Analysis
- Get comprehensive document structure summaries
- List all headings with levels and text
- Count paragraphs and tables
- Preview document content

## 📦 Installation

### Prerequisites
- Python 3.10 or higher
- pip package manager

### Quick Start

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/docx-editor-mcp.git
   cd docx-editor-mcp
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify installation:**
   ```bash
   python -c "from docx import Document; from mcp.server.fastmcp import FastMCP; print('Installation successful!')"
   ```

## ⚙️ Configuration

### Understanding MCP Servers

> **Important:** MCP servers communicate via stdio using JSON-RPC protocol. They are NOT meant to be run directly from the command line for interactive use. Instead, they must be launched by an MCP client.

### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`  
**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "python",
      "args": ["C:\\path\\to\\docx-editor-mcp\\server.py"]
    }
  }
}
```

### Kilocode / VS Code Configuration

Add to your VS Code settings or Kilocode configuration:

```json
{
  "mcp.servers": {
    "docx-editor": {
      "command": "python",
      "args": ["/path/to/docx-editor-mcp/server.py"]
    }
  }
}
```

## 🔧 Available Tools

### Document Creation & Management

| Tool | Description |
|------|-------------|
| [`create_document(filename)`](server.py:61) | Creates a new document with default professional styles |
| [`save_document(filename)`](server.py:182) | Saves the current document to specified path |
| [`load_template(filename)`](server.py:837) | Loads an existing document for modification |

### Content Addition

| Tool | Description |
|------|-------------|
| [`add_heading(text, level)`](server.py:83) | Adds a heading (Level 1-6) with default styling |
| [`add_heading_custom(text, level, font_size)`](server.py:98) | Adds a heading with custom font size |
| [`add_paragraph(text, alignment, indent_first_line)`](server.py:113) | Adds a paragraph with alignment options |
| [`add_formatted_text(paragraph_index, text, bold, italic, font_size, lang)`](server.py:141) | Appends styled text to a paragraph |
| [`add_list_item(text, style)`](server.py:170) | Adds bullet or numbered list item |

### Parameter Extraction

| Tool | Description |
|------|-------------|
| [`extract_document_parameters(filename, compact, all_styles)`](server.py:652) | Extract ALL document parameters as JSON |
| [`extract_core_properties(filename)`](server.py:700) | Extract metadata (author, title, dates, etc.) |
| [`extract_custom_properties(filename)`](server.py:726) | Extract user-defined custom properties |
| [`extract_document_variables(filename)`](server.py:752) | Extract document variables for automation |
| [`extract_section_properties(filename)`](server.py:778) | Extract margins, page size, orientation |
| [`extract_styles_info(filename, all_styles, compact)`](server.py:804) | Extract style definitions |
| [`get_document_structure(filename)`](server.py:1118) | Get headings, paragraphs, tables summary |

### Template Generation

| Tool | Description |
|------|-------------|
| [`apply_template_parameters(parameters_json, output_filename)`](server.py:974) | Create document from JSON parameters |
| [`set_core_property(property_name, value)`](server.py:1054) | Set metadata property |
| [`set_custom_property(property_name, value)`](server.py:1081) | Set custom property |

## 📖 Usage Examples

### Creating a New Document

Ask your AI assistant:

```
Create a new Word document called "report.docx" with:
- A heading "Annual Report 2024"
- A paragraph about company performance
- A bullet list with key achievements
```

The server will execute:
```python
create_document("report.docx")
add_heading("Annual Report 2024", level=1)
add_paragraph("The company has shown remarkable growth this year...")
add_list_item("Revenue increased by 25%", style="List Bullet")
add_list_item("Expanded to 3 new markets", style="List Bullet")
save_document()
```

### Extracting Document Parameters

```
Extract all parameters from "template.docx" and show me the styles used.
```

Returns structured JSON:
```json
{
  "core_properties": {
    "author": "John Doe",
    "title": "Company Template",
    "created": "2024-01-15T10:30:00"
  },
  "sections": [{
    "margins": {
      "top_mm": 15,
      "bottom_mm": 15,
      "left_mm": 20,
      "right_mm": 20
    },
    "orientation": "portrait"
  }],
  "styles": {
    "paragraph_styles": {
      "Normal": {
        "font": {"name": "Times New Roman", "size_pt": 14},
        "paragraph_format": {"alignment": "JUSTIFY", "line_spacing": 1.15}
      }
    }
  }
}
```

### Cloning a Document Template

```
Extract parameters from "template.docx" and create a new document "new_report.docx" with the same formatting.
```

```python
params = extract_document_parameters("template.docx")
apply_template_parameters(params, "new_report.docx")
# Now add your content...
add_heading("New Report", level=1)
add_paragraph("Your content here...")
save_document()
```

### Analyzing Document Structure

```
Load "document.docx" and show me its structure.
```

Returns:
```json
{
  "headings": [
    {"index": 0, "level": "Heading 1", "text": "Introduction"},
    {"index": 5, "level": "Heading 2", "text": "Methodology"}
  ],
  "paragraphs": [
    {"index": 1, "style": "Normal", "text_preview": "This document describes..."},
    {"index": 2, "style": "Normal", "text_preview": "The following sections..."}
  ],
  "tables_count": 2
}
```

## 🛠 Technology Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.10+ |
| **Protocol** | Model Context Protocol (MCP) |
| **MCP Framework** | FastMCP |
| **Document Engine** | python-docx |
| **Communication** | JSON-RPC over stdio |

### Dependencies

```
mcp>=1.0.0
python-docx>=0.8.11
```

### Default Document Styles

New documents are created with professional default styling:

| Element | Style |
|---------|-------|
| **Normal Text** | Times New Roman, 14pt, Justified, 1.15 line spacing |
| **Heading 1** | Times New Roman, 16pt, Centered, No bold |
| **Heading 2** | Times New Roman, 16pt, Centered, No bold |
| **Margins** | Top/Bottom: 15mm, Left/Right: 20mm |
| **First Line Indent** | 12.7mm (1.27 cm) |

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

### Getting Started

1. Fork the repository
2. Clone your fork:
   ```bash
   git clone https://github.com/yourusername/docx-editor-mcp.git
   ```
3. Create a feature branch:
   ```bash
   git checkout -b feature/amazing-feature
   ```
4. Make your changes and commit:
   ```bash
   git commit -m "Add amazing feature"
   ```
5. Push to your branch:
   ```bash
   git push origin feature/amazing-feature
   ```
6. Open a Pull Request

### Contribution Guidelines

- Follow PEP 8 style guidelines
- Add docstrings to all new functions
- Update documentation for any new features
- Test your changes thoroughly before submitting

### Feature Requests & Bug Reports

- Open an issue for bug reports or feature requests
- Provide detailed descriptions and reproduction steps for bugs
- Include examples for feature requests

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  Made with ❤️ for the MCP community
</p>