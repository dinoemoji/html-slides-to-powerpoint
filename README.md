# HTML to PowerPoint Converter

A Python tool that converts HTML slides (stored in JSON format) to PowerPoint (.pptx) presentations. This tool uses Playwright to render HTML and extract visual elements, then recreates them in PowerPoint format.

## Features

- Converts HTML slides from JSON to PowerPoint format
- Preserves text formatting, colors, fonts, and styles
- Handles images (both local files and URLs)
- Supports shapes, backgrounds, borders, and rounded corners
- Maintains table structures with proper formatting
- Preserves styled text elements (chips/pills with backgrounds)

## Requirements

- Python 3.7 or higher
- Playwright browser binaries (installed automatically)

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Install Playwright browsers:
```bash
playwright install chromium
```

## Usage

### Basic Usage

Convert a JSON file containing HTML slides to PowerPoint:

```bash
python3 html_to_pptx.py input.json [output.pptx]
```

- `input.json`: Path to the JSON file containing HTML slides
- `output.pptx`: (Optional) Output PowerPoint file path. If not specified, defaults to `input.pptx`

### JSON Format

The input JSON file should be an array of slide objects, where each slide has:

```json
[
  {
    "id": "slide_1",
    "html": "<!DOCTYPE html>..."
  },
  {
    "id": "slide_2",
    "html": "<!DOCTYPE html>..."
  }
]
```

Each slide object should contain:
- `id`: A unique identifier for the slide
- `html`: The HTML content of the slide (1920x1080 pixels recommended)

### Example

```bash
# Convert with default output name (ms_preso.pptx)
python3 html_to_pptx.py ms_preso.json

# Convert with custom output name
python3 html_to_pptx.py ms_preso.json my_presentation.pptx
```

## How It Works

1. **HTML Rendering**: Uses Playwright to render each HTML slide in a Chromium browser
2. **Element Extraction**: Extracts text, images, shapes, and styling information from the rendered HTML
3. **PowerPoint Creation**: Creates PowerPoint slides using python-pptx, positioning and styling elements to match the HTML

## Slide Dimensions

- Default slide size: 1920x1080 pixels (16:9 aspect ratio)
- PowerPoint slide size: Automatically converted to inches (96 DPI)

## Supported Elements

- **Text**: Headings, paragraphs, styled text with fonts, colors, and alignment
- **Images**: Local files and HTTP/HTTPS URLs with aspect ratio preservation
- **Shapes**: Rectangles, circles, rounded rectangles with fills and borders
- **Tables**: Table structures with cell formatting, borders, and backgrounds
- **Backgrounds**: Solid colors and gradients

## Notes

- The tool processes slides sequentially and may take some time for presentations with many slides
- Images are downloaded from URLs if needed
- Font families are mapped to PowerPoint-compatible fonts (e.g., "Proxima Nova" → "Calibri")
- Circular images and shapes are detected and preserved

## Troubleshooting

### Playwright Browser Not Found
If you see errors about missing browsers, run:
```bash
playwright install chromium
```

### Image Loading Issues
- Ensure image URLs are accessible
- Local image paths should be relative to the HTML file or absolute paths
- Check that image files exist if using local paths

### Font Issues
- Web fonts are mapped to system fonts available in PowerPoint
- Custom fonts may not be available and will fall back to similar system fonts

## Related Tools

This project also includes:
- `pptx_to_json.py`: Converts PowerPoint files to JSON with HTML representation
- `pptx_to_json_relevance.py`: Enhanced version with robust image extraction

