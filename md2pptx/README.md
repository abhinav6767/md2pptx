# MD2PPTX — Markdown to PowerPoint Converter

A modular, CLI-based Python system that converts Markdown (.md) files into structured, visually appealing .pptx presentations using provided Slide Master templates.

## Features

- **Automatic Content Parsing**: Parses Markdown headings, bullet lists, numbered lists, tables, and paragraphs into a structured document model
- **Intelligent Storyline Building**: Automatically constructs a 10–15 slide narrative following the flow: Title → Agenda → Executive Summary → Content → Charts → Conclusion → Thank You
- **Dynamic Chart Generation**: Detects numerical tables and generates professional matplotlib charts (bar, line, pie, area) styled to match the template's color scheme
- **Infographic Generation**: Creates process flows, timelines, and comparison layouts using native PowerPoint shapes
- **Slide Master Compliance**: All slides inherit the provided Slide Master's layouts, themes, colors, and fonts
- **Cross-Platform Compatibility**: Output opens correctly in Microsoft PowerPoint, Google Slides, and LibreOffice Impress
- **Content Summarization**: Handles large documents (up to 5 MB) by intelligently selecting and summarizing the most important content

## Architecture

```
md2pptx/
├── main.py            # CLI entrypoint & pipeline orchestrator
├── parser.py          # Markdown → structured Document model
├── analyzer.py        # Content analysis (tables, metrics, infographics)
├── storyline.py       # 10-15 slide storyline builder
├── charts.py          # Matplotlib chart generation
├── infographics.py    # PowerPoint shape-based infographics
├── renderer.py        # PPTX rendering engine
├── template.py        # Slide Master loader & analyzer
├── design.py          # Design system constants
└── requirements.txt   # Dependencies
```

### Pipeline Flow

```
Markdown → Parser → Analyzer → Storyline Builder → Renderer → .pptx
                                                      ↑
                                            Chart Generator
                                            Infographic Generator
                                            Template Loader
```

## Installation

```bash
pip install -r md2pptx/requirements.txt
```

### Dependencies
- `python-pptx` — PowerPoint file generation
- `matplotlib` — Chart generation
- `Pillow` — Image handling
- `markdown-it-py` — Markdown parsing support

## Usage

```bash
# Basic usage
python -m md2pptx input.md -t template.pptx -o output.pptx

# With custom slide count
python -m md2pptx input.md -t template.pptx --slides 15

# Auto-name output (saved to output/ folder)
python -m md2pptx input.md -t template.pptx
```

### CLI Arguments

| Argument | Required | Default | Description |
|----------|----------|---------|-------------|
| `input` | Yes | — | Path to the Markdown (.md) input file |
| `-t`, `--template` | Yes | — | Path to the Slide Master template (.pptx) |
| `-o`, `--output` | No | `output/<input_name>.pptx` | Output file path |
| `--slides` | No | `12` | Target number of slides (10–15) |

## Examples

```bash
# Accenture Analysis
python -m md2pptx "Test Cases/Accenture Tech Acquisition Analysis.md" \
  -t "Slide Master/Template_Accenture.pptx" \
  -o "output/Accenture.pptx"

# Banking Analysis
python -m md2pptx "Test Cases/Banking ROE.md" \
  -t "Slide Master/Template_Accenture.pptx" \
  -o "output/Banking.pptx"
```

## Slide Types

The system generates the following slide types, adapted to the template's available layouts:

| Slide Type | Layout Used | Content |
|-----------|-------------|---------|
| Cover | Cover/1_Cover | Document title + subtitle |
| Agenda | Title only | Numbered list of topics |
| Executive Summary | Title only | Key bullet points |
| Content | Title only | Section bullets |
| Two-Column | Title only | Split content with divider |
| Table | Title only | Formatted data table |
| Chart | Blank | Matplotlib chart image |
| Infographic | Blank | Shape-based process flow |
| Conclusion | Title only | Key takeaways |
| Thank You | Thank You | End slide |

## Design System

- **Typography**: Title (28pt), Subtitle (16pt), Body (14pt), Caption (10pt)
- **Grid**: 13.33" × 7.50" slides with 0.5" margins
- **Colors**: Extracted from Slide Master theme or defaults (professional blue palette)
- **Charts**: Styled with theme colors, clean axis formatting, value labels
- **Tables**: Alternating row colors, themed header row

## Limitations

- Base64-embedded images in markdown are skipped (charts are generated from table data instead)
- Content summarization uses rule-based heuristics (no LLM required)
- Maximum recommended input file size: 5 MB
- Slide count range: 10–15 slides

## License

MIT
