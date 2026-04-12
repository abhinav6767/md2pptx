# MD2PPTX — AI-Powered Markdown to PowerPoint Engine

> 🏆 Hackathon Submission | Multi-Agent AI Presentation Generator

[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## What Is This?

**MD2PPTX** converts a plain Markdown (`.md`) file into a professional, consultant-grade PowerPoint presentation (`.pptx`) in under 2 minutes — fully automatically.

You provide:
1. A Markdown document (research report, analysis, notes)
2. A PowerPoint Slide Master template (your brand / theme)

The system outputs a complete, visually rich, 10–15 slide presentation that follows the color scheme, fonts, and layout system of your chosen template.

---

## ✨ Key Features

| Feature | Detail |
|---------|--------|
| 🧠 **Multi-Agent AI Pipeline** | 4 specialized Gemini agents collaborate to plan, visualize, and arrange content |
| 🎨 **Template-Faithful Colors** | Extracts all 6 theme accent colors from Slide Master XML — infographics match your brand exactly |
| 📊 **Smart Visual Selection** | AI decides whether each slide should be a chart, swimlane, numbered list, icon cards, or timeline |
| 🏗️ **Native PPTX Shapes** | Consultant-style infographics built with native PowerPoint shapes, not flat images |
| 🔒 **Layout Protection** | Cover, Thank You, and Divider layouts are never altered — template integrity preserved |
| ⚡ **Single CLI Command** | `python -m md2pptx input.md -t template.pptx -o output.pptx` |

---

## 🚀 Quick Start

### 1. Prerequisites
- Python 3.11+
- Google AI Studio API Key (Gemini)
- Unsplash API Key (optional, for images)

### 2. Install Dependencies

```bash
pip install -r md2pptx/requirements.txt
```

### 3. Configure API Keys

Create `md2pptx/.env`:

```env
GOOGLE_API_KEY=your_google_ai_studio_key_here
UNSPLASH_ACCESS_KEY=your_unsplash_key_here  # optional
```

### 4. Run

```bash
python -m md2pptx "your_report.md" -t "your_template.pptx" -o "output/result.pptx"
```

---

## 📋 CLI Reference

```
python -m md2pptx <input> -t <template> [-o <output>] [--slides N]
```

| Argument | Required | Default | Description |
|----------|----------|---------|-------------|
| `input` | ✅ | — | Path to `.md` input file |
| `-t` / `--template` | ✅ | — | Path to Slide Master `.pptx` template |
| `-o` / `--output` | ❌ | `output/<name>.pptx` | Output file path |
| `--slides` | ❌ | `12` | Target slide count (10–15) |

### Example Commands

```bash
# AI Bubble Analysis with Ghost Research template
python -m md2pptx "Code EZ_ Master of Agents _ Files\Sample Files\AI Bubble_ Detection, Prevention, and Investment Strategies\AI Bubble_ Detection, Prevention, and Investment Strategies.md" \
  -t "Code EZ_ Master of Agents _ Files\Sample Files\AI Bubble_ Detection, Prevention, and Investment Strategies\Template_AI Bubble_ Detection, Prevention, and Investment Strategies.pptx" \
  -o "output\AI_Bubble.pptx"

# Accenture Tech Acquisition Analysis
python -m md2pptx "Code EZ_ Master of Agents _ Files\Test Cases\Accenture Tech Acquisition Analysis.md" \
  -t "Code EZ_ Master of Agents _ Files\Slide Master\Template_Accenture Tech Acquisition Analysis.pptx" \
  -o "output\Accenture.pptx"
```

---

## 🏗️ System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        MD2PPTX PIPELINE                         │
└─────────────────────────────────────────────────────────────────┘

  INPUT: report.md + template.pptx
       │
       ▼
┌─────────────┐
│  1. PARSER  │  Converts Markdown into structured Document model
│  parser.py  │  Extracts: sections, headings, tables, bullet lists
└──────┬──────┘
       │
       ▼
┌──────────────┐
│  2. ANALYZER │  Scans content for chart-able tables and
│  analyzer.py │  infographic candidates (metrics, processes, etc.)
└──────┬───────┘
       │
       ▼
┌─────────────────────────────────────────────────────────────────┐
│                   3. MULTI-AGENT AI PIPELINE                    │
│                    (agents/ directory)                          │
│                                                                 │
│  ┌─────────────────┐   ┌─────────────────┐                     │
│  │ Storyline Agent │   │  Layout Agent   │                     │
│  │                 │   │                 │                     │
│  │ Gemini 2.5 Flash│   │ Gemini 2.5 Flash│                     │
│  │                 │   │                 │                     │
│  │ • Plans 12 slide│   │ • Maps each     │                     │
│  │   narrative arc │   │   slide to a    │                     │
│  │ • Assigns visual│   │   Slide Master  │                     │
│  │   type per slide│   │   layout name   │                     │
│  │ • Synthesizes   │→→→│ • Enforces no   │                     │
│  │   bullets from  │   │   Blank layouts │                     │
│  │   full document │   │   for content   │                     │
│  └─────────────────┘   └────────┬────────┘                     │
│                                 │                               │
│  ┌──────────────────┐  ┌────────▼────────┐                     │
│  │  Image Agent     │  │Visualizer Agent │                     │
│  │                  │  │                 │                     │
│  │ • Queries Unsplash│  │ • Maps content  │                     │
│  │   for contextual │  │   type to best  │                     │
│  │   photography    │  │   visual type   │                     │
│  │ • Falls back to  │←←│ • Schedules     │                     │
│  │   AI generation  │  │   charts &      │                     │
│  │                  │  │   infographics  │                     │
│  └──────────────────┘  └─────────────────┘                     │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌──────────────────────────────────────────────────────────────┐
│                    4. RENDERER (renderer.py)                 │
│                                                              │
│  For each slide:                                             │
│  • Loads correct Slide Master layout                         │
│  • Places title + subtitle text (with forced fallback)       │
│  • Routes to the appropriate visual generator:               │
│                                                              │
│    numbered_list → _generate_numbered_list()                 │
│    swimlane      → _generate_swimlane()                      │
│    icon_cards    → _generate_icon_cards()                    │
│    data_table    → _generate_data_table()                    │
│    vertical_timeline → _generate_vertical_timeline()         │
│    chart         → ChartGenerator → matplotlib → PNG        │
│    image         → ImageAgent → Unsplash PNG                 │
│    hero_header   → Full image + oval overlay + cards         │
│    sidebar_split → Primary panel + right infographic         │
│                                                              │
└──────────────────────────────┬───────────────────────────────┘
                               │
                               ▼
                    OUTPUT: presentation.pptx
```

---

## 🎨 Visual Type System

The **StorylineAgent** assigns each slide one of these visual types based on content:

| Visual Type | When Used | Native or AI? |
|-------------|-----------|---------------|
| `numbered_list` | 3–6 mechanisms, factors, key points | ✅ Native PPTX shapes |
| `swimlane` | Parallel pillars, multi-step processes | ✅ Native PPTX shapes |
| `icon_cards` | Categorised recommendations, safeguards | ✅ Native PPTX shapes |
| `data_table` | Quantitative indicators with source/year | ✅ Native PPTX shapes |
| `vertical_timeline` | Chronological phases, progressions | ✅ Native PPTX shapes |
| `chart` | Tabular numerical data (bar/line/pie) | 🤖 AI matplotlib → PNG |
| `image` | Conceptual slides needing illustration | 🌐 Unsplash API → PNG |
| `hero_header` | Executive Summary, Agenda overview | ✅ Native + Image |
| `sidebar_split` | Deep-dive analytical slides | ✅ Native + Image |

---

## 🔑 Key Design Decisions

### 1. Multi-Agent Architecture
Rather than a single monolithic LLM call, we split the problem into 4 specialized agents, each with a narrow scope. This reduces hallucination, makes each step testable, and allows retries on individual stages without regenerating everything.

### 2. Native PPTX Shapes Over Flat Images
Early versions used Matplotlib to generate all infographics as PNG images. This was replaced with **native `python-pptx` shape generation** for all structural layouts (numbered lists, swimlanes, cards etc.). Benefits: text remains editable, colors match the template exactly, and file sizes are smaller.

### 3. Template Color Extraction
The renderer reads all 6 accent colors (`accent1`–`accent6`), `dk1` (dark), `dk2` (secondary dark), and `lt1` (light) from the Slide Master's theme XML. This means the same code produces a red-branded output for Ghost Research and a navy-branded output for Accenture — automatically.

### 4. Protected Layouts
Cover, Thank You, and Divider slides are "protected" — no AI-generated images or shapes are injected onto them. This preserves the premium look of the template's native design for these key slides.

### 5. Guaranteed Title Fallback
If a layout's placeholder system doesn't expose a title slot (common in custom Slide Masters), the renderer manually injects a styled textbox at the title position. No slide is ever left without a visible title.

---

## 📁 Project Structure

```
md2pptx/
├── main.py                  # CLI entrypoint
├── parser.py                # Markdown → Document model
├── analyzer.py              # Content analysis & candidate detection
├── renderer.py              # PPTX rendering engine
├── template.py              # Slide Master color/font/layout extraction
├── design.py                # Grid constants, type scales, color system
├── charts.py                # Chart generation pipeline
├── infographics.py          # Native PPTX shape infographic generators
├── storyline.py             # Storyline-to-PPTX (legacy fallback)
├── requirements.txt         # Dependencies
└── agents/
    ├── orchestrator.py      # Agent pipeline coordinator
    ├── storyline_agent.py   # Slide narrative + visual type planning
    ├── layout_agent.py      # Slide Master layout mapping
    ├── chart_agent.py       # AI chart code generation (Gemini)
    ├── infographic_agent.py # AI infographic code generation (Gemini)
    ├── image_agent.py       # Image sourcing (Unsplash + fallback)
    ├── models.py            # Pydantic data models
    └── client.py            # Gemini API client setup
```

---

## 📦 Dependencies

```
python-pptx>=0.6.21   # PowerPoint file generation
matplotlib>=3.8.0     # Chart rendering
Pillow>=10.0.0        # Image handling
requests>=2.31.0      # HTTP (Unsplash API)
google-genai>=0.3.0   # Gemini AI (storyline, charts, infographics)
pydantic>=2.0.0       # Data validation for agent outputs
python-dotenv>=1.0.0  # Environment variable loading
```

---

## 📄 License

MIT — see [LICENSE](LICENSE) for details.
