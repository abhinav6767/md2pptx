# MD2PPTX — System Architecture & Workflow Documentation

> A complete technical reference for the multi-agent AI presentation generation system.

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [Full Pipeline Walkthrough](#2-full-pipeline-walkthrough)
3. [Agent Descriptions](#3-agent-descriptions)
4. [Visual Type Decision Logic](#4-visual-type-decision-logic)
5. [Template Color System](#5-template-color-system)
6. [Rendering Engine](#6-rendering-engine)
7. [Native Infographic Generators](#7-native-infographic-generators)
8. [Error Handling & Fallbacks](#8-error-handling--fallbacks)
9. [Configuration & Environment](#9-configuration--environment)
10. [Data Models](#10-data-models)

---

## 1. System Overview

MD2PPTX is a **multi-agent AI pipeline** that transforms unstructured Markdown documents into professionally designed PowerPoint presentations. The key insight is decomposing the problem:

```
"Given a 10-page research document, make me a stunning 12-slide consultant deck"
```

...into a sequence of **narrow, verifiable AI tasks**, each handled by a specialized agent.

### Design Philosophy

| Principle | Implementation |
|-----------|---------------|
| **Native over AI Images** | Consultant-style infographics built with `python-pptx` native shapes |
| **Template Fidelity** | Full 6-accent theme color extraction from Slide Master XML |
| **Zero Empty Slides** | Mandatory title + subtitle fallback injection on every slide |
| **Layout Integrity** | Cover, Thank You, Divider layouts protected — never modified |
| **Narrow AI Scopes** | Each agent has one job; failures don't cascade |

---

## 2. Full Pipeline Walkthrough

```
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 1: INPUT PARSING                                             │
│  File: parser.py                                                    │
│                                                                     │
│  Input: report.md                                                   │
│    ↓                                                                │
│  MarkdownParser reads the file section by section                   │
│  Extracts:                                                          │
│    • Document title (first H1)                                      │
│    • Subtitle (first paragraph after H1)                            │
│    • Sections (H2 headings, each with sub-content)                  │
│    • ContentBlocks: HEADING, PARAGRAPH, BULLET_LIST, TABLE, CODE    │
│    • TableData objects with structured rows/columns                 │
│  Output: Document object                                            │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 2: CONTENT ANALYSIS                                          │
│  File: analyzer.py                                                  │
│                                                                     │
│  Scans all sections for:                                            │
│    • ChartCandidates: tables with numerical data → bar/line/pie     │
│    • InfographicCandidates: lists with steps/phases/metrics         │
│    • KeyMetrics: large numbers, %, $, growth figures                │
│  Output: AnalysisResult (passed alongside Document to agents)       │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 3A: TEMPLATE LOADING                                         │
│  File: template.py                                                  │
│                                                                     │
│  Reads the Slide Master .pptx template:                             │
│    • Enumerates all layout names (Cover, Title only, Blank, etc.)   │
│    • Extracts full 6-accent theme color palette from XML:           │
│      dk1 → primary    dk2 → secondary    lt1 → background          │
│      accent1 → accent1  accent2 → accent2  ... accent6 → accent6   │
│    • Detects text brightness → assigns text_dark color              │
│    • Identifies protected layouts (Cover, Thank You, Divider)       │
│  Output: TemplateInfo containing DesignSystem                       │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 3B: MULTI-AGENT ORCHESTRATION                                │
│  File: agents/orchestrator.py                                       │
│                                                                     │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │  Agent 1: StorylineAgent (agents/storyline_agent.py)         │   │
│  │  Model: Gemini 2.5 Flash | Temperature: 0.2                  │   │
│  │                                                              │   │
│  │  Input:  Compressed document text + target slide count       │   │
│  │  Output: AgentStoryline — list of AgentSlide objects          │   │
│  │                                                              │   │
│  │  Each AgentSlide includes:                                   │   │
│  │   • title: string                                            │   │
│  │   • subtitle: string (mandatory — empty never allowed)       │   │
│  │   • bullets: list[str] (max 6, no emojis)                   │   │
│  │   • recommended_visual: one of 9 visual type keys           │   │
│  │   • visual_reference: table name for chart slides           │   │
│  │   • is_two_column: bool                                      │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                              ↓                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │  Agent 2: VisualizerAgent (inlined in orchestrator.py)       │   │
│  │  Model: Gemini 2.5 Flash | Temperature: 0.3                  │   │
│  │                                                              │   │
│  │  Validates and refines visual type assignments               │   │
│  │  Maps content patterns to preferred visual types             │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                              ↓                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │  Agent 3: ImageAgent (agents/image_agent.py)                 │   │
│  │  APIs: Unsplash API → Gemini fallback                        │   │
│  │                                                              │   │
│  │  For slides with visual = image/hero_header/sidebar_split:   │   │
│  │   1. Builds a contextual search query from slide title       │   │
│  │   2. Queries Unsplash for a real photograph                  │   │
│  │   3. Falls back to a placeholder if Unsplash returns nothing │   │
│  │   4. Saves image to /tmp and stores path in image_url        │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                              ↓                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │  Agent 4: LayoutAgent (agents/layout_agent.py)               │   │
│  │  Model: Gemini 2.5 Flash | Temperature: 0.0                  │   │
│  │                                                              │   │
│  │  Maps each slide to an EXACT Slide Master layout name        │   │
│  │  Rules:                                                      │   │
│  │   • First slide → Cover/1_Cover layout                       │   │
│  │   • Last slide → Thank You layout                            │   │
│  │   • Never assign Blank/Thank You to content slides           │   │
│  │   • Content slides → Title only or 1_E_Title variants        │   │
│  └──────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│  STAGE 4: RENDERING                                                 │
│  File: renderer.py                                                  │
│                                                                     │
│  For each slide in AgentStoryline:                                  │
│    1. Add slide using the assigned Slide Master layout              │
│    2. Fill native placeholders (idx 0/10 = title, 11 = subtitle)   │
│    3. If no title placeholder found → inject textbox fallback       │
│    4. If subtitle exists and not in placeholder → inject paragraph  │
│    5. Check is_protected_layout → skip images if Cover/TY/Divider  │
│    6. Route to visual generator based on recommended_visual         │
│    7. Add slide number footer                                       │
│                                                                     │
│  Visual Routing:                                                    │
│    numbered_list    → infographic_gen._generate_numbered_list()     │
│    swimlane         → infographic_gen._generate_swimlane()          │
│    icon_cards       → infographic_gen._generate_icon_cards()        │
│    data_table       → infographic_gen._generate_data_table()        │
│    vertical_timeline → infographic_gen._generate_vertical_timeline()│
│    chart            → chart_gen.generate() → PNG → add_picture()   │
│    image            → slide.shapes.add_picture() at BODY_TOP       │
│    hero_header      → full image + oval overlay + premium_cards     │
│    sidebar_split    → primary panel + infographic on right          │
│    text             → _add_bullets() with formatted runs            │
│                                                                     │
│  Output: Saved .pptx file                                           │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 3. Agent Descriptions

### StorylineAgent (`agents/storyline_agent.py`)

**Purpose:** The "creative director." Plans the entire slide narrative from scratch.

**Inputs:** Compressed document text (first 40,000 chars), target slide count

**Outputs:** `AgentStoryline` — a validated Pydantic model containing 12 `AgentSlide` objects

**Key Prompt Instructions:**
- Never use `text`-only slides — every slide must have a visual type
- Bullets must follow "Header: Detail" format for native layouts
- `data_table` bullets must use pipe format: `"Indicator | Value | Year | Source"`
- Subtitles are mandatory on every slide

---

### LayoutAgent (`agents/layout_agent.py`)

**Purpose:** Maps each slide to the precise layout name from the Slide Master.

**Inputs:** List of available layout names from template, AgentStoryline

**Outputs:** Updated AgentStoryline with `layout_name` populated for each slide

**Key Rules:**
- Slide 1 → Cover pattern layouts
- Slide N (last) → Thank You pattern layouts
- Content slides → NEVER Blank, NEVER Thank You

---

### ChartAgent (`agents/chart_agent.py`)

**Purpose:** Generates AI-written Python code to produce a styled matplotlib chart.

**Inputs:** Table data (headers + rows), chart type, theme hex colors

**Process:**
1. Builds a detailed prompt with data embedded as `data = {...}`
2. Injects guaranteed boilerplate (imports, font settings, DejaVu Sans override)
3. Strips all non-ASCII characters (emoji/glyphs) from data
4. Executes generated code with `exec()` in an isolated namespace
5. Returns the saved PNG path

---

### InfographicAgent (`agents/infographic_agent.py`)

**Purpose:** Generates AI-written Python code for custom matplotlib infographics.

**Note:** Only invoked for `infographic` type slides not covered by the 5 native generators. Most slides now use native shapes instead.

---

### ImageAgent (`agents/image_agent.py`)

**Purpose:** Sources contextually relevant photographs for visual slides.

**Process:**
1. Constructs a search query from slide title + keywords
2. Calls Unsplash API (`/photos/random?query=...`)
3. Downloads the image to a temp directory
4. Returns the local file path (or empty string on failure)

---

## 4. Visual Type Decision Logic

The StorylineAgent decides the visual type using this logic:

```
Content Pattern                    → Visual Type
──────────────────────────────────────────────────
3-6 mechanisms/factors/risks       → numbered_list
Parallel pillars / N-step approach → swimlane
4-6 category:description items     → icon_cards
Stats with value/year/source       → data_table
Chronological phases               → vertical_timeline
Numerical table exists             → chart
Conceptual (needs photo)           → image
Executive Summary / Agenda         → hero_header
Deep-dive / analytical             → sidebar_split
```

---

## 5. Template Color System

Colors are extracted from Slide Master theme XML in `template.py`:

```python
# XML path: //theme/themeElements/clrScheme
dk1  → colors.primary        (brand color, used for circles, headers, accents)
dk2  → colors.secondary      (supporting dark color)
lt1  → colors.background     (slide background)
ac1  → colors.accent1        (first supporting color)
ac2  → colors.accent2        (second supporting color)
ac3  → colors.accent3
ac4  → colors.accent4
ac5  → colors.accent5
```

Both `srgbClr` and `sysClr` (with `lastClr` fallback) are handled.

**Text Color:**
- If `dk1` brightness < 128 → `text_dark = dk1` (e.g., dark navy)
- Otherwise → `text_dark = #1A1A2E` (safe fallback)

**Text on primary backgrounds (sidebar_split, numbered_list circles):**
Always uses `text_light = #FFFFFF` regardless of theme.

---

## 6. Rendering Engine

### Grid System (`design.py`)

All shape placements snap to predefined grid constants:

```python
SLIDE_WIDTH   = Inches(13.33)   SLIDE_HEIGHT  = Inches(7.50)
MARGIN_LEFT   = Inches(0.5)     MARGIN_RIGHT  = Inches(0.5)
TITLE_LEFT    = Inches(0.4)     TITLE_TOP     = Inches(0.55)
SUBTITLE_TOP  = Inches(1.1)
BODY_LEFT     = Inches(0.5)     BODY_TOP      = Inches(1.6)
BODY_WIDTH    = Inches(12.3)    BODY_HEIGHT   = Inches(5.2)
COL_LEFT_LEFT = Inches(0.5)     COL_LEFT_WIDTH = Inches(5.9)
COL_RIGHT_LEFT= Inches(6.7)     COL_RIGHT_WIDTH= Inches(5.9)
```

### Layout Protection

```python
is_protected_layout = any(p in layout_name_lower for p in [
    "cover", "thank you", "thank_you", "divider", "section"
])
# If True: no images injected, native template design preserved
```

### Title Fallback

```python
# After iterating placeholders:
if not title_set:
    self._add_title_shape(slide, slide_content.title)
    title_set = True
```

---

## 7. Native Infographic Generators

All located in `infographics.py`. Each accepts `bounds=(left, top, width, height)`:

### `_generate_numbered_list`
- Vertical rows: colored circle (01, 02, ...) on left + content text on right
- Splits "Header: Detail" format → bold header + muted detail text
- Thin separator line between rows
- Color cycle: primary → secondary → accent1 → ...

### `_generate_swimlane`
- Horizontal N-column layout
- Numbered badge at top → vertical connector → content card
- Card: light `#F7F8FA` background + colored top accent bar
- Supports up to 5 columns

### `_generate_icon_cards`
- 2-column grid (for > 3 items)
- Rounded rectangle cards with left vertical accent bar
- Shadow effect via `card.shadow.visible = True`
- Splits "Category: Description" format

### `_generate_data_table`
- Primary-colored header row + alternating `#F2F6FF` / white body rows
- Column widths: 42% / 20% / 15% / 23%
- Bold + primary-colored "Value" column

### `_generate_vertical_timeline`
- Left rail (Inches 0.5 wide, primary color)
- Colored dots at each entry
- Inline bold header + regular body text using `add_run()`

---

## 8. Error Handling & Fallbacks

| Stage | Error | Fallback |
|-------|-------|----------|
| StorylineAgent API timeout | Raises `ValueError` | Retry up to 1x |
| ImageAgent Unsplash failure | Returns empty string | Image slot skipped |
| ChartAgent bad generated code | `exec()` exception caught | Falls back to bullet text |
| InfographicAgent AI failure | Exception caught with warning | Falls back to `_generate_process_flow` |
| Layout not found in master | Falls back to index 0 | First available layout |
| No title placeholder | `_add_title_shape()` always called | Manual textbox at TITLE_TOP |
| Non-ASCII glyphs in data | `re.sub(r'[^\x00-\x7F]+', '', ...)` | Cleaned input sent to AI |
| Matplotlib missing font | `DejaVu Sans` forced in boilerplate | System sans-serif |

---

## 9. Configuration & Environment

### `.env` file (in `md2pptx/`)

```env
GOOGLE_API_KEY=<your_google_ai_studio_api_key>
UNSPLASH_ACCESS_KEY=<your_unsplash_access_key>   # optional
```

### `design.py` — Adjustable Constants

| Constant | Default | Purpose |
|----------|---------|---------|
| `MAX_BULLETS_PER_SLIDE` | 6 | Caps bullet count per slide |
| `MAX_SLIDES` | 15 | Maximum slide count |
| `MIN_SLIDES` | 10 | Minimum slide count |
| `FONT_TITLE` | 28pt | Title font size |
| `FONT_SUBTITLE` | 16pt | Subtitle font size |
| `FONT_BODY` | 14pt | Body text font size |

---

## 10. Data Models

### `AgentSlide` (agents/models.py)

```python
class AgentSlide(BaseModel):
    title: str                        # Primary slide message
    subtitle: Optional[str]           # One-sentence context (mandatory in prompts)
    bullets: List[str]                # Supporting points (max 6, no emojis)
    speaker_notes: Optional[str]      # Extended context for presenter
    recommended_visual: str           # One of 9 visual type keys
    visual_reference: Optional[str]   # Table name for chart slides
    layout_name: Optional[str]        # Exact Slide Master layout name
    is_two_column: bool               # 50/50 split text layout
    image_url: Optional[str]          # Path to downloaded/generated image
```

### `InfographicCandidate` (analyzer.py)

```python
@dataclass
class InfographicCandidate:
    infographic_type: str    # process, timeline, numbered_list, swimlane, etc.
    title: str               # Slide title (for context)
    items: List[str]         # Bullet strings (often "Header: Detail" format)
    values: List[str]        # Optional numerical values (for metrics type)
    section_title: str       # Source section name
```

---

*Generated: April 2026 | MD2PPTX v2.0 — Multi-Agent Premium Presentation Engine*
