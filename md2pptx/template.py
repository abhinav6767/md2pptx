"""
Template Loader — Loads and analyzes Slide Master .pptx templates.
Extracts layouts, placeholders, theme colors, and fonts.
"""

import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from .design import ThemeColors, ThemeFonts, DesignSystem


@dataclass
class LayoutInfo:
    """Information about a slide layout."""
    name: str
    index: int
    placeholders: Dict[int, dict] = field(default_factory=dict)


@dataclass
class TemplateInfo:
    """Parsed template information."""
    path: str
    slide_width: int = 0
    slide_height: int = 0
    layouts: List[LayoutInfo] = field(default_factory=list)
    layout_map: Dict[str, int] = field(default_factory=dict)
    design: DesignSystem = field(default_factory=DesignSystem)

    # Mapped layout indices for each slide type
    cover_layout_idx: int = 0
    divider_layout_idx: int = -1
    content_layout_idx: int = -1
    blank_layout_idx: int = -1
    thankyou_layout_idx: int = -1


class TemplateLoader:
    """Loads and analyzes a Slide Master .pptx template."""

    # Known layout name patterns for each slide type
    COVER_PATTERNS = ["cover", "1_cover", "2_cover", "0_title", "title company"]
    DIVIDER_PATTERNS = ["divider", "section", "c_section"]
    CONTENT_PATTERNS = ["title only", "title, subtitle", "1_e_title"]
    BLANK_PATTERNS = ["blank"]
    THANKYOU_PATTERNS = ["thank you", "thank_you", "1_thank"]

    def __init__(self, template_path: str):
        self.template_path = template_path
        self.prs = Presentation(template_path)
        self.info = self._analyze()

    def _analyze(self) -> TemplateInfo:
        """Analyze the template and extract all relevant information."""
        info = TemplateInfo(path=self.template_path)
        info.slide_width = self.prs.slide_width
        info.slide_height = self.prs.slide_height

        # Extract layouts
        for sm in self.prs.slide_masters:
            for idx, layout in enumerate(sm.slide_layouts):
                layout_info = LayoutInfo(name=layout.name, index=idx)

                for ph in layout.placeholders:
                    layout_info.placeholders[ph.placeholder_format.idx] = {
                        "type": str(ph.placeholder_format.type),
                        "name": ph.name,
                        "left": ph.left,
                        "top": ph.top,
                        "width": ph.width,
                        "height": ph.height,
                    }

                info.layouts.append(layout_info)
                info.layout_map[layout.name.lower()] = idx

        # Map layout types
        info.cover_layout_idx = self._find_layout(info, self.COVER_PATTERNS, default=0)
        info.divider_layout_idx = self._find_layout(info, self.DIVIDER_PATTERNS, default=-1)
        info.content_layout_idx = self._find_layout(info, self.CONTENT_PATTERNS, default=-1)
        info.blank_layout_idx = self._find_layout(info, self.BLANK_PATTERNS, default=-1)
        info.thankyou_layout_idx = self._find_layout(info, self.THANKYOU_PATTERNS, default=-1)

        # If no content layout, fallback to blank
        if info.content_layout_idx == -1:
            info.content_layout_idx = info.blank_layout_idx

        # Extract theme
        info.design = self._extract_design()

        return info

    def _find_layout(self, info: TemplateInfo, patterns: List[str], default: int = -1) -> int:
        """Find a layout index matching any of the given name patterns."""
        for name_lower, idx in info.layout_map.items():
            for pattern in patterns:
                if pattern in name_lower:
                    return idx
        return default

    def _extract_design(self) -> DesignSystem:
        """Extract design tokens from the template's theme."""
        colors = ThemeColors()
        fonts = ThemeFonts()

        # Try extracting theme colors from the slide master
        try:
            sm = self.prs.slide_masters[0]
            theme = sm.element
            # Extract from theme XML
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            theme_elements = theme.findall(".//a:theme/a:themeElements/a:clrScheme", ns)

            if theme_elements:
                clr = theme_elements[0]
                dk1 = clr.find("a:dk1", ns)
                dk2 = clr.find("a:dk2", ns)
                accent1 = clr.find("a:accent1", ns)
                accent2 = clr.find("a:accent2", ns)

                # Try to extract srgbClr values
                for elem, attr_name in [
                    (dk1, "primary"), (accent1, "accent1"), (accent2, "accent2")
                ]:
                    if elem is not None:
                        srgb = elem.find("a:srgbClr", ns)
                        if srgb is not None:
                            val = srgb.get("val", "")
                            if len(val) == 6:
                                r = int(val[0:2], 16)
                                g = int(val[2:4], 16)
                                b = int(val[4:6], 16)
                                setattr(colors, attr_name, RGBColor(r, g, b))
        except Exception:
            pass  # Use defaults if theme extraction fails

        # Try extracting fonts
        try:
            sm = self.prs.slide_masters[0]
            # Check first slide for font info
            for slide in self.prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.name:
                                    fonts.heading = run.font.name
                                    fonts.body = run.font.name
                                    break
                            if fonts.heading != "Calibri":
                                break
                    if fonts.heading != "Calibri":
                        break
                if fonts.heading != "Calibri":
                    break
        except Exception:
            pass  # Use defaults

        return DesignSystem(colors=colors, fonts=fonts)

    def get_presentation(self) -> Presentation:
        """Return a fresh Presentation based on the template."""
        return Presentation(self.template_path)

    def get_layout(self, layout_idx: int):
        """Get a specific slide layout by index."""
        sm = self.prs.slide_masters[0]
        if 0 <= layout_idx < len(sm.slide_layouts):
            return sm.slide_layouts[layout_idx]
        return sm.slide_layouts[0]


def find_templates(templates_dir: str) -> Dict[str, str]:
    """Find all available template .pptx files in a directory."""
    templates = {}
    if os.path.isdir(templates_dir):
        for fname in os.listdir(templates_dir):
            if fname.endswith(".pptx") and fname.startswith("Template_"):
                key = fname.replace("Template_", "").replace(".pptx", "").strip()
                # Create a short key
                short_key = key.split("_")[0].lower() if "_" in key else key[:20].lower()
                templates[short_key] = os.path.join(templates_dir, fname)
    return templates
