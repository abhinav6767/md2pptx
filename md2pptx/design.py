"""
Design System — Constants for colors, fonts, spacing, and grid layout.
Extracts theme from Slide Master and provides unified design tokens.
"""

from dataclasses import dataclass, field
from typing import List, Tuple, Optional
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


# Slide dimensions (widescreen 13.33 x 7.50 inches)
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# Margins
MARGIN_LEFT = Inches(0.5)
MARGIN_RIGHT = Inches(0.5)
MARGIN_TOP = Inches(0.7)
MARGIN_BOTTOM = Inches(0.5)

# Content area
CONTENT_LEFT = MARGIN_LEFT
CONTENT_TOP = Inches(1.4)  # Below title area
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN_BOTTOM

# Title area
TITLE_LEFT = Inches(0.4)
TITLE_TOP = Inches(0.55)
TITLE_WIDTH = Inches(10.6)
TITLE_HEIGHT = Inches(0.6)

# Subtitle area
SUBTITLE_LEFT = TITLE_LEFT
SUBTITLE_TOP = Inches(1.1)
SUBTITLE_WIDTH = TITLE_WIDTH
SUBTITLE_HEIGHT = Inches(0.3)

# Body area (for content slides)
BODY_LEFT = Inches(0.5)
BODY_TOP = Inches(1.6)
BODY_WIDTH = Inches(12.3)
BODY_HEIGHT = Inches(5.2)

# Two-column layout
COL_LEFT_LEFT = Inches(0.5)
COL_LEFT_WIDTH = Inches(5.9)
COL_RIGHT_LEFT = Inches(6.7)
COL_RIGHT_WIDTH = Inches(5.9)
COL_TOP = Inches(1.6)
COL_HEIGHT = Inches(5.2)

# Chart placement
CHART_LEFT = Inches(0.8)
CHART_TOP = Inches(1.8)
CHART_WIDTH = Inches(11.7)
CHART_HEIGHT = Inches(4.8)

# Footer / slide number area
FOOTER_LEFT = Inches(9.0)
FOOTER_TOP = Inches(6.6)
FOOTER_WIDTH = Inches(2.7)
FOOTER_HEIGHT = Inches(0.15)

# Typography scale
FONT_TITLE = Pt(28)
FONT_SUBTITLE = Pt(16)
FONT_BODY = Pt(14)
FONT_BODY_SMALL = Pt(12)
FONT_CAPTION = Pt(10)
FONT_LABEL = Pt(9)
FONT_SLIDE_NUMBER = Pt(8)

# Cover slide typography
FONT_COVER_TITLE = Pt(36)
FONT_COVER_SUBTITLE = Pt(18)

# Divider typography
FONT_DIVIDER_TITLE = Pt(32)

# Table typography
FONT_TABLE_HEADER = Pt(11)
FONT_TABLE_BODY = Pt(10)

# Max content limits
MAX_BULLETS_PER_SLIDE = 6
MAX_LINES_PER_SLIDE = 8
MAX_CHARS_PER_BULLET = 120
MAX_SLIDES = 15
MIN_SLIDES = 10

# Spacing
LINE_SPACING = Pt(20)
PARAGRAPH_SPACING = Pt(6)
BULLET_INDENT = Inches(0.25)


@dataclass
class ThemeColors:
    """Color palette extracted from Slide Master or defaults."""
    primary: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x3D, 0x6B))      # Dark blue
    secondary: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x7B, 0xC0))    # Medium blue
    accent1: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0xA3, 0xE0))      # Light blue
    accent2: RGBColor = field(default_factory=lambda: RGBColor(0x2E, 0xCC, 0x71))      # Green
    accent3: RGBColor = field(default_factory=lambda: RGBColor(0xF3, 0x97, 0x00))      # Orange
    accent4: RGBColor = field(default_factory=lambda: RGBColor(0xE7, 0x4C, 0x3C))      # Red
    accent5: RGBColor = field(default_factory=lambda: RGBColor(0x8E, 0x44, 0xAD))      # Purple
    text_dark: RGBColor = field(default_factory=lambda: RGBColor(0x33, 0x33, 0x33))
    text_light: RGBColor = field(default_factory=lambda: RGBColor(0xFF, 0xFF, 0xFF))
    text_muted: RGBColor = field(default_factory=lambda: RGBColor(0x77, 0x77, 0x77))
    background: RGBColor = field(default_factory=lambda: RGBColor(0xFF, 0xFF, 0xFF))
    divider_bg: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x3D, 0x6B))
    table_header_bg: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x3D, 0x6B))
    table_row_alt: RGBColor = field(default_factory=lambda: RGBColor(0xF2, 0xF2, 0xF2))

    def chart_colors(self) -> List[str]:
        """Return list of hex colors for matplotlib charts."""
        return [
            self._rgb_to_hex(self.primary),
            self._rgb_to_hex(self.secondary),
            self._rgb_to_hex(self.accent1),
            self._rgb_to_hex(self.accent2),
            self._rgb_to_hex(self.accent3),
            self._rgb_to_hex(self.accent4),
            self._rgb_to_hex(self.accent5),
        ]

    @staticmethod
    def _rgb_to_hex(color: RGBColor) -> str:
        # RGBColor.__str__() returns 'RRGGBB'
        hex_str = str(color)
        return f"#{hex_str.lower()}"


@dataclass
class ThemeFonts:
    """Font families from Slide Master."""
    heading: str = "Calibri"
    body: str = "Calibri"


@dataclass
class DesignSystem:
    """Complete design system combining colors, fonts, and layout rules."""
    colors: ThemeColors = field(default_factory=ThemeColors)
    fonts: ThemeFonts = field(default_factory=ThemeFonts)

    def get_bullet_char(self) -> str:
        return "•"

    def get_dash_char(self) -> str:
        return "–"
