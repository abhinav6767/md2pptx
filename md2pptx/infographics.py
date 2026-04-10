"""
Infographic Generator — Creates programmatic infographics using python-pptx shapes.
Generates process flows, timelines, comparisons, and key metric callouts.
"""

from typing import List, Optional
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from .design import DesignSystem
from .analyzer import InfographicCandidate


class InfographicGenerator:
    """Generates infographic shapes on PPTX slides."""

    def __init__(self, design: DesignSystem):
        self.design = design

    def generate(self, slide, candidate: InfographicCandidate):
        """Generate an infographic on the given slide."""
        
        # 1) Attempt to use AI to generate highly unique infographic image
        try:
            from .agents.infographic_agent import InfographicAgent
            import tempfile
            import os
            from pptx.util import Inches

            out_dir = tempfile.mkdtemp(prefix="md2pptx_infographics_")
            filepath = os.path.join(out_dir, "info_ai.png")
            
            print(f"  [InfographicGenerator] Summoning AI Infographic Agent for stunning {candidate.infographic_type}...")
            agent = InfographicAgent()
            code = agent.generate_infographic_code(candidate, filepath)
            
            # Execute AI code
            local_vars = {"filepath": filepath}
            exec(code, {}, local_vars)
            
            if os.path.exists(filepath):
                # Put the generated AI infographic onto the slide
                slide.shapes.add_picture(
                    filepath,
                    Inches(0.5), Inches(2.0),
                    width=Inches(12.33),
                    height=Inches(5.0)
                )
                return
                
        except Exception as ai_e:
            print(f"  [Warning] AI Infographic generation failed ({ai_e}), falling back to native PPTX shapes.")

        # 2) Fallback to native PP shapes
        if candidate.infographic_type == "process":
            self._generate_process_flow(slide, candidate)
        elif candidate.infographic_type == "timeline":
            self._generate_timeline(slide, candidate)
        elif candidate.infographic_type == "comparison":
            self._generate_comparison(slide, candidate)
        elif candidate.infographic_type == "metrics":
            self._generate_metrics(slide, candidate)
        else:
            self._generate_process_flow(slide, candidate)

    def _generate_process_flow(self, slide, candidate: InfographicCandidate):
        """Generate a horizontal process flow infographic."""
        items = candidate.items
        if not items:
            return

        n = len(items)
        colors = self.design.colors

        # Layout calculations
        start_left = Inches(0.7)
        top = Inches(2.2)
        total_width = Inches(11.9)
        box_height = Inches(2.8)

        # Calculate individual box width and gap
        gap = Inches(0.15)
        available_width = total_width - gap * (n - 1)
        box_width = available_width / n

        # Color cycle for boxes
        box_colors = [
            colors.primary, colors.secondary, colors.accent1,
            colors.accent2, colors.accent3, colors.accent4,
        ]

        for i, item in enumerate(items):
            left = start_left + (box_width + gap) * i
            color = box_colors[i % len(box_colors)]

            # Main box with rounded corners
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(left), int(top), int(box_width), int(box_height)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = color
            shape.line.fill.background()

            # Step number circle
            num_size = Inches(0.5)
            num_left = left + (box_width - num_size) / 2
            num_top = top + Inches(0.2)
            num_shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                int(num_left), int(num_top), int(num_size), int(num_size)
            )
            num_shape.fill.solid()
            num_shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            num_shape.line.fill.background()

            # Number text
            tf = num_shape.text_frame
            tf.text = str(i + 1)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            tf.paragraphs[0].font.size = Pt(14)
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.color.rgb = color
            tf.word_wrap = True

            # Item text
            text_top = top + Inches(0.85)
            text_height = box_height - Inches(1.0)
            text_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(left + Inches(0.1)), int(text_top),
                int(box_width - Inches(0.2)), int(text_height)
            )
            text_shape.fill.background()
            text_shape.line.fill.background()

            tf = text_shape.text_frame
            tf.word_wrap = True
            # Truncate text for display
            display_text = item[:80] + "..." if len(item) > 80 else item
            tf.text = display_text
            for para in tf.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                para.font.size = Pt(10)
                para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Add connecting arrows between boxes
        arrow_top = top + box_height / 2
        for i in range(n - 1):
            arrow_left = start_left + box_width * (i + 1) + gap * i
            arrow_shape = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                int(arrow_left - Inches(0.05)), int(arrow_top - Inches(0.1)),
                int(gap + Inches(0.1)), int(Inches(0.2))
            )
            arrow_shape.fill.solid()
            arrow_shape.fill.fore_color.rgb = colors.text_muted
            arrow_shape.line.fill.background()

    def _generate_timeline(self, slide, candidate: InfographicCandidate):
        """Generate a horizontal timeline infographic."""
        items = candidate.items
        if not items:
            return

        n = len(items)
        colors = self.design.colors

        # Timeline line
        line_top = Inches(3.8)
        line_left = Inches(1.0)
        line_width = Inches(11.3)

        line_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(line_left), int(line_top),
            int(line_width), int(Inches(0.06))
        )
        line_shape.fill.solid()
        line_shape.fill.fore_color.rgb = colors.primary
        line_shape.line.fill.background()

        # Timeline points and labels
        spacing = line_width / max(n - 1, 1) if n > 1 else line_width / 2
        box_colors = [colors.primary, colors.secondary, colors.accent1,
                     colors.accent2, colors.accent3]

        for i, item in enumerate(items):
            x = line_left + spacing * i if n > 1 else line_left + line_width / 2
            color = box_colors[i % len(box_colors)]

            # Dot on timeline
            dot_size = Inches(0.25)
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                int(x - dot_size / 2), int(line_top - dot_size / 2 + Inches(0.03)),
                int(dot_size), int(dot_size)
            )
            dot.fill.solid()
            dot.fill.fore_color.rgb = color
            dot.line.fill.background()

            # Alternating above/below
            if i % 2 == 0:
                text_top = line_top - Inches(1.5)
            else:
                text_top = line_top + Inches(0.4)

            text_width = Inches(2.0)
            text_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(x - text_width / 2), int(text_top),
                int(text_width), int(Inches(1.2))
            )
            text_shape.fill.solid()
            text_shape.fill.fore_color.rgb = color
            text_shape.line.fill.background()

            tf = text_shape.text_frame
            tf.word_wrap = True
            display_text = item[:60] + "..." if len(item) > 60 else item
            tf.text = display_text
            for para in tf.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                para.font.size = Pt(9)
                para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def _generate_comparison(self, slide, candidate: InfographicCandidate):
        """Generate a comparison layout with side-by-side blocks."""
        items = candidate.items
        if not items:
            return

        colors = self.design.colors
        n = len(items)

        # Layout: cards in a row
        start_left = Inches(0.7)
        top = Inches(2.0)
        total_width = Inches(11.9)
        card_height = Inches(3.5)
        gap = Inches(0.2)

        available = total_width - gap * (n - 1)
        card_width = available / n

        card_colors = [colors.primary, colors.secondary, colors.accent1,
                      colors.accent2, colors.accent3, colors.accent4]

        for i, item in enumerate(items):
            left = start_left + (card_width + gap) * i
            color = card_colors[i % len(card_colors)]

            # Card background
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(left), int(top), int(card_width), int(card_height)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = color
            card.line.fill.background()

            # Card text
            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.15)
            tf.margin_right = Inches(0.15)
            tf.margin_top = Inches(0.15)

            display_text = item[:100] + "..." if len(item) > 100 else item
            tf.text = display_text
            for para in tf.paragraphs:
                para.alignment = PP_ALIGN.LEFT
                para.font.size = Pt(10)
                para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def _generate_metrics(self, slide, candidate: InfographicCandidate):
        """Generate key metrics callout boxes."""
        items = candidate.items
        values = candidate.values
        if not items:
            return

        colors = self.design.colors
        n = min(len(items), 4)

        start_left = Inches(1.0)
        top = Inches(2.5)
        total_width = Inches(11.3)
        box_height = Inches(2.5)
        gap = Inches(0.3)

        available = total_width - gap * (n - 1)
        box_width = available / n

        metric_colors = [colors.primary, colors.accent1, colors.accent2, colors.accent3]

        for i in range(n):
            left = start_left + (box_width + gap) * i
            color = metric_colors[i % len(metric_colors)]

            # Metric box
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(left), int(top), int(box_width), int(box_height)
            )
            box.fill.solid()
            box.fill.fore_color.rgb = color
            box.line.fill.background()

            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.15)
            tf.margin_right = Inches(0.15)

            # Value (large)
            if i < len(values) and values[i]:
                p = tf.paragraphs[0]
                p.text = values[i]
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(28)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

            # Label
            if i < len(items):
                p = tf.add_paragraph()
                p.text = items[i][:50]
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(11)
                p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
