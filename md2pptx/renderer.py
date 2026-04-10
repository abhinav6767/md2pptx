"""
PPTX Renderer — Builds the final .pptx file from an AgentStoryline and generated assets.
Places content into Slide Master layouts with proper formatting.
"""

import os
from typing import List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from .design import (
    DesignSystem, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
    SUBTITLE_LEFT, SUBTITLE_TOP, SUBTITLE_WIDTH, SUBTITLE_HEIGHT,
    BODY_LEFT, BODY_TOP, BODY_WIDTH, BODY_HEIGHT,
    COL_LEFT_LEFT, COL_LEFT_WIDTH, COL_RIGHT_LEFT, COL_RIGHT_WIDTH,
    COL_TOP, COL_HEIGHT, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT,
    FONT_TITLE, FONT_SUBTITLE, FONT_BODY, FONT_BODY_SMALL, FONT_CAPTION,
    FONT_COVER_TITLE, FONT_COVER_SUBTITLE, FONT_DIVIDER_TITLE,
    FONT_TABLE_HEADER, FONT_TABLE_BODY, FONT_SLIDE_NUMBER,
    MAX_BULLETS_PER_SLIDE, LINE_SPACING, PARAGRAPH_SPACING,
    FOOTER_LEFT, FOOTER_TOP, FOOTER_WIDTH, FOOTER_HEIGHT,
)
from .template import TemplateLoader, TemplateInfo
from .agents.models import AgentStoryline, AgentSlide
from .charts import ChartGenerator
from .infographics import InfographicGenerator
from .parser import TableData, Document
from .analyzer import ChartCandidate, InfographicCandidate

class PPTXRenderer:
    """Renders an AgentStoryline into a .pptx file using the Slide Master template."""

    def __init__(self, template_loader: TemplateLoader):
        self.template = template_loader
        self.info = template_loader.info
        self.design = self.info.design
        self.chart_gen = ChartGenerator(self.design)
        self.infographic_gen = InfographicGenerator(self.design)

    def render(self, storyline: AgentStoryline, doc: Document, output_path: str) -> str:
        """Render the storyline to a .pptx file."""
        prs = Presentation(self.info.path)

        # Remove template's existing example slides
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        # Create mapping of layout name to index
        layout_map = {l.name: l.index for l in self.info.layouts}

        chart_idx = 0

        # Render each slide
        for i, slide_content in enumerate(storyline.slides):
            safe_title = slide_content.title.encode('ascii', 'ignore').decode('ascii')
            print(f"  Rendering slide {i+1}/{len(storyline.slides)}: [{slide_content.layout_name}] {safe_title[:50]}")
            
            # Find requested layout
            layout_idx = layout_map.get(slide_content.layout_name, 0)
            layout = self._get_layout(prs, layout_idx)
            slide = prs.slides.add_slide(layout)

            # Fill Text Placeholders (Title, Subtitle)
            title_set = False
            for ph in slide.placeholders:
                idx = ph.placeholder_format.idx
                if idx in (0, 10):  # Main title
                    ph.text = slide_content.title
                    self._style_text_frame(ph.text_frame, FONT_TITLE, self.design.colors.text_dark, bold=True)
                    title_set = True
                elif idx == 11 and slide_content.subtitle:  # Subtitle
                    ph.text = slide_content.subtitle
                    self._style_text_frame(ph.text_frame, FONT_SUBTITLE, self.design.colors.text_dark)

            if not title_set:
                self._add_title_shape(slide, slide_content.title)

            # Route visual content
            visual = slide_content.recommended_visual
            
            # Two-column text layout
            if slide_content.is_two_column and slide_content.bullets:
                mid = len(slide_content.bullets) // 2
                self._add_column_bullets(slide, slide_content.bullets[:mid], COL_LEFT_LEFT, COL_TOP, COL_LEFT_WIDTH, COL_HEIGHT)
                self._add_column_bullets(slide, slide_content.bullets[mid:], COL_RIGHT_LEFT, COL_TOP, COL_RIGHT_WIDTH, COL_HEIGHT)
            
            # Single-column text layout
            elif visual == "text" and slide_content.bullets:
                self._add_bullets(slide, slide_content.bullets, center_align=False)

            # Image
            elif visual == "image" and getattr(slide_content, 'image_url', None):
                if os.path.exists(slide_content.image_url):
                    slide.shapes.add_picture(
                        slide_content.image_url,
                        int(CHART_LEFT), int(CHART_TOP),
                        int(CHART_WIDTH), int(CHART_HEIGHT)
                    )

            # Chart Processing
            elif visual == "chart" and slide_content.visual_reference:
                # Find matching table from the document
                matched_table = None
                for t in doc.all_tables:
                    if t.title == slide_content.visual_reference or t.title in slide_content.visual_reference:
                        matched_table = t
                        break
                
                # If no matching table, try first numerical table
                if not matched_table:
                    for t in doc.all_tables:
                        if t.has_numerical_data:
                            matched_table = t
                            break

                if matched_table:
                    # Construct chart candidate dynamically
                    # We default to BAR chart if not specified in prompt, but we can determine heuristic here
                    candidate = ChartCandidate(
                        table=matched_table,
                        chart_type="bar" if len(matched_table.rows) < 10 else "line",
                        section_title=slide_content.title
                    )
                    chart_path = self.chart_gen.generate(candidate, chart_idx)
                    if chart_path and os.path.exists(chart_path):
                        slide.shapes.add_picture(
                            chart_path,
                            int(CHART_LEFT), int(CHART_TOP),
                            int(CHART_WIDTH), int(CHART_HEIGHT)
                        )
                    chart_idx += 1
                else:
                    # Fallback to bullets
                    self._add_bullets(slide, slide_content.bullets, center_align=False)

            # Infographic Processing
            elif visual == "infographic":
                info_type = slide_content.visual_reference or "process"
                info_type = info_type.lower()
                
                mapped_type = "process"
                if "timeline" in info_type:
                    mapped_type = "timeline"
                elif "comparison" in info_type:
                    mapped_type = "comparison"
                elif "metrics" in info_type:
                    mapped_type = "metrics"
                    
                candidate = InfographicCandidate(
                    infographic_type=mapped_type,
                    title=slide_content.title,
                    items=slide_content.bullets
                )
                self.infographic_gen.generate(slide, candidate)


            # Adding Slide Number
            self._add_slide_number(slide, prs)

        # Save
        os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else '.', exist_ok=True)
        prs.save(output_path)
        return output_path

    def _get_layout(self, prs: Presentation, layout_idx: int):
        sm = prs.slide_masters[0]
        if 0 <= layout_idx < len(sm.slide_layouts):
            return sm.slide_layouts[layout_idx]
        return sm.slide_layouts[0]

    def _add_title_shape(self, slide, title: str):
        txbox = slide.shapes.add_textbox(int(TITLE_LEFT), int(TITLE_TOP), int(TITLE_WIDTH), int(TITLE_HEIGHT))
        tf = txbox.text_frame
        tf.text = title
        tf.word_wrap = True
        self._style_text_frame(tf, FONT_TITLE, self.design.colors.text_dark, bold=True)

    def _add_bullets(self, slide, bullets: List[str], center_align: bool = False):
        txbox = slide.shapes.add_textbox(int(BODY_LEFT), int(BODY_TOP), int(BODY_WIDTH), int(BODY_HEIGHT))
        tf = txbox.text_frame
        tf.word_wrap = True
        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"  •  {bullet}"
            p.font.size = FONT_BODY
            p.font.color.rgb = self.design.colors.text_dark
            p.space_after = Pt(8)
            p.space_before = Pt(4)
            if center_align:
                # Per new guidelines, body text must not be centered. Forcing left alignment instead.
                p.alignment = PP_ALIGN.LEFT

    def _add_column_bullets(self, slide, bullets: List[str], left, top, width, height):
        txbox = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
        tf = txbox.text_frame
        tf.word_wrap = True
        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"  •  {bullet}"
            p.font.size = FONT_BODY_SMALL
            p.font.color.rgb = self.design.colors.text_dark
            p.space_after = Pt(6)
            # Per new guidelines, body text must not be centered. We leave it default left-aligned.
            p.alignment = PP_ALIGN.LEFT

    def _add_slide_number(self, slide, prs: Presentation):
        slide_num = len(prs.slides)
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 4:
                ph.text = str(slide_num)
                return
        txbox = slide.shapes.add_textbox(int(FOOTER_LEFT), int(FOOTER_TOP), int(FOOTER_WIDTH), int(FOOTER_HEIGHT))
        tf = txbox.text_frame
        tf.text = str(slide_num)
        tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
        tf.paragraphs[0].font.size = FONT_SLIDE_NUMBER
        tf.paragraphs[0].font.color.rgb = self.design.colors.text_muted

    def _style_text_frame(self, text_frame, font_size: Pt, color: RGBColor, bold: bool = False, alignment=None):
        for para in text_frame.paragraphs:
            para.font.size = font_size
            para.font.color.rgb = color
            para.font.bold = bold
            if alignment:
                para.alignment = alignment
