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

        layout_map = {l.name: l.index for l in self.info.layouts}

        chart_idx = 0
        image_idx = 0
        mask_shapes = [
            MSO_SHAPE.ROUNDED_RECTANGLE, 
            MSO_SHAPE.OVAL, 
            MSO_SHAPE.SNIP_1_RECTANGLE, 
            MSO_SHAPE.ROUND_2_DIAG_RECTANGLE,
            MSO_SHAPE.ROUND_1_RECTANGLE
        ]

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

            # 0. Forced Fallback for Title/Subtitle (if placeholders are missing)
            # This ensures no slide is ever left empty or unbranded.
            if not title_set:
                self._add_title_shape(slide, slide_content.title)
                title_set = True
            
            # Add short para (subtitle) fallback
            if slide_content.subtitle:
                # Check if subtitle was already set via placeholders
                subtitle_set_already = False
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx == 11 and ph.text:
                        subtitle_set_already = True
                        break
                
                if not subtitle_set_already:
                    self._add_subtitle_paragraph(slide, slide_content.subtitle)

            visual = slide_content.recommended_visual
            has_bullets = bool(slide_content.bullets)
            
            # --- Detect if this is a protected layout (Cover, Thank You, Divider)
            # These layouts have their own rich template backgrounds — never overlay images on them.
            layout_name_lower = (slide_content.layout_name or "").lower()
            is_protected_layout = any(p in layout_name_lower for p in [
                "cover", "thank you", "thank_you", "divider", "section"
            ])
            
            has_image = (
                not is_protected_layout
                and visual in ("image", "ultra_dense", "hero_header", "sidebar_split")
                and getattr(slide_content, 'image_url', None)
                and os.path.exists(slide_content.image_url)
            )

            # --- Z-Order Override for Premium Layouts ---
            # Native placeholders sit at the bottom. For custom solid backgrounds, we must
            # clear the placeholders and redraw text on top to prevent obstruction.
            if visual in ("hero_header", "sidebar_split"):
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        shape.text_frame.clear()  # Hide native text to allow manual layering on top
                # Premium layouts manage their own titles.
                title_set = True 
            
            # 1. Two-column text layout
            if slide_content.is_two_column and has_bullets:
                mid = len(slide_content.bullets) // 2
                self._add_column_bullets(slide, slide_content.bullets[:mid], COL_LEFT_LEFT, COL_TOP, COL_LEFT_WIDTH, COL_HEIGHT)
                self._add_column_bullets(slide, slide_content.bullets[mid:], COL_RIGHT_LEFT, COL_TOP, COL_RIGHT_WIDTH, COL_HEIGHT)
            
            # 2. Single-column text layout
            elif visual == "text" and has_bullets:
                self._add_bullets(slide, slide_content.bullets, center_align=False)

            # 3. Mixed Media: Text + Image (50/50 split with exact padding)
            elif visual == "image" and has_bullets and has_image:
                # Text on left
                self._add_column_bullets(slide, slide_content.bullets, COL_LEFT_LEFT, COL_TOP, COL_LEFT_WIDTH, COL_HEIGHT)
                
                # Enforce precise margin padding between columns
                pic = slide.shapes.add_picture(
                    slide_content.image_url,
                    int(COL_RIGHT_LEFT), int(COL_TOP),
                    int(COL_RIGHT_WIDTH), int(COL_HEIGHT)
                )
                pic.auto_shape_type = mask_shapes[image_idx % len(mask_shapes)]
                image_idx += 1

            # 4. Image Only (Full-width image placed BELOW title, not covering it)
            elif visual == "image" and has_image:
                # Place image safely in the body zone (below title area)
                pic = slide.shapes.add_picture(
                    slide_content.image_url,
                    int(BODY_LEFT), int(BODY_TOP),
                    int(BODY_WIDTH), int(BODY_HEIGHT)
                )
                pic.auto_shape_type = mask_shapes[image_idx % len(mask_shapes)]
                image_idx += 1

            # 5. Chart Processing (with optional text left if requested)
            elif visual == "chart" and slide_content.visual_reference:
                matched_table = None
                for t in doc.all_tables:
                    if t.title == slide_content.visual_reference or t.title in slide_content.visual_reference:
                        matched_table = t
                        break
                
                if not matched_table:
                    for t in doc.all_tables:
                        if t.has_numerical_data:
                            matched_table = t
                            break

                if matched_table:
                    candidate = ChartCandidate(
                        table=matched_table,
                        chart_type="bar" if len(matched_table.rows) < 10 else "line",
                        section_title=slide_content.title
                    )
                    chart_path = self.chart_gen.generate(candidate, chart_idx)
                    if chart_path and os.path.exists(chart_path):
                        if has_bullets:
                            # Mixed media Text + Chart
                            self._add_column_bullets(slide, slide_content.bullets, COL_LEFT_LEFT, COL_TOP, COL_LEFT_WIDTH, COL_HEIGHT)
                            slide.shapes.add_picture(
                                chart_path,
                                int(COL_RIGHT_LEFT), int(COL_TOP),
                                int(COL_RIGHT_WIDTH), int(COL_HEIGHT)
                            )
                        else:
                            # Full width chart
                            slide.shapes.add_picture(
                                chart_path,
                                int(CHART_LEFT), int(CHART_TOP),
                                int(CHART_WIDTH), int(CHART_HEIGHT)
                            )
                    chart_idx += 1
                else:
                    self._add_bullets(slide, slide_content.bullets, center_align=False)

            # 5.5 Native Consultant Layouts — route bullets directly to premium generators
            elif visual in ("numbered_list", "swimlane", "icon_cards", "data_table", "vertical_timeline"):
                import copy
                # Map visual type to infographic type name
                info_type_map = {
                    "numbered_list": "numbered_list",
                    "swimlane": "swimlane",
                    "icon_cards": "icon_cards",
                    "data_table": "data_table",
                    "vertical_timeline": "vertical_timeline",
                }
                info_candidate = InfographicCandidate(
                    infographic_type=info_type_map[visual],
                    title=slide_content.title,
                    items=copy.deepcopy(slide_content.bullets),
                    values=[]
                )
                body_bounds = (BODY_LEFT, BODY_TOP, BODY_WIDTH, BODY_HEIGHT)
                self.infographic_gen.generate(slide, info_candidate, bounds=body_bounds)

            # 5.5. Ultra-Dense (Title + Text Left + Image Top-Right + Infographic Bottom-Right)
            elif visual == "ultra_dense":
                # 1. Text on Left
                self._add_column_bullets(slide, slide_content.bullets, COL_LEFT_LEFT, COL_TOP, COL_LEFT_WIDTH, COL_HEIGHT)
                
                # 2. Image on Top Right (if exists)
                img_exists = has_image
                if img_exists:
                    img_height = COL_HEIGHT / 2 - Inches(0.15)
                    pic = slide.shapes.add_picture(
                        slide_content.image_url,
                        int(COL_RIGHT_LEFT), int(COL_TOP),
                        int(COL_RIGHT_WIDTH), int(img_height)
                    )
                    pic.auto_shape_type = mask_shapes[image_idx % len(mask_shapes)]
                    image_idx += 1
                
                # 3. Miniature Infographic on Bottom Right
                info_top = COL_TOP + (COL_HEIGHT / 2) + Inches(0.15) if img_exists else COL_TOP
                info_height = COL_HEIGHT / 2 - Inches(0.15) if img_exists else COL_HEIGHT
                bounds = (COL_RIGHT_LEFT, info_top, COL_RIGHT_WIDTH, info_height)
                
                info_type = slide_content.visual_reference or "process"
                info_type = info_type.lower()
                mapped_type = "process"
                if "timeline" in info_type: mapped_type = "timeline"
                elif "comparison" in info_type: mapped_type = "comparison"
                elif "metrics" in info_type: mapped_type = "metrics"
                
                import copy
                info_candidate = InfographicCandidate(
                    infographic_type=mapped_type,
                    title=slide_content.title,
                    items=copy.deepcopy(slide_content.bullets)
                )
                self.infographic_gen.generate(slide, info_candidate, bounds=bounds)

            # 5.6 Premium: Hero Header
            elif visual == "hero_header":
                # Top 45% Image
                hero_height = Inches(3.3)
                if has_image:
                    pic = slide.shapes.add_picture(
                        slide_content.image_url,
                        0, 0, Inches(13.33), int(hero_height)
                    )
                
                # Intersecting semi-circle (Oval slightly off bottom)
                oval_width = Inches(8.0)
                oval_height = Inches(4.5)
                oval_left = (Inches(13.33) - oval_width) / 2
                oval_top = hero_height - (oval_height / 2)
                
                oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(oval_left), int(oval_top), int(oval_width), int(oval_height))
                oval.fill.solid()
                oval.fill.fore_color.rgb = self.design.colors.primary
                oval.line.fill.background()
                
                # Title manually layered on top
                tx_box = slide.shapes.add_textbox(int(oval_left + Inches(0.5)), int(oval_top + Inches(1.0)), int(oval_width - Inches(1.0)), Inches(1.5))
                tx_box.text_frame.word_wrap = True
                tx_box.text_frame.text = slide_content.title.upper()
                self._style_text_frame(tx_box.text_frame, Pt(36), self.design.colors.text_light, bold=True, alignment=PP_ALIGN.CENTER)
                
                # Generate bottom metrics
                info_candidate = InfographicCandidate(infographic_type="premium_cards", title="", items=slide_content.bullets)
                bounds = (Inches(0.5), hero_height + Inches(1.5), Inches(12.33), Inches(2.5))
                self.infographic_gen.generate(slide, info_candidate, bounds=bounds)

            # 5.7 Premium: Sidebar Split
            elif visual == "sidebar_split":
                # Left Sidebar 33%
                sb_width = Inches(4.44)
                sb_top_height = Inches(3.75)
                
                # Solid primary color top panel
                rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, int(sb_width), int(sb_top_height))
                rect.fill.solid()
                rect.fill.fore_color.rgb = self.design.colors.primary
                rect.line.fill.background()
                
                # Title overlaid securely in white text (always visible on primary background)
                tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.5), int(sb_width - Inches(0.6)), Inches(2.5))
                tb.text_frame.word_wrap = True
                p_title = tb.text_frame.paragraphs[0]
                p_title.text = slide_content.title
                p_title.font.size = Pt(26)
                p_title.font.bold = True
                p_title.font.color.rgb = self.design.colors.text_light  # ALWAYS white on primary bg
                
                if slide_content.subtitle:
                    p_sub = tb.text_frame.add_paragraph()
                    p_sub.text = slide_content.subtitle
                    p_sub.font.size = Pt(12)
                    p_sub.font.bold = False
                    p_sub.font.color.rgb = self.design.colors.text_light  # ALWAYS white on primary bg


                # Image Bottom (Curved)
                if has_image:
                    img_top = sb_top_height
                    img_height = Inches(3.75)
                    pic = slide.shapes.add_picture(
                        slide_content.image_url,
                        Inches(0.2), int(img_top + Inches(0.2)), int(sb_width - Inches(0.4)), int(img_height - Inches(0.4))
                    )
                    pic.auto_shape_type = MSO_SHAPE.OVAL 

                # Process graphics mapping to the right 66%
                info_type = slide_content.visual_reference or "process"
                if "comparison" in info_type.lower(): mapped_type = "comparison"
                else: mapped_type = "gear_process"
                
                info_candidate = InfographicCandidate(infographic_type=mapped_type, title="", items=slide_content.bullets)
                bounds = (sb_width + Inches(0.3), Inches(1.0), Inches(13.33) - sb_width - Inches(0.6), Inches(5.5))
                self.infographic_gen.generate(slide, info_candidate, bounds=bounds)

            # 6. Infographic Processing (Uses bullets to generate shapes)
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

    def _add_subtitle_paragraph(self, slide, subtitle: str):
        """Adds a short summary paragraph (executive summary) below the title."""
        txbox = slide.shapes.add_textbox(int(SUBTITLE_LEFT), int(SUBTITLE_TOP), int(SUBTITLE_WIDTH), int(SUBTITLE_HEIGHT))
        tf = txbox.text_frame
        tf.text = subtitle
        tf.word_wrap = True
        # Using a slightly larger font for the top para to make it stand out
        self._style_text_frame(tf, Pt(15), self.design.colors.text_muted, bold=False, italic=True)

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

    def _style_text_frame(self, text_frame, font_size: Pt, color: RGBColor, bold: bool = False, italic: bool = False, alignment=None):
        for para in text_frame.paragraphs:
            para.font.size = font_size
            para.font.color.rgb = color
            para.font.bold = bold
            para.font.italic = italic
            if alignment:
                para.alignment = alignment
