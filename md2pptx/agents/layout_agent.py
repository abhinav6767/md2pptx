from typing import List, Dict, Any
from .models import AgentStoryline
from .client import get_client
import json

class LayoutAgent:
    """Uses Gemini to assign the correct PPTX layout name to each slide, ensuring constraints."""

    def __init__(self):
        self.client = get_client()

    def map_layouts(self, storyline: AgentStoryline, available_layouts: List[str]) -> AgentStoryline:

        # Build compact slide descriptors that include visual hints
        slide_descs = []
        for s in storyline.slides:
            slide_descs.append({
                "title": s.title,
                "recommended_visual": s.recommended_visual,
                "is_two_column": s.is_two_column,
                "has_bullets": bool(s.bullets),
            })

        n_slides = len(slide_descs)
        prompt = f"""
        You are an expert presentation layout engine. Your job is to select the exact layout name
        from the available Slide Master layouts for each slide.

        Available Layouts:
        {json.dumps(available_layouts)}

        Rules (apply in strict priority order):
        1. Slide index 0 (the FIRST slide / Cover) MUST use a layout whose name contains
           "Cover", "Title Company", or "0_Title". NO OTHER slide may use this layout.
        2. Slide index {n_slides - 1} (the VERY LAST slide / Thank You) MUST use a layout
           whose name contains "Thank You" or "thank_you" or "1_Thank", if one exists.
           NO OTHER slide may use this layout.
        3. Slides with recommended_visual IN (text, chart, infographic, image, numbered_list,
           swimlane, icon_cards, data_table, vertical_timeline) AND has_bullets = true:
           MUST use a layout whose name contains "Content", "Title, Content",
           "Title and Content", "1_E_Title", or "Title Only".
        4. Two-column slides (is_two_column = true): prefer layouts with "Two Content" or "Comparison".
        5. Slides with recommended_visual IN (hero_header, sidebar_split): use "Blank" or a minimal layout.
        6. "Divider" / "Section" layouts are ONLY for section transition slides between major topics,
           never for content slides that have bullet points or charts.
        7. "Blank" ONLY as a last resort when no other layout fits a content slide.
        8. NEVER assign a Cover/Thank-You/Divider layout to a middle content slide.

        Slides to evaluate (index, title, visual type, has_bullets):
        {json.dumps(slide_descs, indent=2)}

        Output:
        A JSON array (one element per slide in order). Each element:
          {{ "layout_name": "<exact string from Available Layouts>" }}
        """

        from pydantic import RootModel, BaseModel, Field

        class LayoutOutput(BaseModel):
            layout_name: str = Field(description="Exact name from available layouts")

        class LayoutArray(RootModel):
            root: List[LayoutOutput]

        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config={
                'response_mime_type': 'application/json',
                'response_schema': LayoutArray,
                'temperature': 0.0,
            },
        )

        updates = json.loads(response.text)

        # --- Identify safe fallback layouts for guard-rails ---
        def _best_layout(candidates, available, default_idx=0):
            """Return the first layout name that matches any candidate pattern."""
            for name in available:
                for c in candidates:
                    if c in name.lower():
                        return name
            return available[default_idx]

        cover_patterns    = ["cover", "0_title", "title company"]
        content_patterns  = ["content", "title, content", "title and content", "1_e_title", "title only"]
        thankyou_patterns = ["thank you", "thank_you", "1_thank"]
        divider_patterns  = ["divider", "section", "c_section"]
        blank_patterns    = ["blank"]

        cover_layout    = _best_layout(cover_patterns, available_layouts)
        content_layout  = _best_layout(content_patterns, available_layouts)
        thankyou_layout = _best_layout(thankyou_patterns, available_layouts)
        fallback_layout = _best_layout(blank_patterns, available_layouts, default_idx=0)

        n_slides = len(storyline.slides)
        for i, slide in enumerate(storyline.slides):
            if i < len(updates):
                chosen = updates[i].get('layout_name', available_layouts[0])
                # Safety: ensure the chosen name is actually among available layouts
                if chosen not in available_layouts:
                    chosen = content_layout

                # --- Hard guard-rails (deterministic, override any LLM mistake) ---
                chosen_lower = chosen.lower()

                # Rule A: Only slide 0 can use the Cover layout
                is_cover = any(p in chosen_lower for p in cover_patterns)
                if is_cover and i != 0:
                    chosen = content_layout

                # Rule B: Only the last slide can use Thank-You layout
                is_thankyou = any(p in chosen_lower for p in thankyou_patterns)
                if is_thankyou and i != n_slides - 1:
                    chosen = content_layout

                # Rule C: Divider layouts only for non-first, non-last slides
                #         AND only when there are no bullets (pure section headers)
                is_divider = any(p in chosen_lower for p in divider_patterns)
                if is_divider and (i == 0 or i == n_slides - 1 or slide.bullets):
                    chosen = content_layout if slide.bullets else fallback_layout

                slide.layout_name = chosen

        return storyline
