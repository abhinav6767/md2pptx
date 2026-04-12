from typing import List, Dict, Any
from .models import AgentStoryline
from .client import get_client
import json

class LayoutAgent:
    """Uses Gemini to assign the correct PPTX layout name to each slide, ensuring constraints."""
    
    def __init__(self):
        self.client = get_client()

    def map_layouts(self, storyline: AgentStoryline, available_layouts: List[str]) -> AgentStoryline:
        
        prompt = f"""
        You are an expert presentation layout engine. Your job is to select the exact layout name from the available Slide Master layouts for each slide.
        
        Available Layouts:
        {json.dumps(available_layouts)}
        
        Rules for mapping:
        1. The first slide (Cover) should use a layout with "Cover" or "Title" in the name.
        2. Content slides with text/bullets should use a layout with "Title" or "Content" or "Title and Vertical Text".
        3. AVOID "Blank" or "Thank You" for any slide that contains bullets, charts, or infographics. "Blank" should only be used as a last resort if no other layout fits.
        4. Section dividers should use a layout with "Divider" or "Section" in the name if available.
        5. "Thank You" should ONLY be used for the final slide. Do NOT use it for content slides.
        
        Slides to evaluate:
        {json.dumps([s.model_dump(exclude={'speaker_notes', 'bullets'}) for s in storyline.slides], indent=2)}
        
        Output:
        Return a JSON array where each element contains `layout_name` (must be exactly one of the Available Layouts).
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
        
        for i, slide in enumerate(storyline.slides):
            if i < len(updates):
                upd = updates[i]
                slide.layout_name = upd.get('layout_name', available_layouts[0])
                
        return storyline
