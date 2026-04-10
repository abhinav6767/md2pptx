from typing import List, Dict, Any
from .models import AgentStoryline
from .client import get_client
from ..parser import Document
import json

class VisualizerAgent:
    """Uses Gemini to assign the best visual strategy (chart, infographic, image) to each slide."""
    
    def __init__(self):
        self.client = get_client()

    def assign_visuals(self, storyline: AgentStoryline, tables_info: List[Dict[str, Any]]) -> AgentStoryline:
        
        # We need a schema for just the visual assignments to avoid rewriting the whole object if not needed
        # Or we can ask Gemini to return a list of updates.
        
        prompt = f"""
        You are an expert presentation designer. Your task is to review a presentation's slides and assign the most impactful visual strategy for each.
        
        Options for `recommended_visual`:
        - "text": AVOID THIS unless absolutely necessary (e.g., Agenda). Text-only slides are severely penalized in our guidelines.
        - "chart": The slide should visualize data (must be used if numerical tables exist).
        - "infographic": The slide text describes a process, timeline, or multi-part comparison.
        - "image": The slide is conceptual. Default to this instead of "text" to provide a visual anchor (requires you to write a great image generation prompt).
        
        You must also provide:
        - `visual_reference`: If 'chart', provide the EXACT title of the table to use. If 'image', provide a highly descriptive 1-line image generation prompt. If 'infographic', provide the type ("process", "timeline", "comparison", "metrics").
        - `is_two_column`: Boolean. Set to true if `recommended_visual` is "text" AND there are more than 4 bullet points, demanding a split layout.
        
        Available Tables (with numerical data):
        {json.dumps(tables_info, indent=2)}
        
        Slides to evaluate:
        {json.dumps([s.model_dump(exclude={'speaker_notes'}) for s in storyline.slides], indent=2)}
        
        Output:
        Return a JSON array where each object corresponds to a slide, in order, containing: `recommended_visual`, `visual_reference`, and `is_two_column`.
        """
        
        from pydantic import RootModel, BaseModel, Field
        class VisualOutput(BaseModel):
            recommended_visual: str = Field(description="text, chart, infographic, image")
            visual_reference: str = Field(None, description="Table title, infographic type, or image prompt")
            is_two_column: bool = Field(False)
            
        class VisualArray(RootModel):
            root: List[VisualOutput]
            
        response = self.client.models.generate_content(
            model='gemini-3.1-pro-preview',
            contents=prompt,
            config={
                'response_mime_type': 'application/json',
                'response_schema': VisualArray,
                'temperature': 0.1,
            },
        )
        
        updates = json.loads(response.text)
        
        # Apply updates
        for i, slide in enumerate(storyline.slides):
            if i < len(updates):
                upd = updates[i]
                slide.recommended_visual = upd.get('recommended_visual', 'text')
                slide.visual_reference = upd.get('visual_reference')
                slide.is_two_column = upd.get('is_two_column', False)
                
        return storyline
