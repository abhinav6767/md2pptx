from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any

class AgentSlide(BaseModel):
    title: str = Field(description="The main title of the slide.")
    subtitle: Optional[str] = Field(None, description="An optional subtitle.")
    bullets: List[str] = Field(description="List of concise bullet points for the slide (max 6 bullets). Use appropriate emojis.")
    speaker_notes: Optional[str] = Field(None, description="Speaker notes providing more context.")
    # Fields to be filled by subsequent agents
    recommended_visual: str = Field("text", description="One of: text, chart, infographic, image")
    visual_reference: Optional[str] = Field(None, description="Reference to a table title, infographic type, or image prompt")
    layout_name: Optional[str] = Field(None, description="The specific Slide Master layout to use")
    is_two_column: bool = Field(False, description="Whether the text should be split into two columns")

class AgentStoryline(BaseModel):
    slides: List[AgentSlide] = Field(description="The list of 10 to 15 slides forming the presentation.")
