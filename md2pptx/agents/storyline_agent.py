from typing import List
from ..parser import Document
from .models import AgentStoryline
from .client import get_client

class StorylineAgent:
    """Uses Gemini to rewrite markdown content into a structured storyline."""
    
    def __init__(self, target_slides: int = 12):
        self.client = get_client()
        self.target_slides = target_slides

    def generate(self, doc: Document) -> AgentStoryline:
        # Extract text to send to Gemini
        # For large documents, we chunk or summarize. Here we send a compressed version.
        content = f"# {doc.title}\n{doc.subtitle}\n\n"
        for i, sec in enumerate(doc.sections):
            text = sec.get_text_content()
            # Compress to avoid exceeding context or generating too much
            content += f"## {sec.title}\n{text[:1000]}\n"
        
        prompt = f"""
        You are an expert presentation designer. Your task is to extract the key narrative from the provided text and structure it into exactly {self.target_slides} slides.
        
        Guidelines:
        1. Narrative Flow: Ensure the first slide is the Cover, the second is an Agenda, followed by Executive Summary, core content, and a Conclusion slide, ending with a Thank You slide.
        2. Content Structure: Rewrite the text to be punchy and structured. AVOID excessive bullet points. Keep text minimal (max 6-8 lines total per slide) and prioritize a clear visual hierarchy (one key message per slide).
        3. Infographic & Visual Focus: Structure the text so it can seamlessly translate into visual layouts (e.g., process steps, comparisons, key metrics) rather than paragraphs.
        4. Emojis: Use tasteful emojis to make the content visually engaging (1-2 per slide maximum). Target conceptual and data-driven insights.
        
        Source Content:
        {content[:30000]} # Trim to fit typical Gemini prompts if needed
        """
        
        response = self.client.models.generate_content(
            model='gemini-3.1-pro-preview',
            contents=prompt,
            config={
                'response_mime_type': 'application/json',
                'response_schema': AgentStoryline,
                'temperature': 0.2,
            },
        )
        
        if not response.text:
             raise ValueError("Storyline Agent failed to return a response.")
             
        # Parse output dynamically
        import json
        data = json.loads(response.text)
        return AgentStoryline(**data)
