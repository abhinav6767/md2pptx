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
        You are an elite management consultant and presentation designer. Your task is to extract the key narrative from the provided text and structure it into exactly {self.target_slides} highly professional presentation slides.
        
        Guidelines:
        1. Narrative Flow: 
           - Start with a Cover slide, then Agenda, then Executive Summary.
           - Build the core content logically, ending with a Conclusion/Takeaways slide, and finally a Thank You slide.
        2. Infographic First Approach (CRITICAL):
           - Before placing text, ask yourself: “Can this be visualized?”.
           - If YES → Set 'recommended_visual' to 'infographic' (e.g., process flow, timeline, metrics, comparison). Convert the text into short labels or steps rather than paragraphs.
           - If NO → Keep text minimal (max 6 lines total per slide). Avoid bullet overload. Focus on 1 clear takeaway per slide.
        3. Visuals over Text:
           - Completely avoid paragraphs. Synthesize into punchy, consultant-grade lists or diagrams.
           - We strongly prefer mixed-media slides (e.g. Text + Image side-by-side) over text-only slides. Use 'image' or 'chart' or 'infographic' as 'recommended_visual' whenever data or concepts support it.
           - Premium Layout 1 (Hero Header): If the slide is an Executive Summary, Agenda, or High-Level overview, set 'recommended_visual' to 'hero_header'. This crafts a stunning full-width top image with 4 sleek metric cards at the bottom.
           - Premium Layout 2 (Sidebar Split): If the slide is a deep-dive, process, or comparison, set 'recommended_visual' to 'sidebar_split'. This creates a solid colored 33% left sidebar mapping into a large canvas infographics area on the right.
           - Avoid 'ultra_dense' unless there is too much data to fit in a standard layout. Choose 'hero_header' or 'sidebar_split' first for maximum premium impact.
        4. Hierarchy & Tone: 
           - Provide ONE clear key message per slide.
           - Establish clear hierarchy: Title (Primary message) > Subtitle (optional context) > Body Content (supporting points).
           - Use tasteful emojis as modern bullet icons where appropriate.
           
        Source Content:
        {content[:40000]} # Trim to fit typical Gemini prompts if needed
        """
        
        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
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
