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

        2. Visual-First Design (CRITICAL — follow strictly):
           NEVER use `text` or bullet-only slides. Every slide must have a visual. Choose the best type:

           - `numbered_list`: Use for 3-6 mechanisms, factors, or key points where order matters (e.g. "5 Key Risks", "4 Drivers"). Bullets must follow format "Header Title: Description detail".
           - `swimlane`: Use for parallel or multi-track processes/pillars where each item has equal weight (e.g. "3 Strategic Pillars", "5-step Approach"). Bullets must follow "Header: Description" format.
           - `icon_cards`: Use for categorised safeguards, recommendations, or frameworks with 4-6 distinct topics (e.g. "Institutional, Retail, Operational safeguards"). Bullets must follow "Category Name: What it means" format.
           - `data_table`: Use when content contains quantitative indicators, metrics, or statistics with multiple dimensions. Bullets must follow "Indicator | Value | Year | Source" pipe-separated format.
           - `vertical_timeline`: Use for historical progressions, phase rollouts, or step-by-step sequences in chronological order.
           - `chart`: Use when there is explicit tabular data to visualize (bar, line, pie). Set visual_reference to the name of the relevant table.
           - `hero_header`: Use for Executive Summary, Agenda/Overview, or high-level overview slides. Generates a stunning full-width header image with metric cards.
           - `sidebar_split`: Use for deep-dive or analytical slides with a single strong theme on the left and process/flow on the right.
           - `image`: Use when a realistic photographic illustration will enhance the content.

        3. Hierarchy & Tone:
           - Every slide MUST have a Title (primary message) and MUST have a Subtitle (one sentence of context).
           - Limit to 1 key message per slide. Keep bullets punchy and ≤ 12 words each.
           - USE emojis at the START of bullets where they genuinely aid clarity. Choose topic-relevant
             emojis (e.g. 📊 for data, ⚠️ for risk, ✅ for benefits, 🚀 for growth, 💡 for insight,
             🏦 for finance, 🌍 for global, ⚙️ for process, 🔬 for research). Limit to 1 emoji per bullet.
            
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
