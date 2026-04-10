from ..parser import Document
from ..template import TemplateLoader
from .storyline_agent import StorylineAgent
from .visualizer_agent import VisualizerAgent
from .layout_agent import LayoutAgent
from .image_agent import ImageAgent
from .models import AgentStoryline

class MultiAgentOrchestrator:
    """Orchestrates the entire multi-agent pipeline."""
    
    def __init__(self, target_slides: int = 12):
        self.storyline_agent = StorylineAgent(target_slides=target_slides)
        self.visualizer_agent = VisualizerAgent()
        self.layout_agent = LayoutAgent()
        self.image_agent = ImageAgent()

    def run(self, doc: Document, template_loader: TemplateLoader) -> AgentStoryline:
        print("\n    [Orchestrator] 1/4 - Storyline Agent (Rewriting & Emojis)...")
        storyline = self.storyline_agent.generate(doc)
        print(f"      -> Generated {len(storyline.slides)} slides.")

        print("\n    [Orchestrator] 2/4 - Visualizer Agent (Charts, Infographics, Images)...")
        # Extract table metadata to pass to Gemini
        tables_info = []
        for i, t in enumerate(doc.all_tables):
            if t.has_numerical_data:
                tables_info.append({"title": t.title or f"Table_{i}", "headers": t.headers, "rows": len(t.rows)})
        
        storyline = self.visualizer_agent.assign_visuals(storyline, tables_info)

        print("\n    [Orchestrator] 3/4 - Image Agent (Generating Visuals)...")
        for i, slide in enumerate(storyline.slides):
            if slide.recommended_visual == 'image' and slide.visual_reference:
                # Try to generate an image
                img_path = self.image_agent.generate_image(slide.visual_reference, i)
                if img_path:
                    slide.image_url = img_path
                else:
                    # Fallback to text if img gen failed
                    slide.recommended_visual = 'text'

        print("\n    [Orchestrator] 4/4 - Layout Assessment Agent (Grid Mapping)...")
        available_layouts = [l.name for l in template_loader.info.layouts]
        storyline = self.layout_agent.map_layouts(storyline, available_layouts)

        return storyline
