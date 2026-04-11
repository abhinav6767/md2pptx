import os
import json
from .client import get_client

class InfographicAgent:
    """Uses Gemini to generate custom matplotlib Python code for stunning infographics."""
    def __init__(self):
        self.client = get_client()

    def generate_infographic_code(self, candidate, filepath: str, theme_colors: list = None) -> str:
        # Strip all emojis/unsupported glyphs to prevent matplotlib FT2Font C++ crash
        import re
        def clean_txt(t): return re.sub(r'[^\w\s.,!?-]', '', str(t))
        clean_items = [clean_txt(i) for i in candidate.items] if candidate.items else []
        
        data_json = json.dumps({
            "type": candidate.infographic_type,
            "items": clean_items,
            "values": candidate.values,
            "theme_colors": theme_colors or ['#1f77b4', '#ff7f0e', '#2ca02c']
        }, indent=2)

        prompt = f"""
You are an expert Python data visualization and infographic designer.
Your task is to write ONLY executable Python code (NO markdown blocks, NO explanations, NO ```python tags) that generates a highly visually stunning infographic utilizing `matplotlib`.

CRITICAL RULES:
1. Write raw standard Python code only. Do not enclose it in markdown blocks!
2. You MUST set the backend and import common libraries exactly at the top of your code: 
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import re
import math

3. Render an amazing, deeply stylized {candidate.infographic_type} infographic. Use bubbles, flowcharts, or icons drawn with matplotlib shapes/text. Do not just draw a simple list. We want a very modern "wow" factor infographic.
4. CRITICAL COLOR RULE: You MUST exclusively use the hex colors provided in `data['theme_colors']`. DO NOT INVENT OR REFER TO UNDEFINED VARIABLES (like `light_text_color`). If you need a color, index the array `data['theme_colors'][i]`.
5. If the type is 'process', draw a linked node flow. If 'timeline', draw a beautiful timeline curve. If 'comparison', draw side-by-side gauge or metric layouts.
6. Save the figure exactly using the variable `filepath` (it is pre-injected into the local namespace).
7. Only use `fig.savefig(filepath, bbox_inches='tight', dpi=150, transparent=True)`. Do not use `plt.show()`.
8. Turn off all axes grids, ticks, and borders so it looks strictly like an infographic, not a graph (`ax.axis('off')`).
9. DO NOT include raw Windows file paths in your code (this causes 'bad escape \P' errors). Use the literal variable `filepath` instead of a string.
10. You MUST instantiate the data inline. DO NOT assume `data` is already defined. Start your code with exactly:

data = {data_json}

# Provide the raw Python code below:
"""
        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config={'temperature': 0.8} # creative
        )
        
        code = response.text.replace("```python", "").replace("```", "").strip()
        
        # --- GUARANTEED BOILERPLATE ---
        # Prepend critical imports and data binding. Then remove any DUPLICATE top-level
        # import statements that the AI may have written, but ONLY if they appear at
        # column 0 (unindented) to avoid breaking indented code blocks.
        boilerplate_lines = [
            "import matplotlib",
            "matplotlib.use('Agg')",
            "import matplotlib.pyplot as plt",
            "import matplotlib.patches as mpatches",
            "from matplotlib.patches import FancyBboxPatch, Circle, FancyArrowPatch, Arc, Wedge",
            "import numpy as np",
            "import re",
            "import math",
            "import textwrap",
            "from textwrap import wrap",
        ]
        # Remove duplicate top-level import lines from AI code
        ai_lines = code.splitlines()
        filtered = [l for l in ai_lines if l.strip() not in boilerplate_lines or l != l.lstrip()]
        code = "\n".join(filtered)
        
        boilerplate = "\n".join(boilerplate_lines) + f"\n\ndata = {data_json}\n"
        return boilerplate + "\n" + code
