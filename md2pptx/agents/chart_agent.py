from .client import get_client
import json

class ChartAgent:
    """Uses Gemini to generate custom matplotlib Python code for dynamic, beautiful charts."""
    def __init__(self):
        self.client = get_client()

    def generate_chart_code(self, candidate, filepath: str, theme_colors: list = None) -> str:
        table = candidate.table
        chart_type = candidate.chart_type
        
        # Strip all emojis/unsupported glyphs to prevent matplotlib FT2Font C++ crash
        import re
        def clean_txt(t): return re.sub(r'[^\w\s.,!?-]', '', str(t))
        clean_headers = [clean_txt(h) for h in table.headers] if table.headers else []
        clean_rows = [[clean_txt(c) for c in row] for row in table.rows] if table.rows else []
        
        data_json = json.dumps({
            "title": clean_txt(candidate.section_title or table.title or "Chart Data"),
            "headers": clean_headers,
            "rows": clean_rows,
            "suggested_type": chart_type,
            "theme_colors": theme_colors or ['#1f77b4', '#ff7f0e', '#2ca02c']
        }, indent=2)

        prompt = f"""
You are an expert Python data visualization and infographic designer.
Your task is to write ONLY executable Python code (NO markdown blocks, NO explanations, NO ```python tags) that generates a highly visually stunning business chart utilizing `matplotlib`.

CRITICAL RULES:
1. Write raw standard Python code only. Do not enclose it in markdown blocks!
2. You MUST set the backend and import common libraries exactly at the top of your code: 
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

3. Render a beautiful, modern {chart_type} chart utilizing the data below. Make it look stunning, premium, and unique (e.g. use gradients, stylized markers, dropshadows, or dark modern aesthetics).
4. CRITICAL COLOR RULE: You MUST exclusively use the hex colors provided in `data['theme_colors']`. DO NOT INVENT OR REFER TO UNDEFINED VARIABLES. If you need a color, index the array `data['theme_colors'][i]`.
5. Save the figure exactly using the variable `filepath` (it is pre-injected into the local namespace, do NOT write its string literal).
6. Only use `fig.savefig(filepath, bbox_inches='tight', dpi=150)`. Do not use `plt.show()`.
7. Convert string data to floats properly using string replacement (remove $, %, commas).
8. Handle any data errors gracefully (skip invalid rows).
9. DO NOT include raw Windows file paths in your code (this causes 'bad escape \P' errors). Use the literal variable `filepath` instead of a string.
10. You MUST instantiate the data inline. DO NOT assume `data` is already defined. Start your code with exactly:

data = {data_json}

# Provide the raw Python code below:
"""
        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config={'temperature': 0.7} # slightly creative to make unique charts
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
