from .client import get_client
import json

class ChartAgent:
    """Uses Gemini to generate custom matplotlib Python code for dynamic, beautiful charts."""
    def __init__(self):
        self.client = get_client()

    def generate_chart_code(self, candidate, filepath: str) -> str:
        table = candidate.table
        chart_type = candidate.chart_type
        
        data_json = json.dumps({
            "title": candidate.section_title or table.title or "Chart Data",
            "headers": table.headers,
            "rows": table.rows,
            "suggested_type": chart_type
        }, indent=2)

        prompt = f"""
You are an expert Python data visualization designer.
Your task is to write ONLY executable Python code (NO markdown blocks, NO explanations, NO ```python tags) that generates a highly attractive, unique, and professional corporate chart.

CRITICAL RULES:
1. Write raw standard Python code only. Do not enclose it in markdown blocks!
2. You MUST set the backend: 
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

3. Render a beautiful, modern {chart_type} chart utilizing the data below. Make it look stunning, premium, and unique (e.g. use beautiful corporate colors, gradients, stylized markers, dropshadows, or dark modern aesthetics).
4. Save the figure exactly to the variable `filepath` (it is pre-injected into the local namespace, do NOT overwrite it).
5. Only use `fig.savefig(filepath, bbox_inches='tight', dpi=150)`. Do not use `plt.show()`.
6. Convert string data to floats properly using string replacement (remove $, %, commas).
7. Handle any data errors gracefully (skip invalid rows).

Dataset:
{data_json}

# Provide the raw Python code below:
"""
        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config={'temperature': 0.7} # slightly creative to make unique charts
        )
        
        code = response.text.replace("```python", "").replace("```", "").strip()
        return code
