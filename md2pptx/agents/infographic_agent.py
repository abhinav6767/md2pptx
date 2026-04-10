import os
import json
from .client import get_client

class InfographicAgent:
    """Uses Gemini to generate custom matplotlib Python code for stunning infographics."""
    def __init__(self):
        self.client = get_client()

    def generate_infographic_code(self, candidate, filepath: str) -> str:
        data_json = json.dumps({
            "type": candidate.infographic_type,
            "items": candidate.items,
            "values": candidate.values
        }, indent=2)

        prompt = f"""
You are an expert Python data visualization and infographic designer.
Your task is to write ONLY executable Python code (NO markdown blocks, NO explanations, NO ```python tags) that generates a highly visually stunning infographic utilizing `matplotlib`.

CRITICAL RULES:
1. Write raw standard Python code only. Do not enclose it in markdown blocks!
2. You MUST set the backend: 
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

3. Render an amazing, deeply stylized {candidate.infographic_type} infographic. Use beautiful corporate color schemes, bubbles, flowcharts, or icons drawn with matplotlib shapes/text. Do not just draw a simple list. We want a very modern "wow" factor infographic.
4. If the type is 'process', draw a linked node flow. If 'timeline', draw a beautiful timeline curve. If 'comparison', draw side-by-side gauge or metric layouts.
5. Save the figure exactly to the variable `filepath` (it is pre-injected into the local namespace).
6. Only use `fig.savefig(filepath, bbox_inches='tight', dpi=150, transparent=True)`. Do not use `plt.show()`.
7. Turn off all axes grids, ticks, and borders so it looks strictly like an infographic, not a graph (`ax.axis('off')`).

Data to visualize:
{data_json}

# Provide the raw Python code below:
"""
        response = self.client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config={'temperature': 0.8} # creative
        )
        
        code = response.text.replace("```python", "").replace("```", "").strip()
        return code
