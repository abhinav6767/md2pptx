"""
Chart Generator — Creates high-quality matplotlib charts from tabular data.
Charts are styled to match the Slide Master's visual language.
"""

import os
import io
import re
import tempfile
from typing import List, Optional, Tuple

import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from mpl_toolkits.mplot3d import Axes3D

from .analyzer import ChartCandidate
from .design import DesignSystem


class ChartGenerator:
    """Generates matplotlib charts styled to match the presentation theme."""

    def __init__(self, design: DesignSystem, output_dir: str = None):
        self.design = design
        self.output_dir = output_dir or tempfile.mkdtemp(prefix="md2pptx_charts_")
        os.makedirs(self.output_dir, exist_ok=True)
        self._setup_style()

    def _setup_style(self):
        """Configure matplotlib global style."""
        plt.rcParams.update({
            'figure.facecolor': 'white',
            'axes.facecolor': 'white',
            'font.family': 'sans-serif',
            'font.sans-serif': [self.design.fonts.body, 'Calibri', 'Arial', 'sans-serif'],
            'font.size': 10,
            'axes.titlesize': 14,
            'axes.labelsize': 10,
            'xtick.labelsize': 9,
            'ytick.labelsize': 9,
            'legend.fontsize': 9,
            'axes.linewidth': 0, # remove 2D borders for 3D illusion
        })

    def generate(self, candidate: ChartCandidate, chart_index: int = 0) -> Optional[str]:
        """Generate a chart from a ChartCandidate and return the file path."""
        filepath = os.path.join(self.output_dir, f"chart_{chart_index}.png")
        try:
            table = candidate.table
            if not table.rows or not table.headers:
                return None
                
            # Attempt AI Generated Chart first for unique attractiveness
            try:
                from .agents.chart_agent import ChartAgent
                print(f"  [ChartGenerator] Summoning AI Chart Agent for dynamic, unique {candidate.chart_type} chart...")
                agent = ChartAgent()
                code = agent.generate_chart_code(
                    candidate, 
                    filepath, 
                    theme_colors=self.design.colors.chart_colors()
                )
                
                # Execute AI code in a controlled namespace
                local_vars = {"filepath": filepath}
                exec(code, {}, local_vars)
                
                if os.path.exists(filepath):
                    return filepath
            except Exception as ai_e:
                print(f"  [Warning] AI Chart generation failed ({ai_e}), falling back to static rules.")

            # Fallback to static logic
            chart_type = candidate.chart_type
            colors = self.design.colors.chart_colors()

            if chart_type == "bar":
                return self._generate_bar_chart_3d(candidate, colors, chart_index)
            elif chart_type == "line":
                return self._generate_line_chart_3d(candidate, colors, chart_index)
            else:
                return self._generate_bar_chart_3d(candidate, colors, chart_index)

        except Exception as e:
            print(f"  [Warning] Chart generation failed: {e}")
            return None

    def _parse_value(self, cell: str) -> Optional[float]:
        """Parse a cell value to a float."""
        cleaned = cell.strip()
        if not cleaned or cleaned.upper() in ('N/A', '-', '—', 'N/D', ''):
            return None

        # Remove common formatting
        cleaned = cleaned.replace(',', '').replace('$', '').replace('€', '')
        cleaned = cleaned.replace('%', '').replace('+', '').replace('~', '')
        cleaned = cleaned.replace('≈', '').replace('<', '').replace('>', '')

        # Handle "billion" / "million"
        multiplier = 1.0
        if 'billion' in cleaned.lower():
            multiplier = 1e9
            cleaned = cleaned.lower().replace('billion', '').strip()
        elif 'million' in cleaned.lower():
            multiplier = 1e6
            cleaned = cleaned.lower().replace('million', '').strip()
        elif 'trillion' in cleaned.lower():
            multiplier = 1e12
            cleaned = cleaned.lower().replace('trillion', '').strip()

        try:
            return float(cleaned) * multiplier
        except (ValueError, TypeError):
            return None

    def _generate_bar_chart_3d(self, candidate: ChartCandidate,
                            colors: List[str], idx: int) -> Optional[str]:
        """Generate a 3D bar chart."""
        table = candidate.table
        labels = []
        data_series = {}

        for row in table.rows:
            if candidate.label_column < len(row):
                label = row[candidate.label_column].strip()
                if len(label) > 15:
                    label = label[:12] + "..."
                labels.append(label)

        for col_idx in candidate.data_columns:
            if col_idx < len(table.headers):
                col_name = table.headers[col_idx].strip()
                values = []
                for row in table.rows:
                    if col_idx < len(row):
                        val = self._parse_value(row[col_idx])
                        values.append(val if val is not None else 0)
                    else:
                        values.append(0)
                data_series[col_name] = values

        if not labels or not data_series:
            return None

        fig = plt.figure(figsize=(11, 5.5), dpi=150)
        ax = fig.add_subplot(111, projection='3d')
        ax.view_init(elev=20., azim=-35) # Professional corporate angle

        x = np.arange(len(labels))
        num_series = len(data_series)
        width = 0.6 / max(num_series, 1)

        import matplotlib.patches as mpatches
        legend_patches = []

        for i, (name, values) in enumerate(data_series.items()):
            offset = (i - num_series / 2 + 0.5) * width
            color = colors[i % len(colors)]
            
            y_pos = np.zeros(len(x))
            z_pos = np.zeros(len(x))
            dx = np.ones(len(x)) * width * 0.8
            dy = np.ones(len(x)) * 0.4
            
            # Using alpha for stylized 3D look
            ax.bar3d(x + offset, y_pos, z_pos, dx, dy, values, color=color, alpha=0.85, shade=True)
            legend_patches.append(mpatches.Patch(color=color, label=name))

        # Hide y axis to look like pure 3d bars mapped on x-z plane
        ax.set_yticks([]) 
        
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=25, ha='right')
        
        ax.set_title(candidate.section_title or table.title or "Data Overview",
                    fontsize=16, fontweight='bold', color='#212121', pad=20)

        if candidate.y_label:
            ax.set_zlabel(candidate.y_label, fontsize=10, labelpad=10)

        if num_series > 1:
            ax.legend(handles=legend_patches, loc='upper right', framealpha=0.9, bbox_to_anchor=(1.1, 1))

        # Format Z axis numbers
        ax.zaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: self._format_value(x)))

        # Clean up background pane colors
        ax.xaxis.pane.fill = False
        ax.yaxis.pane.fill = False
        ax.zaxis.pane.fill = False
        ax.grid(b=True, color='grey', linestyle='-.', linewidth=0.3, alpha=0.2)

        plt.tight_layout()
        filepath = os.path.join(self.output_dir, f"chart_{idx}.png")
        fig.savefig(filepath, bbox_inches='tight', dpi=150, facecolor='white')
        plt.close(fig)
        return filepath

    def _generate_line_chart_3d(self, candidate: ChartCandidate,
                             colors: List[str], idx: int) -> Optional[str]:
        """Generate a 3D line chart tracking trends at different depths."""
        table = candidate.table
        labels = []
        data_series = {}

        for row in table.rows:
            if candidate.label_column < len(row):
                labels.append(row[candidate.label_column].strip())

        for col_idx in candidate.data_columns:
            if col_idx < len(table.headers):
                col_name = table.headers[col_idx].strip()
                values = []
                for row in table.rows:
                    if col_idx < len(row):
                        val = self._parse_value(row[col_idx])
                        values.append(val if val is not None else 0)
                    else:
                        values.append(0)
                data_series[col_name] = values

        if not labels or not data_series:
            return None

        fig = plt.figure(figsize=(11, 5.5), dpi=150)
        ax = fig.add_subplot(111, projection='3d')
        ax.view_init(elev=15., azim=-45) # Slight tilt for tracking lines forward

        x = np.arange(len(labels))

        for i, (name, values) in enumerate(data_series.items()):
            color = colors[i % len(colors)]
            # Give each line a slightly different depth 'y'
            y = np.ones(len(x)) * i
            
            # Plot 3D line
            ax.plot3D(x, y, values, color=color, linewidth=3.5, label=name, marker='o', markersize=6)
            
            # Dropdown shadows (fill between line and floor)
            for j in range(len(x) - 1):
                # We can't easily natively fill_between in 3d, so we skip complex polygon filling to prevent bugs
                pass

        ax.set_yticks(np.arange(len(data_series)))
        ax.set_yticklabels(list(data_series.keys()), rotation=-15, ha='left')
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=25, ha='right')

        ax.set_title(candidate.section_title or table.title or "Trend Analysis",
                    fontsize=16, fontweight='bold', color='#212121', pad=20)

        ax.zaxis.set_major_formatter(mticker.FuncFormatter(lambda z, _: self._format_value(z)))

        ax.xaxis.pane.fill = False
        ax.yaxis.pane.fill = False
        ax.zaxis.pane.fill = False
        ax.grid(b=True, color='grey', linestyle='-.', linewidth=0.3, alpha=0.3)

        plt.tight_layout()
        filepath = os.path.join(self.output_dir, f"chart_{idx}.png")
        fig.savefig(filepath, bbox_inches='tight', dpi=150, facecolor='white')
        plt.close(fig)
        return filepath

    def _format_value(self, value: float) -> str:
        """Format a numeric value for display."""
        if value == 0:
            return "0"
        abs_val = abs(value)
        if abs_val >= 1e12:
            return f"{value/1e12:.1f}T"
        elif abs_val >= 1e9:
            return f"{value/1e9:.1f}B"
        elif abs_val >= 1e6:
            return f"{value/1e6:.1f}M"
        elif abs_val >= 1e3:
            return f"{value/1e3:.1f}K"
        elif abs_val < 1:
            return f"{value:.2f}"
        else:
            return f"{value:.1f}"
