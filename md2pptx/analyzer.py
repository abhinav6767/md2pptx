"""
Content Analyzer — Analyzes parsed document content to detect data types,
key insights, and content suitable for visualization.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple
from .parser import Document, Section, TableData, ContentBlock, BlockType


@dataclass
class ChartCandidate:
    """A table identified as suitable for chart generation."""
    table: TableData
    chart_type: str = "bar"  # bar, pie, line, area
    section_title: str = ""
    x_label: str = ""
    y_label: str = ""
    data_columns: List[int] = field(default_factory=list)
    label_column: int = 0


@dataclass
class InfographicCandidate:
    """Content identified as suitable for infographic generation."""
    infographic_type: str = "process"  # process, timeline, comparison, metrics
    title: str = ""
    items: List[str] = field(default_factory=list)
    values: List[str] = field(default_factory=list)
    section_title: str = ""


@dataclass
class KeyMetric:
    """An extracted key metric/statistic."""
    value: str = ""
    label: str = ""
    context: str = ""


@dataclass
class AnalysisResult:
    """Complete analysis of a document."""
    chart_candidates: List[ChartCandidate] = field(default_factory=list)
    infographic_candidates: List[InfographicCandidate] = field(default_factory=list)
    key_metrics: List[KeyMetric] = field(default_factory=list)
    section_importance: Dict[str, float] = field(default_factory=dict)
    total_word_count: int = 0
    has_toc: bool = False
    has_exec_summary: bool = False
    has_conclusion: bool = False


class ContentAnalyzer:
    """Analyzes document content for visualization opportunities and key insights."""

    # Patterns for detecting metrics
    METRIC_RE = re.compile(
        r'(\$?\d+[\d,.]*\s*(?:billion|million|%|percent|trillion)?)',
        re.IGNORECASE
    )
    PERCENTAGE_RE = re.compile(r'(\d+(?:\.\d+)?)\s*%')
    CURRENCY_RE = re.compile(r'\$\s*(\d+(?:\.\d+)?)\s*(billion|million|trillion)?', re.IGNORECASE)

    # Keywords indicating process/flow content
    PROCESS_KEYWORDS = [
        "step", "phase", "stage", "process", "workflow", "pipeline",
        "first", "second", "third", "then", "next", "finally"
    ]
    TIMELINE_KEYWORDS = [
        "timeline", "chronolog", "history", "evolution", "year",
        "2020", "2021", "2022", "2023", "2024", "2025"
    ]
    COMPARISON_KEYWORDS = [
        "comparison", "versus", "vs", "compare", "differ", "advantage",
        "disadvantage", "pro", "con", "strength", "weakness"
    ]

    def analyze(self, doc: Document) -> AnalysisResult:
        """Perform complete analysis of the document."""
        result = AnalysisResult()

        # Check structural elements
        result.has_toc = doc.get_toc_section() is not None
        result.has_exec_summary = doc.get_executive_summary() is not None
        result.has_conclusion = doc.get_conclusion() is not None

        # Analyze tables for charts
        result.chart_candidates = self._find_chart_candidates(doc)

        # Analyze content for infographics
        result.infographic_candidates = self._find_infographic_candidates(doc)

        # Extract key metrics
        result.key_metrics = self._extract_key_metrics(doc)

        # Calculate section importance
        result.section_importance = self._calculate_importance(doc)

        # Total word count
        for sec in doc.sections:
            result.total_word_count += sec.word_count()

        return result

    def _find_chart_candidates(self, doc: Document) -> List[ChartCandidate]:
        """Find tables suitable for chart generation."""
        candidates = []

        for table in doc.all_tables:
            if not table.has_numerical_data:
                continue
            if len(table.rows) < 2:
                continue

            candidate = ChartCandidate(
                table=table,
                section_title=table.title or "Data"
            )

            # Determine chart type based on data shape
            num_rows = len(table.rows)
            num_cols = len(table.headers)

            # Find numerical columns
            for col_idx in range(num_cols):
                has_numbers = False
                for row in table.rows:
                    if col_idx < len(row):
                        cleaned = re.sub(r'[,$%~+≈<>()N/A\s]', '', row[col_idx])
                        cleaned = cleaned.replace('billion', '').replace('million', '').strip()
                        try:
                            float(cleaned)
                            has_numbers = True
                            break
                        except (ValueError, TypeError):
                            continue
                if has_numbers:
                    candidate.data_columns.append(col_idx)

            if not candidate.data_columns:
                continue

            # Set label column (first non-numerical column)
            for col_idx in range(num_cols):
                if col_idx not in candidate.data_columns:
                    candidate.label_column = col_idx
                    break

            # Determine chart type
            if num_rows <= 5 and len(candidate.data_columns) == 1:
                # Check if data looks like parts of a whole
                candidate.chart_type = "bar"
            elif num_rows > 5:
                candidate.chart_type = "bar"
            elif len(candidate.data_columns) >= 2:
                candidate.chart_type = "bar"  # Grouped bar for multiple columns
            else:
                candidate.chart_type = "bar"

            # Check for time-series data (line chart)
            if table.headers:
                first_header = table.headers[0].lower()
                if any(kw in first_header for kw in ["year", "quarter", "month", "date", "period", "fiscal"]):
                    candidate.chart_type = "line"

            candidate.x_label = table.headers[candidate.label_column] if table.headers else ""
            if candidate.data_columns:
                candidate.y_label = table.headers[candidate.data_columns[0]] if len(table.headers) > candidate.data_columns[0] else ""

            candidates.append(candidate)

        return candidates

    def _find_infographic_candidates(self, doc: Document) -> List[InfographicCandidate]:
        """Find content suitable for infographic visualization."""
        candidates = []
        content_sections = doc.get_content_sections()

        for sec in content_sections:
            sec_text = sec.get_text_content().lower()

            # Check for process/flow content
            if self._has_keywords(sec_text, self.PROCESS_KEYWORDS):
                # Look for numbered or bullet lists that describe steps
                for block in sec.content_blocks:
                    if block.block_type in (BlockType.NUMBERED_LIST, BlockType.BULLET_LIST):
                        if 3 <= len(block.items) <= 7:
                            candidates.append(InfographicCandidate(
                                infographic_type="process",
                                title=sec.title,
                                items=block.items[:6],
                                section_title=sec.title,
                            ))
                            break

            # Check subsections for similar patterns
            for sub in sec.subsections:
                sub_text = sub.get_text_content().lower()

                if self._has_keywords(sub_text, self.COMPARISON_KEYWORDS):
                    for block in sub.content_blocks:
                        if block.block_type in (BlockType.BULLET_LIST, BlockType.NUMBERED_LIST):
                            if 2 <= len(block.items) <= 6:
                                candidates.append(InfographicCandidate(
                                    infographic_type="comparison",
                                    title=sub.title,
                                    items=block.items[:6],
                                    section_title=sec.title,
                                ))
                                break

        return candidates[:3]  # Limit to 3 infographics max

    def _extract_key_metrics(self, doc: Document) -> List[KeyMetric]:
        """Extract key metrics and statistics from the document."""
        metrics = []
        exec_summary = doc.get_executive_summary()

        if exec_summary:
            text = exec_summary.get_text_content()
            # Find percentages
            for m in self.PERCENTAGE_RE.finditer(text):
                start = max(0, m.start() - 60)
                end = min(len(text), m.end() + 30)
                context = text[start:end].strip()
                metrics.append(KeyMetric(
                    value=m.group(0),
                    label=self._extract_metric_label(text, m.start()),
                    context=context
                ))

            # Find currency values
            for m in self.CURRENCY_RE.finditer(text):
                start = max(0, m.start() - 60)
                end = min(len(text), m.end() + 30)
                context = text[start:end].strip()
                metrics.append(KeyMetric(
                    value=m.group(0),
                    label=self._extract_metric_label(text, m.start()),
                    context=context
                ))

        return metrics[:8]  # Limit to 8 key metrics

    def _extract_metric_label(self, text: str, pos: int) -> str:
        """Extract a short label for a metric from surrounding text."""
        start = max(0, pos - 50)
        chunk = text[start:pos].strip()
        # Get last few words before the number
        words = chunk.split()[-5:]
        return " ".join(words)

    def _calculate_importance(self, doc: Document) -> Dict[str, float]:
        """Calculate importance scores for each section."""
        importance = {}
        content_sections = doc.get_content_sections()

        if not content_sections:
            return importance

        # Calculate based on word count, tables, and position
        total_words = sum(sec.word_count() for sec in content_sections)
        if total_words == 0:
            total_words = 1

        for i, sec in enumerate(content_sections):
            score = 0.0

            # Word count factor (normalized)
            word_ratio = sec.word_count() / total_words
            score += min(word_ratio * 3, 1.0)  # Cap at 1.0

            # Position factor (earlier sections slightly more important)
            position_factor = 1.0 - (i / max(len(content_sections), 1)) * 0.3
            score *= position_factor

            # Tables boost importance
            tables = sec.get_tables()
            if tables:
                score += 0.3
                if any(t.has_numerical_data for t in tables):
                    score += 0.2

            # Subsections add depth
            if sec.subsections:
                score += min(len(sec.subsections) * 0.1, 0.3)

            importance[sec.title] = round(score, 3)

        return importance

    def _has_keywords(self, text: str, keywords: List[str]) -> bool:
        """Check if text contains any of the keywords."""
        return any(kw in text for kw in keywords)
