"""
Storyline Builder — Constructs a coherent 10-15 slide narrative from parsed content.
Maps content to slide types following the mandated flow:
Title → Agenda → Executive Summary → Section Content → Charts & Data → Conclusion
"""

from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from enum import Enum, auto

from .parser import Document, Section, ContentBlock, BlockType, TableData
from .analyzer import AnalysisResult, ChartCandidate, InfographicCandidate
from .design import MAX_SLIDES, MIN_SLIDES, MAX_BULLETS_PER_SLIDE, MAX_CHARS_PER_BULLET


class SlideType(Enum):
    COVER = auto()
    AGENDA = auto()
    EXECUTIVE_SUMMARY = auto()
    SECTION_DIVIDER = auto()
    CONTENT = auto()
    CONTENT_TWO_COLUMN = auto()
    CHART = auto()
    TABLE = auto()
    INFOGRAPHIC = auto()
    KEY_METRICS = auto()
    CONCLUSION = auto()
    THANK_YOU = auto()


@dataclass
class SlideContent:
    """Content for a single slide in the storyline."""
    slide_type: SlideType
    title: str = ""
    subtitle: str = ""
    body_text: str = ""
    bullets: List[str] = field(default_factory=list)
    table: Optional[TableData] = None
    chart_candidate: Optional[ChartCandidate] = None
    infographic: Optional[InfographicCandidate] = None
    left_bullets: List[str] = field(default_factory=list)  # For two-column
    right_bullets: List[str] = field(default_factory=list)  # For two-column
    left_title: str = ""
    right_title: str = ""
    section_number: int = 0


@dataclass
class Storyline:
    """Complete slide storyline."""
    slides: List[SlideContent] = field(default_factory=list)
    document_title: str = ""
    document_subtitle: str = ""

    @property
    def slide_count(self) -> int:
        return len(self.slides)


class StorylineBuilder:
    """Builds a storyline from parsed document and analysis results."""

    def __init__(self, target_slides: int = 12):
        self.target_slides = max(MIN_SLIDES, min(target_slides, MAX_SLIDES))

    def build(self, doc: Document, analysis: AnalysisResult) -> Storyline:
        """Build a complete storyline."""
        storyline = Storyline(
            document_title=doc.title,
            document_subtitle=doc.subtitle
        )

        # 1. Title slide (always)
        storyline.slides.append(SlideContent(
            slide_type=SlideType.COVER,
            title=doc.title,
            subtitle=doc.subtitle or self._generate_subtitle(doc),
        ))

        # 2. Agenda slide (if TOC exists and we have room)
        agenda_items = self._build_agenda(doc)
        if agenda_items and self.target_slides >= 12:
            storyline.slides.append(SlideContent(
                slide_type=SlideType.AGENDA,
                title="Agenda",
                bullets=agenda_items[:MAX_BULLETS_PER_SLIDE],
            ))

        # 3. Executive Summary (if exists)
        exec_summary = doc.get_executive_summary()
        if exec_summary:
            summary_bullets = self._summarize_section(exec_summary, max_bullets=MAX_BULLETS_PER_SLIDE)
            storyline.slides.append(SlideContent(
                slide_type=SlideType.EXECUTIVE_SUMMARY,
                title="Executive Summary",
                bullets=summary_bullets,
            ))

        # 4. Calculate remaining slide budget
        used = len(storyline.slides)
        # Reserve: 1 for conclusion, 1 for thank you
        reserved = 2
        # Charts/data slides
        chart_slides = min(len(analysis.chart_candidates), 2)
        # Infographic slides
        infographic_slides = min(len(analysis.infographic_candidates), 1)

        content_budget = self.target_slides - used - reserved - chart_slides - infographic_slides
        content_budget = max(content_budget, 3)  # At least 3 content slides

        # 5. Section content slides
        content_sections = doc.get_content_sections()
        section_slides = self._build_content_slides(
            content_sections, analysis, content_budget
        )
        storyline.slides.extend(section_slides)

        # 6. Chart & Data slides
        for i, chart in enumerate(analysis.chart_candidates[:chart_slides]):
            storyline.slides.append(SlideContent(
                slide_type=SlideType.CHART,
                title=chart.section_title or f"Data Analysis {i+1}",
                chart_candidate=chart,
            ))

        # 7. Infographic slides
        for infographic in analysis.infographic_candidates[:infographic_slides]:
            storyline.slides.append(SlideContent(
                slide_type=SlideType.INFOGRAPHIC,
                title=infographic.title,
                infographic=infographic,
            ))

        # 8. Conclusion / Key Takeaways
        conclusion = doc.get_conclusion()
        if conclusion:
            conclusion_bullets = self._summarize_section(conclusion, max_bullets=MAX_BULLETS_PER_SLIDE)
            storyline.slides.append(SlideContent(
                slide_type=SlideType.CONCLUSION,
                title="Key Takeaways",
                bullets=conclusion_bullets,
            ))
        else:
            # Generate conclusion from key metrics
            storyline.slides.append(SlideContent(
                slide_type=SlideType.CONCLUSION,
                title="Key Takeaways",
                bullets=self._generate_conclusion(doc, analysis),
            ))

        # 9. Thank You slide (always)
        storyline.slides.append(SlideContent(
            slide_type=SlideType.THANK_YOU,
            title="Thank You",
        ))

        # Ensure we're within bounds
        self._trim_to_target(storyline)

        return storyline

    def _generate_subtitle(self, doc: Document) -> str:
        """Generate a subtitle if none exists."""
        exec_sum = doc.get_executive_summary()
        if exec_sum:
            text = exec_sum.get_text_content()
            # Take first sentence
            sentences = text.split('.')
            if sentences:
                return sentences[0].strip()[:120]
        return ""

    def _build_agenda(self, doc: Document) -> List[str]:
        """Build agenda items from main sections."""
        content_sections = doc.get_content_sections()
        items = []
        for sec in content_sections:
            title = sec.title.strip()
            # Remove numbering prefix like "1. " or "1.1. "
            title = self._strip_numbering(title)
            if title and len(title) > 3:
                items.append(title)
        return items[:8]

    def _summarize_section(self, section: Section, max_bullets: int = 6) -> List[str]:
        """Summarize a section into bullet points."""
        bullets = []

        # First, collect existing bullet points
        for block in section.content_blocks:
            if block.block_type in (BlockType.BULLET_LIST, BlockType.NUMBERED_LIST):
                for item in block.items:
                    cleaned = self._truncate(item, MAX_CHARS_PER_BULLET)
                    if cleaned and len(cleaned) > 10:
                        bullets.append(cleaned)

        # If not enough bullets, extract from paragraphs
        if len(bullets) < 3:
            for block in section.content_blocks:
                if block.block_type == BlockType.PARAGRAPH and block.text:
                    sentences = self._extract_key_sentences(block.text)
                    for sent in sentences:
                        cleaned = self._truncate(sent, MAX_CHARS_PER_BULLET)
                        if cleaned and cleaned not in bullets and len(cleaned) > 15:
                            bullets.append(cleaned)

        # Also check subsections
        for sub in section.subsections:
            sub_bullets = self._summarize_section(sub, max_bullets=2)
            bullets.extend(sub_bullets)

        # Deduplicate and limit
        seen = set()
        unique_bullets = []
        for b in bullets:
            key = b[:50].lower()
            if key not in seen:
                seen.add(key)
                unique_bullets.append(b)

        return unique_bullets[:max_bullets]

    def _build_content_slides(self, sections: List[Section],
                               analysis: AnalysisResult,
                               budget: int) -> List[SlideContent]:
        """Build content slides from sections within the given budget."""
        slides = []

        if not sections:
            return slides

        # Sort sections by importance
        sorted_sections = sorted(
            sections,
            key=lambda s: analysis.section_importance.get(s.title, 0),
            reverse=True
        )

        # Select top sections that fit in budget
        selected = sorted_sections[:budget]

        # Re-order by original position
        original_order = {sec.title: i for i, sec in enumerate(sections)}
        selected.sort(key=lambda s: original_order.get(s.title, 999))

        section_num = 1
        for sec in selected:
            # Check if section has enough content for a divider + content
            has_subs = len(sec.subsections) > 0

            # Build content bullets from section + its subsections
            all_bullets = []

            # Direct content
            for block in sec.content_blocks:
                if block.block_type in (BlockType.BULLET_LIST, BlockType.NUMBERED_LIST):
                    for item in block.items:
                        cleaned = self._truncate(item, MAX_CHARS_PER_BULLET)
                        if cleaned and len(cleaned) > 10:
                            all_bullets.append(cleaned)
                elif block.block_type == BlockType.PARAGRAPH and block.text:
                    sents = self._extract_key_sentences(block.text)
                    for s in sents[:2]:
                        cleaned = self._truncate(s, MAX_CHARS_PER_BULLET)
                        if cleaned and len(cleaned) > 15:
                            all_bullets.append(cleaned)

            # Subsection content
            for sub in sec.subsections:
                sub_content = self._summarize_section(sub, max_bullets=2)
                all_bullets.extend(sub_content)

            # Deduplicate
            seen = set()
            unique = []
            for b in all_bullets:
                key = b[:40].lower()
                if key not in seen:
                    seen.add(key)
                    unique.append(b)
            all_bullets = unique

            # Check for tables
            section_tables = sec.get_tables()
            for sub in sec.subsections:
                section_tables.extend(sub.get_tables())

            # Create slide
            title = self._strip_numbering(sec.title)

            if section_tables and section_tables[0].rows:
                # Table slide
                table = section_tables[0]
                # Limit table rows for display
                if len(table.rows) > 8:
                    table = TableData(
                        title=table.title,
                        headers=table.headers,
                        rows=table.rows[:8],
                        has_numerical_data=table.has_numerical_data
                    )
                slides.append(SlideContent(
                    slide_type=SlideType.TABLE,
                    title=title,
                    table=table,
                    section_number=section_num,
                ))
            elif len(all_bullets) > MAX_BULLETS_PER_SLIDE and has_subs:
                # Two-column layout for dense content
                mid = len(all_bullets) // 2
                left = all_bullets[:mid][:MAX_BULLETS_PER_SLIDE]
                right = all_bullets[mid:][:MAX_BULLETS_PER_SLIDE]
                slides.append(SlideContent(
                    slide_type=SlideType.CONTENT_TWO_COLUMN,
                    title=title,
                    left_bullets=left,
                    right_bullets=right,
                    section_number=section_num,
                ))
            else:
                slides.append(SlideContent(
                    slide_type=SlideType.CONTENT,
                    title=title,
                    bullets=all_bullets[:MAX_BULLETS_PER_SLIDE],
                    section_number=section_num,
                ))

            section_num += 1

        return slides

    def _generate_conclusion(self, doc: Document, analysis: AnalysisResult) -> List[str]:
        """Generate conclusion bullets from analysis."""
        bullets = []
        # Use key metrics
        for metric in analysis.key_metrics[:4]:
            if metric.context:
                bullets.append(self._truncate(metric.context, MAX_CHARS_PER_BULLET))

        if not bullets:
            bullets = [
                f"Analysis covers {len(doc.sections)} major sections",
                f"Total content: {analysis.total_word_count:,} words analyzed",
            ]
            if analysis.chart_candidates:
                bullets.append(f"{len(analysis.chart_candidates)} data visualizations generated")

        return bullets[:MAX_BULLETS_PER_SLIDE]

    def _trim_to_target(self, storyline: Storyline):
        """Trim storyline to target slide count if over."""
        while len(storyline.slides) > MAX_SLIDES:
            # Remove content slides from the middle (keep cover, exec summary, conclusion, thank you)
            protected_types = {SlideType.COVER, SlideType.EXECUTIVE_SUMMARY,
                             SlideType.CONCLUSION, SlideType.THANK_YOU}
            for i in range(len(storyline.slides) - 1, -1, -1):
                if storyline.slides[i].slide_type not in protected_types:
                    storyline.slides.pop(i)
                    break
            else:
                break  # Can't remove any more

    def _extract_key_sentences(self, text: str) -> List[str]:
        """Extract key sentences from a paragraph."""
        sentences = []
        # Split on sentence boundaries
        parts = text.replace('. ', '.\n').split('\n')
        for part in parts:
            part = part.strip()
            if len(part) > 20:  # Skip very short fragments
                sentences.append(part)
        return sentences[:3]

    def _truncate(self, text: str, max_len: int) -> str:
        """Truncate text to max length, preserving whole words."""
        if len(text) <= max_len:
            return text
        truncated = text[:max_len]
        # Find last space
        last_space = truncated.rfind(' ')
        if last_space > max_len * 0.6:
            truncated = truncated[:last_space]
        return truncated.rstrip('.,;:') + "..."

    def _strip_numbering(self, title: str) -> str:
        """Remove section numbering from title."""
        import re
        # Remove patterns like "1. ", "1.1. ", "1.2.3 "
        cleaned = re.sub(r'^\d+(\.\d+)*\.?\s+', '', title)
        return cleaned.strip()
