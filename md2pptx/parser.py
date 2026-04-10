"""
Markdown Parser — Parses .md files into a structured document model.
Handles headings, paragraphs, lists, tables, images, and references.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from enum import Enum, auto


class BlockType(Enum):
    HEADING = auto()
    PARAGRAPH = auto()
    BULLET_LIST = auto()
    NUMBERED_LIST = auto()
    TABLE = auto()
    IMAGE = auto()
    BLOCKQUOTE = auto()
    CODE_BLOCK = auto()


@dataclass
class TableData:
    """Parsed markdown table."""
    title: str = ""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    has_numerical_data: bool = False

    def __post_init__(self):
        self._check_numerical()

    def _check_numerical(self):
        """Check if table contains numerical data suitable for charting."""
        for row in self.rows:
            for cell in row:
                cleaned = re.sub(r'[,$%~+≈<>()]', '', cell.strip())
                cleaned = cleaned.replace('billion', '').replace('million', '').replace('N/A', '').strip()
                try:
                    float(cleaned)
                    self.has_numerical_data = True
                    return
                except (ValueError, TypeError):
                    continue


@dataclass
class ContentBlock:
    """A single content block within a section."""
    block_type: BlockType
    text: str = ""
    items: List[str] = field(default_factory=list)  # For lists
    table: Optional[TableData] = None
    level: int = 0  # Heading level or list nesting
    image_data: str = ""  # Base64 or path


@dataclass
class Section:
    """A document section starting with a heading."""
    title: str = ""
    level: int = 1  # h1=1, h2=2, etc.
    content_blocks: List[ContentBlock] = field(default_factory=list)
    subsections: List['Section'] = field(default_factory=list)

    def get_text_content(self) -> str:
        """Get all text content concatenated."""
        parts = []
        for block in self.content_blocks:
            if block.block_type == BlockType.PARAGRAPH:
                parts.append(block.text)
            elif block.block_type in (BlockType.BULLET_LIST, BlockType.NUMBERED_LIST):
                parts.extend(block.items)
        return " ".join(parts)

    def get_tables(self) -> List[TableData]:
        """Get all tables in this section."""
        tables = []
        for block in self.content_blocks:
            if block.block_type == BlockType.TABLE and block.table:
                tables.append(block.table)
        return tables

    def word_count(self) -> int:
        """Estimate word count for content density."""
        count = len(self.get_text_content().split())
        for sub in self.subsections:
            count += sub.word_count()
        return count


@dataclass
class Document:
    """Parsed markdown document."""
    title: str = ""
    subtitle: str = ""
    sections: List[Section] = field(default_factory=list)
    all_tables: List[TableData] = field(default_factory=list)

    def get_toc_section(self) -> Optional[Section]:
        """Find the Table of Contents section."""
        for sec in self.sections:
            if "table of contents" in sec.title.lower() or "toc" in sec.title.lower():
                return sec
        return None

    def get_executive_summary(self) -> Optional[Section]:
        """Find the Executive Summary section."""
        for sec in self.sections:
            title_lower = sec.title.lower()
            if "executive summary" in title_lower or "summary" in title_lower:
                return sec
        return None

    def get_conclusion(self) -> Optional[Section]:
        """Find the Conclusion section."""
        for sec in self.sections:
            title_lower = sec.title.lower()
            if any(kw in title_lower for kw in ["conclusion", "key takeaway",
                                                   "key insight", "recommendation",
                                                   "strategic implication"]):
                return sec
        return None

    def get_content_sections(self) -> List[Section]:
        """Get main content sections (excluding TOC, exec summary, conclusion, references)."""
        skip_keywords = [
            "table of contents", "toc", "executive summary", "summary",
            "conclusion", "key takeaway", "key insight", "reference",
            "source documentation", "recommendation", "strategic implication"
        ]
        result = []
        for sec in self.sections:
            title_lower = sec.title.lower()
            if not any(kw in title_lower for kw in skip_keywords):
                result.append(sec)
        return result


class MarkdownParser:
    """Parses Markdown text into a structured Document."""

    # Regex patterns
    HEADING_RE = re.compile(r'^(#{1,6})\s+(.+)$', re.MULTILINE)
    TABLE_TITLE_RE = re.compile(r'^Title:\s*(.+)$', re.MULTILINE)
    TABLE_ROW_RE = re.compile(r'^\|(.+)\|$')
    TABLE_SEP_RE = re.compile(r'^\|[\s:|-]+\|$')
    BULLET_RE = re.compile(r'^(\s*)[-*+]\s+(.+)$')
    NUMBERED_RE = re.compile(r'^(\s*)\d+[.)]\s+(.+)$')
    IMAGE_RE = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
    LINK_RE = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
    BLOCKQUOTE_RE = re.compile(r'^>\s*(.+)$')

    def parse(self, text: str) -> Document:
        """Parse markdown text into a Document."""
        doc = Document()
        lines = text.split('\n')

        # Clean up: remove excessive blank lines
        cleaned_lines = []
        prev_blank = False
        for line in lines:
            is_blank = line.strip() == ''
            if is_blank and prev_blank:
                continue
            cleaned_lines.append(line)
            prev_blank = is_blank

        lines = cleaned_lines

        # First pass: extract title from first H1
        for line in lines:
            m = self.HEADING_RE.match(line)
            if m and len(m.group(1)) == 1:
                doc.title = self._clean_text(m.group(2))
                break

        # Second pass: find subtitle (first H3 after title, if any)
        found_title = False
        for line in lines:
            m = self.HEADING_RE.match(line)
            if m:
                level = len(m.group(1))
                if level == 1:
                    found_title = True
                    continue
                if found_title and level == 3:
                    doc.subtitle = self._clean_text(m.group(2))
                    break
                if found_title and level == 2:
                    break  # No subtitle

        # Third pass: build sections
        current_section = None
        current_subsection = None
        i = 0
        pending_table_title = ""

        while i < len(lines):
            line = lines[i]

            # Check for heading
            m = self.HEADING_RE.match(line)
            if m:
                level = len(m.group(1))
                title = self._clean_text(m.group(2))

                if level == 1:
                    # Skip the document title
                    i += 1
                    continue
                elif level == 2:
                    if current_section:
                        doc.sections.append(current_section)
                    current_section = Section(title=title, level=level)
                    current_subsection = None
                elif level >= 3 and current_section:
                    current_subsection = Section(title=title, level=level)
                    current_section.subsections.append(current_subsection)

                i += 1
                continue

            # Check for table title line (e.g., "Title: ...")
            tm = self.TABLE_TITLE_RE.match(line)
            if tm:
                pending_table_title = tm.group(1).strip()
                i += 1
                continue

            # Check for table
            if self.TABLE_ROW_RE.match(line.strip()):
                table, end_idx = self._parse_table(lines, i, pending_table_title)
                pending_table_title = ""
                if table:
                    block = ContentBlock(block_type=BlockType.TABLE, table=table)
                    target = current_subsection if current_subsection else current_section
                    if target:
                        target.content_blocks.append(block)
                    doc.all_tables.append(table)
                i = end_idx
                continue

            # Check for bullet list
            bm = self.BULLET_RE.match(line)
            if bm:
                items, end_idx = self._parse_list(lines, i, is_numbered=False)
                block = ContentBlock(block_type=BlockType.BULLET_LIST, items=items)
                target = current_subsection if current_subsection else current_section
                if target:
                    target.content_blocks.append(block)
                i = end_idx
                continue

            # Check for numbered list
            nm = self.NUMBERED_RE.match(line)
            if nm:
                items, end_idx = self._parse_list(lines, i, is_numbered=True)
                block = ContentBlock(block_type=BlockType.NUMBERED_LIST, items=items)
                target = current_subsection if current_subsection else current_section
                if target:
                    target.content_blocks.append(block)
                i = end_idx
                continue

            # Check for image (skip base64 images)
            img_m = self.IMAGE_RE.search(line)
            if img_m and not img_m.group(2).startswith("data:"):
                block = ContentBlock(
                    block_type=BlockType.IMAGE,
                    text=img_m.group(1),
                    image_data=img_m.group(2)
                )
                target = current_subsection if current_subsection else current_section
                if target:
                    target.content_blocks.append(block)
                i += 1
                continue

            # Skip base64 images
            if "data:image" in line:
                i += 1
                continue

            # Regular paragraph
            text = line.strip()
            if text and current_section:
                # Clean references like [1](url) from text
                text = self._clean_references(text)
                if text:
                    block = ContentBlock(block_type=BlockType.PARAGRAPH, text=text)
                    target = current_subsection if current_subsection else current_section
                    if target:
                        target.content_blocks.append(block)

            i += 1

        # Add last section
        if current_section:
            doc.sections.append(current_section)

        return doc

    def _parse_table(self, lines: List[str], start: int,
                     title: str = "") -> Tuple[Optional[TableData], int]:
        """Parse a markdown table starting at the given line index."""
        table = TableData(title=title)
        i = start

        # Parse header row
        if i < len(lines) and self.TABLE_ROW_RE.match(lines[i].strip()):
            cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
            table.headers = [self._clean_text(c) for c in cells]
            i += 1
        else:
            return None, start + 1

        # Skip separator row
        if i < len(lines) and self.TABLE_SEP_RE.match(lines[i].strip()):
            i += 1

        # Parse data rows
        while i < len(lines):
            line = lines[i].strip()
            if not self.TABLE_ROW_RE.match(line):
                break
            if self.TABLE_SEP_RE.match(line):
                i += 1
                continue
            cells = [c.strip() for c in line.strip('|').split('|')]
            table.rows.append([self._clean_text(c) for c in cells])
            i += 1

        if table.headers:
            table._check_numerical()
            return table, i
        return None, start + 1

    def _parse_list(self, lines: List[str], start: int,
                    is_numbered: bool = False) -> Tuple[List[str], int]:
        """Parse a bullet or numbered list."""
        items = []
        pattern = self.NUMBERED_RE if is_numbered else self.BULLET_RE
        i = start

        while i < len(lines):
            line = lines[i]
            m = pattern.match(line)
            if m:
                item_text = self._clean_text(m.group(2))
                item_text = self._clean_references(item_text)
                if item_text:
                    items.append(item_text)
                i += 1
            elif line.strip() == '':
                # Check if list continues after blank line
                if i + 1 < len(lines) and pattern.match(lines[i + 1]):
                    i += 1
                else:
                    break
            elif line.startswith('  ') or line.startswith('\t'):
                # Continuation of previous item
                if items:
                    cont = self._clean_text(line.strip())
                    cont = self._clean_references(cont)
                    items[-1] += " " + cont
                i += 1
            else:
                break

        return items, i

    def _clean_text(self, text: str) -> str:
        """Clean markdown formatting from text."""
        # Remove bold/italic
        text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        text = re.sub(r'\*(.+?)\*', r'\1', text)
        text = re.sub(r'___(.+?)___', r'\1', text)
        text = re.sub(r'__(.+?)__', r'\1', text)
        text = re.sub(r'_(.+?)_', r'\1', text)
        # Remove inline code
        text = re.sub(r'`(.+?)`', r'\1', text)
        # Clean up extra spaces
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def _clean_references(self, text: str) -> str:
        """Remove citation references like [1](url) from text."""
        # Remove [N](url) patterns
        text = re.sub(r'\s*\[\d+\]\([^)]+\)', '', text)
        # Remove trailing references
        text = re.sub(r'\s*\[[\d,\s]+\]\([^)]+\)', '', text)
        return text.strip()


def parse_markdown_file(filepath: str) -> Document:
    """Parse a markdown file and return a Document."""
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        text = f.read()
    parser = MarkdownParser()
    return parser.parse(text)
