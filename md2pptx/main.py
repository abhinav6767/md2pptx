"""
MD2PPTX — Main CLI entrypoint
Converts Markdown (.md) files to structured, visually appealing .pptx presentations.

Usage:
    python -m md2pptx input.md --template template.pptx --output output.pptx
    python -m md2pptx input.md -t template.pptx -o output.pptx --slides 12
"""

import argparse
import os
import sys
import time

from .parser import parse_markdown_file
from .analyzer import ContentAnalyzer
from .storyline import StorylineBuilder
from .template import TemplateLoader, find_templates
from .renderer import PPTXRenderer


def main():
    parser = argparse.ArgumentParser(
        prog="md2pptx",
        description="Convert Markdown files to professional PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m md2pptx input.md -t template.pptx -o output.pptx
  python -m md2pptx input.md -t template.pptx --slides 12
  python -m md2pptx input.md -t template.pptx -o output/ --slides 15
        """
    )

    parser.add_argument("input", help="Path to the Markdown (.md) input file")
    parser.add_argument("-t", "--template", required=True,
                       help="Path to the Slide Master template (.pptx) file")
    parser.add_argument("-o", "--output", default=None,
                       help="Output .pptx file path (default: input_name.pptx in output/)")
    parser.add_argument("--slides", type=int, default=12,
                       help="Target number of slides (10-15, default: 12)")

    args = parser.parse_args()

    # Validate input
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        sys.exit(1)

    if not args.input.endswith('.md'):
        print(f"Error: Input file must be a .md file: {args.input}")
        sys.exit(1)

    # Check file size (max 5 MB)
    file_size = os.path.getsize(args.input)
    if file_size > 5 * 1024 * 1024:
        print(f"Warning: Input file is {file_size / (1024*1024):.1f} MB (max recommended: 5 MB)")
        print("  Processing may be slow for very large files.")

    if not os.path.exists(args.template):
        print(f"Error: Template file not found: {args.template}")
        sys.exit(1)

    # Determine output path
    if args.output is None:
        input_name = os.path.splitext(os.path.basename(args.input))[0]
        output_dir = os.path.join(os.path.dirname(args.input), "output")
        os.makedirs(output_dir, exist_ok=True)
        args.output = os.path.join(output_dir, f"{input_name}.pptx")
    elif os.path.isdir(args.output):
        input_name = os.path.splitext(os.path.basename(args.input))[0]
        args.output = os.path.join(args.output, f"{input_name}.pptx")

    # Validate slide count
    target_slides = max(10, min(args.slides, 15))

    print("=" * 60)
    print("  MD2PPTX — Markdown to PowerPoint Converter")
    print("=" * 60)
    print(f"  Input:    {args.input}")
    print(f"  Template: {args.template}")
    print(f"  Output:   {args.output}")
    print(f"  Target:   {target_slides} slides")
    print("=" * 60)

    start_time = time.time()

    # Step 1: Parse Markdown
    print("\n[1/5] Parsing Markdown...")
    doc = parse_markdown_file(args.input)
    print(f"  Title: {doc.title}")
    print(f"  Sections: {len(doc.sections)}")
    print(f"  Tables: {len(doc.all_tables)}")

    # Step 2: Agent Orchestration
    print("\n[2/5] Initiating Multi-Agent AI Pipeline...")
    from .agents import MultiAgentOrchestrator
    orchestrator = MultiAgentOrchestrator(target_slides=target_slides)
    
    # We still need the template loaded before orchestrator finishes so the Layout Agent knows available layouts
    print("\n[3/5] Loading template for Agent reference...")
    template_loader = TemplateLoader(args.template)
    info = template_loader.info
    
    storyline = orchestrator.run(doc, template_loader)

    # We already loaded the template in Step 3.
    # The layouts are ready in info.layouts.

    # Step 5: Render PPTX
    print("\n[5/5] Rendering .pptx...")
    renderer = PPTXRenderer(template_loader)
    output_path = renderer.render(storyline, doc, args.output)

    elapsed = time.time() - start_time
    print(f"\n{'=' * 60}")
    print(f"  Done! Generated {len(storyline.slides)} slides in {elapsed:.1f}s")
    print(f"  Output: {output_path}")
    print(f"{'=' * 60}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
