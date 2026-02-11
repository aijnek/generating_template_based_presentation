#!/usr/bin/env python3
"""
Presentation Creator CLI

Creates PowerPoint presentations from template and slide specifications.

Usage:
    python create_presentation.py \\
        --template template.pptx \\
        --slides slides.json \\
        --output output.pptx \\
        [--config config.json]
"""

import sys
import json
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.presentation_builder import (
    PresentationBuilder,
    TemplateConfig,
    create_simple_presentation
)


def load_slides_spec(spec_path):
    """Load slide specifications from JSON file"""
    with open(spec_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_from_files(template_path, slides_spec_path, output_path, config_path=None):
    """Create presentation from files"""
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not Path(slides_spec_path).exists():
        raise FileNotFoundError(f"Slides spec not found: {slides_spec_path}")
    
    print(f"Loading slides: {slides_spec_path}")
    slides = load_slides_spec(slides_spec_path)
    print(f"✓ Loaded {len(slides)} slide(s)\n")
    
    print(f"Creating presentation from: {template_path}")
    builder = create_simple_presentation(
        template_path=template_path,
        output_path=output_path,
        slides_data=slides,
        config_path=config_path
    )
    
    print(f"✓ Created {builder.slide_count} slide(s)")
    print(f"✓ Saved to: {output_path}")
    
    return builder


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Create PowerPoint presentation from template'
    )
    
    parser.add_argument('-t', '--template', required=True,
                       help='Path to template file')
    parser.add_argument('-s', '--slides', required=True,
                       help='Path to slides spec JSON')
    parser.add_argument('-o', '--output', required=True,
                       help='Output path for presentation')
    parser.add_argument('-c', '--config', 
                       help='Path to config JSON (optional)')
    
    args = parser.parse_args()
    
    try:
        create_from_files(
            template_path=args.template,
            slides_spec_path=args.slides,
            output_path=args.output,
            config_path=args.config
        )
        print("\n✅ Success!")
    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
