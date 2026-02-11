#!/usr/bin/env python3
"""
Template Analyzer

Analyzes PowerPoint templates to extract layout information and generate
configuration files.

Usage:
    python analyze_template.py <template.pptx> [--output <config>]
"""

import sys
import json
from pathlib import Path
import os

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.presentation_builder import PresentationBuilder, TemplateConfig


def analyze_template(template_path):
    """Analyze a PowerPoint template"""
    builder = PresentationBuilder(template_path)
    layouts = builder.list_layouts()
    
    # Categorize layouts
    categorized = {
        'title_slides': [],
        'content_slides': [],
        'two_column_slides': [],
        'blank_slides': [],
        'section_headers': [],
        'other': []
    }
    
    for layout in layouts:
        idx = layout['index']
        name = layout['name'].lower()
        placeholders = layout['placeholders']
        
        has_title = any('title' in ph['name'].lower() for ph in placeholders)
        content_count = sum(1 for ph in placeholders 
                          if 'content' in ph['name'].lower() or 
                             'text' in ph['name'].lower())
        
        # Categorize
        if 'title' in name and 'slide' in name and idx == 0:
            categorized['title_slides'].append(idx)
        elif 'blank' in name or 'empty' in name:
            categorized['blank_slides'].append(idx)
        elif 'section' in name or 'header' in name:
            categorized['section_headers'].append(idx)
        elif content_count >= 2 or 'two' in name or '2' in name:
            categorized['two_column_slides'].append(idx)
        elif has_title and content_count >= 1:
            categorized['content_slides'].append(idx)
        else:
            categorized['other'].append(idx)
    
    # Generate recommended config
    config = {}
    if categorized['title_slides']:
        config['title_slide'] = categorized['title_slides'][0]
    elif layouts:
        config['title_slide'] = 0
    
    if categorized['content_slides']:
        config['content_slide'] = categorized['content_slides'][0]
    elif len(layouts) > 1:
        config['content_slide'] = 1
    
    if categorized['two_column_slides']:
        config['two_column'] = categorized['two_column_slides'][0]
    
    if categorized['section_headers']:
        config['section_header'] = categorized['section_headers'][0]
    
    if categorized['blank_slides']:
        config['blank'] = categorized['blank_slides'][0]
    
    return {
        'template_name': Path(template_path).name,
        'total_layouts': len(layouts),
        'layouts': layouts,
        'categorized': categorized,
        'recommended_config': config
    }


def print_report(analysis):
    """Print human-readable report"""
    print("=" * 70)
    print(f"TEMPLATE ANALYSIS: {analysis['template_name']}")
    print("=" * 70)
    print(f"\nTotal Layouts: {analysis['total_layouts']}\n")
    
    print("LAYOUT CATEGORIES:")
    print("-" * 70)
    for category, indices in analysis['categorized'].items():
        if indices:
            print(f"\n{category.upper().replace('_', ' ')}:")
            for idx in indices:
                layout = analysis['layouts'][idx]
                print(f"  [{idx}] {layout['name']}")
    
    print("\n" + "-" * 70)
    print("\nRECOMMENDED CONFIGURATION:")
    print("-" * 70)
    print(json.dumps(analysis['recommended_config'], indent=2))
    print()


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Analyze PowerPoint template')
    parser.add_argument('template', help='Path to PowerPoint template')
    parser.add_argument('-o', '--output', help='Output base path (without extension)')
    
    args = parser.parse_args()
    
    if not Path(args.template).exists():
        print(f"Error: Template not found: {args.template}", file=sys.stderr)
        sys.exit(1)
    
    print(f"Analyzing template: {args.template}\n")
    analysis = analyze_template(args.template)
    
    if args.output:
        base = Path(args.output)
        
        # Save config
        config_path = base.with_suffix('.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(analysis['recommended_config'], f, indent=2, ensure_ascii=False)
        print(f"✓ Saved config: {config_path}")
        
        # Save full analysis
        analysis_path = base.with_name(base.stem + '_analysis.json')
        with open(analysis_path, 'w', encoding='utf-8') as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False, default=str)
        print(f"✓ Saved analysis: {analysis_path}")
        
        # Save report
        report_path = base.with_suffix('.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            old_stdout = sys.stdout
            sys.stdout = f
            print_report(analysis)
            sys.stdout = old_stdout
        print(f"✓ Saved report: {report_path}\n")
    else:
        print_report(analysis)


if __name__ == '__main__':
    main()
