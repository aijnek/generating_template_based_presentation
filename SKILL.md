---
name: generating-template-based-presentation
description: Generating PowerPoint presentations that perfectly conform to any user-provided template. Automatically analyzes template layouts, generates configuration, and builds slides with complete template fidelity. Use when users upload a PowerPoint template and want to create presentations using it, or mention "template," "corporate template," "brand guidelines" with presentations.
license: MIT
---

# Generating Template-Based Presentations

Create professional PowerPoint presentations with perfect template fidelity. This skill analyzes any template, maps its layouts, and generates slides that exactly match the template's styling.

## When to Use This Skill

Trigger this skill when:
- User uploads a .pptx template file and wants to create presentations
- User mentions "use my template", "corporate template", "brand guidelines"
- User wants to generate multiple presentations with consistent styling
- User says "create a presentation using this template"
- User needs to automate presentation generation while maintaining design consistency

**Examples of user requests:**
- "Create a presentation using this template with 5 slides about AI"
- "Use my corporate template to make a quarterly report"
- "Generate a pitch deck that matches our brand guidelines"
- "Make 10 presentations from this template with different data"

## Core Workflow

### Phase 1: Analyze Template

**When**: User uploads a template or you need to understand layout structure

```bash
python scripts/analyze_template.py <template.pptx> --output config
```

**This produces:**
- `config.json` - Layout mapping (indices to semantic names)
- `config_analysis.json` - Full structural data
- `config.txt` - Human-readable report

**What it does:**
1. Extracts all layout names and indices
2. Identifies placeholder types (title, content, images, tables)
3. Categorizes layouts (title slides, content slides, two-column, etc.)
4. Generates recommended configuration

**Example output:**
```json
{
  "title_slide": 0,
  "content_slide": 4,
  "two_column": 7,
  "section_header": 3,
  "closing": 12
}
```

### Phase 2: Plan Content Structure

**When**: After analysis, work with user to define what slides they need

**Ask the user:**
1. How many slides?
2. What content for each slide?
3. Which layout for each slide?

**Create a slides specification** (JSON or interactive):
```json
[
  {
    "type": "title",
    "title": "Main Title",
    "subtitle": "Subtitle Text"
  },
  {
    "type": "content",
    "title": "Key Points",
    "content": ["Point 1", "Point 2", "Point 3"]
  },
  {
    "type": "two_column",
    "title": "Comparison",
    "left": ["Option A", "Feature 1"],
    "right": ["Option B", "Feature 2"]
  }
]
```

### Phase 3: Generate Presentation

**Option A: Using CLI**
```bash
python scripts/create_presentation.py \
  --template template.pptx \
  --config config.json \
  --slides slides.json \
  --output output.pptx
```

**Option B: Using Python API**
```python
from core.presentation_builder import PresentationBuilder, TemplateConfig

# Load configuration
config = TemplateConfig.from_file('config.json')

# Initialize builder with template
builder = PresentationBuilder('template.pptx', config)

# Add slides
builder.add_title_slide("My Presentation", "Subtitle")
builder.add_content_slide("Agenda", [
    "Introduction",
    "Main Content", 
    "Conclusion"
])

# Save
builder.save('output.pptx')
```

**Option C: Interactive Mode**
```bash
python scripts/create_presentation.py --interactive
```

This guides the user through each step interactively.

## Slide Types

### Title Slide
```json
{"type": "title", "title": "Main Title", "subtitle": "Optional Subtitle"}
```

```python
builder.add_title_slide("Main Title", "Subtitle")
```

### Content Slide (Bullets)
```json
{
  "type": "content",
  "title": "Slide Title", 
  "content": ["Bullet 1", "Bullet 2", "Bullet 3"]
}
```

```python
builder.add_content_slide("Slide Title", ["Bullet 1", "Bullet 2"])
```

### Two-Column Slide
```json
{
  "type": "two_column",
  "title": "Comparison",
  "left": ["Left 1", "Left 2"],
  "right": ["Right 1", "Right 2"]
}
```

```python
builder.add_two_column_slide("Comparison", 
    left_content=["Left 1"], 
    right_content=["Right 1"])
```

### Custom Layout
```json
{"type": "custom", "layout": 5}
```

```python
slide = builder.add_slide(5)  # Layout index
slide.shapes.title.text = "Custom Title"
```

## Core API Reference

### TemplateConfig

Manages layout mappings from indices to semantic names.

```python
from core.presentation_builder import TemplateConfig

# Create from dict
config = TemplateConfig({
    'title_slide': 0,
    'content_slide': 4
})

# Load from file
config = TemplateConfig.from_file('config.json')

# Save to file
config.save('config.json')

# Get layout index
index = config.get_layout_index('title_slide')  # Returns 0

# Add new mapping
config.add_layout('my_special_layout', 8)
```

### PresentationBuilder

Main class for creating presentations from templates.

```python
from core.presentation_builder import PresentationBuilder

# Initialize
builder = PresentationBuilder('template.pptx', config)

# Inspect available layouts
builder.print_layouts()                    # Print to console
layouts = builder.list_layouts()           # Get as list
info = builder.get_layout_info(4)         # Get specific layout

# Add slides
slide = builder.add_slide(0)                    # By index
slide = builder.add_slide('content_slide')      # By config key
slide = builder.add_slide('Custom Layout Name') # By layout name

# Helper methods
builder.add_title_slide(title, subtitle=None)
builder.add_content_slide(title, content=None)
builder.add_two_column_slide(title, left_content, right_content)

# Properties
builder.slide_count        # Number of slides
builder.layout_count       # Number of available layouts

# Save
builder.save('output.pptx')
```

## Common Patterns

### Pattern 1: Quick Single Presentation

```python
from core.presentation_builder import create_simple_presentation

slides_data = [
    {"type": "title", "title": "Q1 Report"},
    {"type": "content", "title": "Summary", "content": ["Item 1", "Item 2"]}
]

create_simple_presentation(
    template_path='template.pptx',
    output_path='output.pptx',
    slides_data=slides_data,
    config_path='config.json'
)
```

### Pattern 2: Batch Processing Multiple Presentations

```python
from core.presentation_builder import PresentationBuilder, TemplateConfig

config = TemplateConfig.from_file('config.json')

# Generate one presentation per data item
for item in data_items:
    builder = PresentationBuilder('template.pptx', config)
    
    builder.add_title_slide(item['title'])
    builder.add_content_slide("Details", item['details'])
    
    builder.save(f'output_{item["id"]}.pptx')
```

### Pattern 3: Data-Driven Generation

```python
import json
from core.presentation_builder import PresentationBuilder, TemplateConfig

# Load data from external source
with open('data.json') as f:
    data = json.load(f)

config = TemplateConfig.from_file('config.json')
builder = PresentationBuilder('template.pptx', config)

# Generate slides from data
for section in data['sections']:
    builder.add_content_slide(
        title=section['heading'],
        content=section['points']
    )

builder.save('data_driven_output.pptx')
```

### Pattern 4: Multiple Templates

```python
from core.presentation_builder import PresentationBuilder, TemplateConfig

templates = {
    'corporate': ('corp_template.pptx', 'corp_config.json'),
    'sales': ('sales_template.pptx', 'sales_config.json')
}

for template_type, (template_path, config_path) in templates.items():
    config = TemplateConfig.from_file(config_path)
    builder = PresentationBuilder(template_path, config)
    
    # Build slides...
    
    builder.save(f'{template_type}_output.pptx')
```

## Understanding Template Analysis

When you run `analyze_template.py`, it categorizes layouts:

**Title Slides**: Layouts with "title" and "slide" in name, usually index 0
**Content Slides**: Layouts with title + body/content placeholders  
**Two-Column Slides**: Layouts with 2+ content placeholders  
**Section Headers**: Layouts with "section" or "header" in name  
**Blank Slides**: Layouts with "blank" or "empty" in name  

The analyzer provides **recommendations**, but you should:
1. Review the human-readable report (`config.txt`)
2. Check if recommended indices make sense
3. Adjust configuration based on actual template structure
4. Test with a few slides before batch processing

## Troubleshooting

### Issue: "Template not found"
**Cause**: Path to template is incorrect  
**Fix**: Use absolute paths or verify file exists
```python
from pathlib import Path
assert Path('template.pptx').exists()
```

### Issue: "No placeholder on this slide"
**Cause**: Layout doesn't have expected placeholders  
**Fix**: 
1. Run `analyze_template.py` to see actual structure
2. Use correct layout index
3. Manually iterate through placeholders:
```python
for shape in slide.placeholders:
    if hasattr(shape, 'text_frame'):
        shape.text = "Your text"
        break
```

### Issue: Text not appearing
**Cause**: Accessing wrong placeholder or no text_frame  
**Fix**: Check placeholder structure
```python
layout_info = builder.get_layout_info(4)
print(layout_info['placeholders'])
```

### Issue: Config doesn't match template
**Cause**: Template was updated after config generation  
**Fix**: Re-run analysis
```bash
python scripts/analyze_template.py template.pptx --output config
```

### Issue: Slide looks different from template
**Cause**: This skill cannot modify layouts/themes, only use existing ones  
**Fix**: Ensure you're using the correct layout index. The skill preserves all template styling - if it looks different, you're likely using the wrong layout.

## What This Skill Can Do

✅ **Preserve template styling** - All colors, fonts, themes maintained  
✅ **Use existing layouts** - Access any layout in the template  
✅ **Fill placeholders** - Title, content, bullets automatically  
✅ **Batch generation** - Create many presentations from one template  
✅ **Configuration management** - Reusable configs per template  
✅ **Multiple templates** - Switch between templates easily  

## What This Skill Cannot Do

❌ **Modify layouts** - Cannot create new layouts or change existing ones  
❌ **Change themes** - Cannot alter colors, fonts, or design  
❌ **Add animations** - Cannot add transitions or animations  
❌ **Insert images** - Image placeholders require manual handling  
❌ **Create tables** - Table placeholders need manual population  
❌ **Modify slide masters** - Master slides are read-only  

**Philosophy**: This skill uses templates as-is. Think of it as a form-filling tool that works with any template structure.

## Advanced: Working with Placeholders

For complex scenarios, work directly with placeholders:

```python
slide = builder.add_slide(7)

# Iterate through all placeholders
for shape in slide.placeholders:
    ph_idx = shape.placeholder_format.idx
    ph_type = shape.placeholder_format.type
    
    print(f"Placeholder {ph_idx}: {ph_type}")
    
    # Check if it can hold text
    if hasattr(shape, 'text_frame'):
        shape.text = "Custom content"
    
    # Check if it's for images
    if ph_type == 18:  # PICTURE placeholder
        # Handle image insertion
        pass
```

## Best Practices

1. **Always analyze templates first** - Never assume layout structure
2. **Save configs per template** - Keep them with templates for reuse
3. **Validate content before generation** - Check slide specs match layouts
4. **Test with small batches first** - Generate 2-3 slides, verify, then scale
5. **Handle errors gracefully** - Not all templates have all layout types
6. **Document your configs** - Add comments about which index is which
7. **Version control templates and configs** - Track changes together

## Dependencies

```bash
pip install python-pptx
```

No other dependencies required.

## File Structure

```
template-presentation-generator/
├── SKILL.md                      # This file
├── README.md                     # Quick start guide
├── core/
│   └── presentation_builder.py   # Core library
├── scripts/
│   ├── analyze_template.py       # Template analysis tool
│   └── create_presentation.py    # CLI creator
├── evals/
│   ├── evals.json                # Test cases
│   └── files/                    # Test templates
└── requirements.txt              # Dependencies
```

## Examples in Detail

See README.md for:
- Complete workflow examples
- Real-world use cases (weekly reports, client presentations)
- Batch processing scripts
- Integration patterns

## Version

1.0.0 - Initial release
