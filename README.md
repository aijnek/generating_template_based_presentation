# Generating Template-Based Presentations

Agent skill that can generate PowerPoint presentations that perfectly match any template's styling.

下記に従ってscriptを走らせることもできるが，基本的にClaude DesktopのAgent Skillとして利用することを想定している．

## Quick Start

### 1. Install Dependencies

```bash
pip install python-pptx
```

### 2. Analyze Your Template

```bash
python scripts/analyze_template.py your_template.pptx --output config
```

This creates:
- `config.json` - Layout configuration
- `config_analysis.json` - Full analysis data
- `config.txt` - Human-readable report

### 3. Create Slides Specification

Create `slides.json`:

```json
[
  {
    "type": "title",
    "title": "My Presentation",
    "subtitle": "An Amazing Deck"
  },
  {
    "type": "content",
    "title": "Key Points",
    "content": ["Point 1", "Point 2", "Point 3"]
  }
]
```

### 4. Generate Presentation

```bash
python scripts/create_presentation.py \
  --template your_template.pptx \
  --config config.json \
  --slides slides.json \
  --output output.pptx
```

## Usage Examples

### Example 1: Python API

```python
from core.presentation_builder import PresentationBuilder, TemplateConfig

# Load config
config = TemplateConfig.from_file('config.json')

# Create builder
builder = PresentationBuilder('template.pptx', config)

# Add slides
builder.add_title_slide("My Title")
builder.add_content_slide("Agenda", ["Item 1", "Item 2"])

# Save
builder.save('output.pptx')
```

### Example 2: Batch Processing

```python
from core.presentation_builder import PresentationBuilder, TemplateConfig

config = TemplateConfig.from_file('config.json')

for customer in customers:
    builder = PresentationBuilder('template.pptx', config)
    builder.add_title_slide(f"Proposal for {customer.name}")
    builder.add_content_slide("Services", customer.services)
    builder.save(f'{customer.id}_proposal.pptx')
```

### Example 3: Data-Driven

```python
import json
from core.presentation_builder import create_simple_presentation

with open('data.json') as f:
    slides_data = json.load(f)

create_simple_presentation(
    template_path='template.pptx',
    output_path='output.pptx',
    slides_data=slides_data,
    config_path='config.json'
)
```

## Slide Types

### Title Slide
```json
{"type": "title", "title": "Main Title", "subtitle": "Subtitle"}
```

### Content Slide
```json
{
  "type": "content",
  "title": "Slide Title",
  "content": ["Bullet 1", "Bullet 2"]
}
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

### Custom Layout
```json
{"type": "custom", "layout": 5}
```

## Configuration Format

The `config.json` maps semantic names to layout indices:

```json
{
  "title_slide": 0,
  "content_slide": 4,
  "two_column": 7,
  "section_header": 3,
  "closing": 12,
  "blank": 13
}
```

Use these names in your code:
```python
builder.add_slide('section_header')  # Uses layout 3
builder.add_slide('closing')         # Uses layout 12
```

## Limitations

- ✅ Preserves all template styling
- ✅ Uses any template layout
- ✅ Batch generation supported
- ❌ Cannot modify layouts or themes
- ❌ Cannot add animations
- ❌ Images/tables need manual handling

## Troubleshooting

**Problem**: "Template not found"  
**Solution**: Use absolute paths or verify file exists

**Problem**: Text not appearing  
**Solution**: Run `analyze_template.py` to check layout structure

**Problem**: Wrong layout used  
**Solution**: Check config.json matches your template's layout indices

## Documentation

See `SKILL.md` for complete API reference and advanced usage.

## License

MIT
