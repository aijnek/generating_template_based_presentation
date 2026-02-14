# Presentation
The root object for a presentation is Presentation. 

A presentation is loaded by constructing a new Presentation instance, passing in the path to a presentation to be loaded:

```python
from pptx import Presentation

path = 'slide-deck-foo.pptx'
prs = Presentation(path)
```

# Slide Master
A slide master controls the global design of a presentation, defining default layouts, fonts, colors, and background for all slides so you can keep a consistent look and update styles in one place. One presentation can have multiple slide masters, but it is sufficient to consider the first slide master in practice.

The (first) slide master of a presentation is loaded like this:

```python
slide_master = prs.slide_master
```

# Slide Layouts
A slide layout is like a template for a slide. Whatever is on the slide layout “shows through” on a slide created with it and formatting choices made on the slide layout are inherited by the slide. 

In PowerPoint, slide layouts belong to a slide master, not directly to a presentation, so normally you have to access the slide layouts via their slide master. However, the python-pptx library makes it easy to access the layouts directly from the presentation:

```python
for i, layout in enumerate(prs.slide_layouts):
    print(f'--- Layout {i}: {layout.name} ---')
```

# Slides
The slides in a presentation belong to the presentation object and are accessed using the slides attribute:

```python
slides = prs.slides
```

## Extract Layout for Slides
The layout used for each slide can be accessed like this:   
```python
for i, slide in enumerate(prs.slides):
    layout = slide.slide_layout
    print(f'Slide {i}: layout="{layout.name}"')
```

## Adding a Slide
Every slide in a presentation is based on a slide layout. Not surprising then that you have to specify which slide layout to use when you create a new slide.

Adding a slide is accomplished by calling the `add_slide()` method on the slides attribute of the presentation. A slide layout must be passed in to specify the layout the new slide should take on:

```python
title_slide_layout = prs.slide_layouts[0]
new_slide = prs.slides.add_slide(title_slide_layout)
```

## Removing a Slide
python-pptx doesn't provide a built-in method to remove a slide. The following helper function accomplishes this by directly manipulating the underlying XML elements:

```python
def remove_slide(prs, slide_index):
    """Remove a slide from the presentation.

    Args:
        prs: Presentation object
        slide_index: 0-based index of the slide to remove
    """
    # 1. Get the target slide ID element from the presentation's slide ID list
    sldIdLst = prs._element.sldIdLst
    target_sldId = sldIdLst.sldId_lst[slide_index]

    # 2. Remove the slide ID from the list
    sldIdLst.remove(target_sldId)

    # 3. Drop the relationship (this also removes the slide part from the package)
    prs.part.drop_rel(target_sldId.rId)
```

# Shapes
Pretty much anything on a slide is a shape; the only thing I can think of that can appear on a slide that’s not a shape is a slide background. There are between six and ten different types of shape, depending how you count. I’ll explain some of the general shape concepts you’ll need to make sense of how to work with them and then we’ll jump right into working with the specific types.

Technically there are six and only six different types of shapes that can be placed on a slide:

## Auto Shape
This is a regular shape, like a rectangle, an ellipse, or a block arrow. They come in a large variety of preset shapes, in the neighborhood of 180 different ones. An auto shape can have a fill and an outline, and can contain text. Some auto shapes have adjustments, the little yellow diamonds you can drag to adjust how round the corners of a rounded rectangle are for example. A text box is also an auto shape, a rectangular one, just by default without a fill and without an outline.

## Picture
A raster image, like a photograph or clip art is referred to as a picture in PowerPoint. It's its own kind of shape with different behaviors than an auto shape. Note that an auto shape can have a picture fill, in which an image "shows through" as the background of the shape instead of a fill color or gradient. That's a different thing. But cool.

## Graphic Frame
This is the technical name for the container that holds a table, a chart, a smart art diagram, or media clip. You can’t add one of these by itself, it just shows up in the file when you add a graphical object. You probably won’t need to know anything more about these.

## Group Shape
In PowerPoint, a set of shapes can be grouped, allowing them to be selected, moved, resized, and even filled as a unit. When you group a set of shapes a group shape gets created to contain those member shapes. You can’t actually see these except by their bounding box when the group is selected.

## Line/Connector
Lines are different from auto shapes because, well, they’re linear. Some lines can be connected to other shapes and stay connected when the other shape is moved. These aren’t supported yet either so I don’t know much more about them. I’d better get to these soon though, they seem like they’d be very handy.

## Content Part
I actually have only the vaguest notion of what these are. It has something to do with embedding “foreign” XML like SVG in with the presentation. I’m pretty sure PowerPoint itself doesn’t do anything with these. My strategy is to ignore them. Working good so far.

# Placeholders
Intuitively, a placeholder is a pre-formatted container into which content can be placed. 

## A Placeholder is a Shape
Placeholders are an orthogonal category of shape, which is to say multiple shape types can be placeholders. In particular, the auto shape (p:sp element), picture (p:pic element), and graphic frame (p:graphicFrame) shape types can be a placeholder. The group shape (p:grpSp), connector (p:cxnSp), and content part (p:contentPart) shapes cannot be a placeholder. A graphic frame placeholder can contain a table, a chart, or SmartArt.

## Placeholder Types
There are 18 types of placeholder.

### Title, Center Title, Subtitle, Body
These placeholders typically appear on a conventional “word chart” containing text only, often organized as a title and a series of bullet points. All of these placeholders can accept text only.

### Content
This multi-purpose placeholder is the most commonly used for the body of a slide. When unpopulated, it displays 6 buttons to allow insertion of a table, a chart, SmartArt, a picture, clip art, or a media clip.

### Picture, Clip Art
These both allow insertion of an image. The insert button on a clip art placeholder brings up the clip art gallery rather than an image file chooser, but otherwise these behave the same.

### Chart, Table, Smart Art
These three allow the respective type of rich graphical content to be inserted.

### Media Clip
Allows a video or sound recording to be inserted.

### Date, Footer, Slide Number
These three appear on most slide masters and slide layouts, but do not behave as most users would expect. These also commonly appear on the Notes Master and Handout Master.

### Header
Only valid on the Notes Master and Handout Master.

### Vertical Body, Vertical Object, Vertical Title
Used with vertically oriented languages such as Japanese.

## Unpopulated vs. Populated
A placeholder on a slide can be empty or filled. This is most evident with a picture placeholder. When unpopulated, a placeholder displays customizable prompt text. A rich content placeholder will also display one or more content insertion buttons when empty.

A text-only placeholder enters “populated” mode when the first character of text is entered and returns to “unpopulated” mode when the last character of text is removed. A rich-content placeholder enters populated mode when content such as a picture is inserted and returns to unpopulated mode when that content is deleted. In order to delete a populated placeholder, the shape must be deleted twice. The first delete removes the content and restores the placeholder to unpopulated mode. An additional delete will remove the placeholder itself. A deleted placeholder can be restored by reapplying the layout.

## Placeholders Inherit
A placeholder appearing on a slide is only part of the overall placeholder mechanism. Placeholder behavior requires three different categories of placeholder shape; those that exist on a slide master, those on a slide layout, and those that ultimately appear on a slide in a presentation.

These three categories of placeholder participate in a property inheritance hierarchy, either as an inheritor, an inheritee, or both. Placeholder shapes on masters are inheritees only. Conversely placeholder shapes on slides are inheritors only. Placeholders on slide layouts are both, a possible inheritor from a slide master placeholder and an inheritee to placeholders on slides linked to that layout.

A layout inherits from its master differently than a slide inherits from its layout. A layout placeholder inherits from the master placeholder sharing the same type. A slide placeholder inherits from the layout placeholder having the same idx value.

In general, all formatting properties are inherited from the "parent" placeholder. This includes position and size as well as fill, line, and font. Any directly applied formatting overrides the corresponding inherited value. Directly applied formatting can be removed by reapplying the layout.



## Access a placeholder

Every placeholder is also a shape, and so can be accessed using the shapes property of a slide. However, when looking for a particular placeholder, the placeholders property can make things easier.

The most reliable way to access a known placeholder is by its idx value. The idx value of a placeholder is the integer key of the slide layout placeholder it inherits properties from. As such, it remains stable throughout the life of the slide and will be the same for any slide created using that layout.

It’s usually easy enough to take a look at the placeholders on a slide and pick out the one you want:

```python
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> for shape in slide.placeholders:
...     print('%d %s' % (shape.placeholder_format.idx, shape.name))
...
0  Title 1
1  Picture Placeholder 2
2  Text Placeholder 3
```

… then, having the known index in hand, to access it directly:

```python
>>> slide.placeholders[1]
<pptx.parts.slide.PicturePlaceholder object at 0x10d094590>
>>> slide.placeholders[2].name
'Text Placeholder 3'
```

> **Note:** Item access on the placeholders collection is like that of a dictionary rather than a list. While the key used above is an integer, the lookup is on idx values, not position in a sequence. If the provided value does not match the idx value of one of the placeholders, KeyError will be raised. idx values are not necessarily contiguous.

In general, the idx value of a placeholder from a built-in slide layout (one provided with PowerPoint) will be between 0 and 5. The title placeholder will always have idx 0 if present and any other placeholders will follow in sequence, top to bottom and left to right. A placeholder added to a slide layout by a user in PowerPoint will receive an idx value starting at 10.


## Identify and Characterize a placeholder

A placeholder behaves differently that other shapes in some ways. In particular, the value of its shape_type attribute is unconditionally MSO_SHAPE_TYPE.PLACEHOLDER regardless of what type of placeholder it is or what type of content it contains:

```python
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> for shape in slide.shapes:
...     print('%s' % shape.shape_type)
...
PLACEHOLDER (14)
PLACEHOLDER (14)
PLACEHOLDER (14)
```

To find out more, it’s necessary to inspect the contents of the placeholder’s placeholder_format attribute. All shapes have this attribute, but accessing it on a non-placeholder shape raises ValueError. The is_placeholder attribute can be used to determine whether a shape is a placeholder:

```python
>>> for shape in slide.shapes:
...     if shape.is_placeholder:
...         phf = shape.placeholder_format
...         print('%d, %s' % (phf.idx, phf.type))
...
0, TITLE (1)
1, PICTURE (18)
2, BODY (2)
```

Another way a placeholder acts differently is that it inherits its position and size from its layout placeholder. This inheritance is overridden if the position and size of a placeholder are changed.


## Insert content into a placeholder

Certain placeholder types have specialized methods for inserting content. In the current release, the picture, table, and chart placeholders have content insertion methods. Text can be inserted into title and body placeholders in the same way text is inserted into an auto shape.


### PicturePlaceholder.insert_picture()

The picture placeholder has an insert_picture() method:

```python
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> placeholder = slide.placeholders[1]  # idx key, not position
>>> placeholder.name
'Picture Placeholder 2'
>>> placeholder.placeholder_format.type
PICTURE (18)
>>> picture = placeholder.insert_picture('my-image.png')
```

> **Note:** A reference to a picture placeholder becomes invalid after its insert_picture() method is called. This is because the process of inserting a picture replaces the original p:sp XML element with a new p:pic element containing the picture. Any attempt to use the original placeholder reference after the call will raise AttributeError. The new placeholder is the return value of the insert_picture() call and may also be obtained from the placeholders collection using the same idx key.

A picture inserted in this way is stretched proportionately and cropped to fill the entire placeholder. Best results are achieved when the aspect ratio of the source image and placeholder are the same. If the picture is taller in aspect than the placeholder, its top and bottom are cropped evenly to fit. If it is wider, its left and right sides are cropped evenly. Cropping can be adjusted using the crop properties on the placeholder, such as crop_bottom.


### TablePlaceholder.insert_table()

The table placeholder has an insert_table() method. The built-in template has no layout containing a table placeholder, so this example assumes a starting presentation named having-table-placeholder.pptx having a table placeholder with idx 10 on its second slide layout:

```python
>>> prs = Presentation('having-table-placeholder.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[1])
>>> placeholder = slide.placeholders[10]  # idx key, not position
>>> placeholder.name
'Table Placeholder 1'
>>> placeholder.placeholder_format.type
TABLE (12)
>>> graphic_frame = placeholder.insert_table(rows=2, cols=2)
>>> table = graphic_frame.table
>>> len(table.rows), len(table.columns)
(2, 2)
```

A table inserted in this way has the position and width of the original placeholder. Its height is proportional to the number of rows.

Like all rich-content insertion methods, a reference to a table placeholder becomes invalid after its insert_table() method is called. This is because the process of inserting rich content replaces the original p:sp XML element with a new element, a p:graphicFrame in this case, containing the rich-content object. Any attempt to use the original placeholder reference after the call will raise AttributeError. The new placeholder is the return value of the insert_table() call and may also be obtained from the placeholders collection using the original idx key, 10 in this case.

> **Note:** The return value of the insert_table() method is a PlaceholderGraphicFrame object, which has all the properties and methods of a GraphicFrame object along with those specific to placeholders. The inserted table is contained in the graphic frame and can be obtained using its table property.


### ChartPlaceholder.insert_chart()

The chart placeholder has an insert_chart() method. The presentation template built into python-pptx has no layout containing a chart placeholder, so this example assumes a starting presentation named having-chart-placeholder.pptx having a chart placeholder with idx 10 on its second slide layout:

```python
>>> from pptx.chart.data import ChartData
>>> from pptx.enum.chart import XL_CHART_TYPE

>>> prs = Presentation('having-chart-placeholder.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[1])

>>> placeholder = slide.placeholders[10]  # idx key, not position
>>> placeholder.name
'Chart Placeholder 9'
>>> placeholder.placeholder_format.type
CHART (12)

>>> chart_data = ChartData()
>>> chart_data.categories = ['Yes', 'No']
>>> chart_data.add_series('Series 1', (42, 24))

>>> graphic_frame = placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)
>>> chart = graphic_frame.chart
>>> chart.chart_type
PIE (5)
```

A chart inserted in this way has the position and size of the original placeholder.

Note the return value from insert_chart() is a PlaceholderGraphicFrame object, not the chart itself. A PlaceholderGraphicFrame object has all the properties and methods of a GraphicFrame object along with those specific to placeholders. The inserted chart is contained in the graphic frame and can be obtained using its chart property.

Like all rich-content insertion methods, a reference to a chart placeholder becomes invalid after its insert_chart() method is called. This is because the process of inserting rich content replaces the original p:sp XML element with a new element, a p:graphicFrame in this case, containing the rich-content object. Any attempt to use the original placeholder reference after the call will raise AttributeError. The new placeholder is the return value of the insert_chart() call and may also be obtained from the placeholders collection using the original idx key, 10 in this case.


## Setting the slide title

Almost all slide layouts have a title placeholder, which any slide based on the layout inherits when the layout is applied. Accessing a slide’s title is a common operation and there’s a dedicated attribute on the shape tree for it:

```python
title_placeholder = slide.shapes.title
title_placeholder.text = 'Air-speed Velocity of Unladen Swallows'
```

# Working with text

Auto shapes and table cells can contain text. Other shapes can't. Text is always manipulated the same way, regardless of its container.

Text exists in a hierarchy of three levels:

- **Shape.text_frame**
- **TextFrame.paragraphs**
- **_Paragraph.runs**

All the text in a shape is contained in its **text frame**. A text frame has vertical alignment, margins, wrapping and auto-fit behavior, a rotation angle, some possible 3D visual features, and can be set to format its text into multiple columns. It also contains a sequence of paragraphs, which always contains at least one paragraph, even when empty.

A paragraph has line spacing, space before, space after, available bullet formatting, tabs, outline/indentation level, and horizontal alignment. A paragraph can be empty, but if it contains any text, that text is contained in one or more runs.

A run exists to provide character level formatting, including font typeface, size, and color, an optional hyperlink target URL, bold, italic, and underline styles, strikethrough, kerning, and a few capitalization styles like all caps.

Let's run through these one by one. Only features available in the current release are shown.

## Accessing the text frame

As mentioned, not all shapes have a text frame. So if you're not sure and you don't want to catch the possible exception, you'll want to check before attempting to access it:

```python
for shape in slide.shapes:
    if not shape.has_text_frame:
        continue
    text_frame = shape.text_frame
    # do things with the text frame
    ...
```

## Accessing paragraphs

A text frame always contains at least one paragraph. This causes the process of getting multiple paragraphs into a shape to be a little clunkier than one might like. Say for example you want a shape with three paragraphs:

```python
paragraph_strs = [
    'Egg, bacon, sausage and spam.',
    'Spam, bacon, sausage and spam.',
    'Spam, egg, spam, spam, bacon and spam.'
]

text_frame = shape.text_frame
text_frame.clear()  # remove any existing paragraphs, leaving one empty one

p = text_frame.paragraphs[0]
p.text = paragraph_strs[0]

for para_str in paragraph_strs[1:]:
    p = text_frame.add_paragraph()
    p.text = para_str
```

## Adding text

Only runs can actually contain text. Assigning a string to the `.text` attribute on a shape, text frame, or paragraph is a shortcut method for placing text in a run contained by those objects. The following two snippets produce the same result:

```python
shape.text = 'foobar'

# is equivalent to ...

text_frame = shape.text_frame
text_frame.clear()
p = text_frame.paragraphs[0]
run = p.add_run()
run.text = 'foobar'
```

## Applying text frame-level formatting

The following produces a shape with a single paragraph, a slightly wider bottom than top margin (these default to 0.05"), no left margin, text aligned top, and word wrapping turned off. In addition, the auto-size behavior is set to adjust the width and height of the shape to fit its text. Note that vertical alignment is set on the text frame. Horizontal alignment is set on each paragraph:

```python
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE

text_frame = shape.text_frame
text_frame.text = 'Spam, eggs, and spam'
text_frame.margin_bottom = Inches(0.08)
text_frame.margin_left = 0
text_frame.vertical_anchor = MSO_ANCHOR.TOP
text_frame.word_wrap = False
text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
```

The possible values for `TextFrame.auto_size` and `TextFrame.vertical_anchor` are specified by the enumeration `MSO_AUTO_SIZE` and `MSO_VERTICAL_ANCHOR` respectively.

## Applying paragraph formatting

The following produces a shape containing three left-aligned paragraphs, the second and third indented (like sub-bullets) under the first:

```python
from pptx.enum.text import PP_ALIGN

paragraph_strs = [
    'Egg, bacon, sausage and spam.',
    'Spam, bacon, sausage and spam.',
    'Spam, egg, spam, spam, bacon and spam.'
]

text_frame = shape.text_frame
text_frame.clear()

p = text_frame.paragraphs[0]
p.text = paragraph_strs[0]
p.alignment = PP_ALIGN.LEFT

for para_str in paragraph_strs[1:]:
    p = text_frame.add_paragraph()
    p.text = para_str
    p.alignment = PP_ALIGN.LEFT
    p.level = 1
```

## Applying character formatting

Character level formatting is applied at the run level, using the `.font` attribute. The following formats a sentence in 18pt Calibri Bold and applies the theme color Accent 1.

```python
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt

text_frame = shape.text_frame
text_frame.clear()  # not necessary for newly-created shape

p = text_frame.paragraphs[0]
run = p.add_run()
run.text = 'Spam, eggs, and spam'

font = run.font
font.name = 'Calibri'
font.size = Pt(18)
font.bold = True
font.italic = None  # cause value to be inherited from theme
font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
```

If you prefer, you can set the font color to an absolute RGB value. Note that this will not change color when the theme is changed:

```python
font.color.rgb = RGBColor(0xFF, 0x7F, 0x50)
```

A run can also be made into a hyperlink by providing a target URL:

```python
run.hyperlink.address = 'https://github.com/scanny/python-pptx'
```

# Working with Notes Slides

A slide can have notes associated with it. These are perhaps most commonly encountered in the notes pane, below the slide in PowerPoint "Normal" view where it may say "Click to add notes".

The notes added here appear each time that slide is present in the main pane. They also appear in Presenter View and in Notes Page view, both available from the menu.

Notes can contain rich text, commonly bullets, bold, varying font sizes and colors, etc. The Notes Page view has somewhat more powerful tools for editing the notes text than the note pane in Normal view.

In the API and the underlying XML, the object that contains the text is known as a Notes Slide. This is because internally, a notes slide is actually a specialized instance of a slide. It contains shapes, many of which are placeholders, and allows inserting of new shapes such as pictures (a logo perhaps) auto shapes, tables, and charts. Consequently, working with a notes slide is very much like working with a regular slide.

Each slide can have zero or one notes slide. A notes slide is created the first time it is used, generally perhaps by adding notes text to a slide. Once created, it stays, even if all the text is deleted.

## The Notes Master

A new notes slide is created using the Notes Master as a template. A presentation has no notes master when newly created in PowerPoint. One is created according to a PowerPoint-internal preset default the first time it is needed, which is generally when the first notes slide is created. It's possible one can also be created by entering the Notes Master view and almost certainly is created by editing the master found there (haven't tried it though). A presentation can have at most one notes master.

The notes master governs the look and feel of notes pages, which can be viewed on-screen but are really designed for printing out. So if you want your notes page print-outs to look different from the default, you can make a lot of customizations by editing the notes master. You access the notes master editor using View > Master > Notes Master on the menu (on my version at least). Notes slides created using python-pptx will have the look and feel of the notes master in the presentation file you opened to create the presentation.

On creation, certain placeholders (slide image, notes, slide number) are copied from the notes master onto the new notes slide (if they have not been removed from the master). These "cloned" placeholders inherit position, size, and formatting from their corresponding notes master placeholder. If the position, size, or formatting of a notes slide placeholder is changed, the changed property is no long inherited (unchanged properties, however, continue to be inherited).

## Notes Slide basics

Enough talk, let's show some code. Let's say you have a slide you're working with and you want to see if it has a notes slide yet:
```python
>>> slide.has_notes_slide
False
```

Ok, not yet. Good. Let's add some notes:
```python
>>> notes_slide = slide.notes_slide
>>> text_frame = notes_slide.notes_text_frame
>>> text_frame.text = 'foobar'
```

Alright, simple enough. Let's look at what happened here:

- `slide.notes_slide` gave us the notes slide. In this case, it first created that notes slide based on the notes master. If there was no notes master, it created that too. So a lot of things can happen behind the scenes with this call the first time you call it, but if we called it again it would just give us back the reference to the same notes slide, which it caches, once retrieved.

- `notes_slide.notes_text_frame` gave us the `TextFrame` object that contains the actual notes. The reason it's not just `notes_slide.text_frame` is that there are potentially more than one. What this is doing behind the scenes is finding the placeholder shape that contains the notes (as opposed to the slide image, header, slide number, etc.) and giving us that particular text frame.

- A text frame in a notes slide works the same as one in a regular slide. More precisely, a text frame on a shape in a notes slide works the same as in any other shape. We used the `.text` property to quickly pop some text in there.

Using the text frame, you can add an arbitrary amount of text, formatted however you want.

## Notes Slide Placeholders

What we haven't explicitly seen so far is the shapes on a slide master. It's easy to get started with that:
```python
>>> notes_placeholder = notes_slide.notes_placeholder
```

This notes placeholder is just like a body placeholder we saw a couple sections back. You can change its position, size, and many other attributes, as well as get at its text via its text frame.

You can also access the other placeholders:
```python
>>> for placeholder in notes_slide.placeholders:
...     print placeholder.placeholder_format.type
...
SLIDE_IMAGE (101)
BODY (2)
SLIDE_NUMBER (13)
```

and also the shapes (a superset of the placeholders):
```python
>>> for shape in notes_slide.shapes:
...     print shape
...
<pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e890>
<pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e750>
<pptx.shapes.placeholder.NotesSlidePlaceholder object at 0x11091e990>
```

In the common case, the notes slide contains only placeholders. However, if you added an image, for example, to the notes slide, that would show up as well. Note that if you added that image to the notes master, perhaps a logo, it would appear on the notes slide "visually", but would not appear as a shape in the notes slide shape collection. Rather, it is visually "inherited" from the notes master.