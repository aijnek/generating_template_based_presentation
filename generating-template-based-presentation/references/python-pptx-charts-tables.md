# Working with charts

python-pptx supports adding charts and modifying existing ones. Most chart types other than 3D types are supported.

## Adding a chart

The following code adds a single-series column chart in a new presentation:

```python
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

# create presentation with 1 slide ------
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

# define chart data ---------------------
chart_data = CategoryChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

# add chart to slide --------------------
x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)

prs.save('chart-01.pptx')
```

## Customizing things a bit

The remaining code will leave out code we've already seen and only show imports, for example, when they're used for the first time, just to keep the focus on the new bits.

Let's create a multi-series chart to use for these examples:

```python
chart_data = ChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Q1 Sales', (19.2, 21.4, 16.7))
chart_data.add_series('Q2 Sales', (22.3, 28.6, 15.2))
chart_data.add_series('Q3 Sales', (20.4, 26.3, 14.2))

graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)

chart = graphic_frame.chart
```

Notice that we captured the shape reference returned by the `add_chart()` call as `graphic_frame` and then extracted the chart object from the graphic frame using its `chart` property. We'll need the chart reference to get to the properties we'll need in the next steps.

The `add_chart()` method doesn't directly return the chart object. That's because a chart is not itself a shape. Rather it's a graphical (DrawingML) object contained in the graphic frame shape. Tables work this way too, also being contained in a graphic frame shape.

## XY and Bubble charts

The charts so far use a discrete set of values for the independent variable (the X axis, roughly speaking). These are perfect when your values fall into a well-defined set of categories. However, there are many cases, particularly in science and engineering, where the independent variable is a continuous value, such as temperature or frequency. These are supported in PowerPoint by XY (aka. scatter) charts.

A bubble chart is essentially an XY chart where the marker size is used to reflect an additional value, effectively adding a third dimension to the chart.

Because the independent variable is continuous, in general, the series do not all share the same X values. This requires a somewhat different data structure and that is provided for by distinct `XyChartData` and `BubbleChartData` objects used to specify the data behind charts of these types:

```python
chart_data = XyChartData()

series_1 = chart_data.add_series('Model 1')
series_1.add_data_point(0.7, 2.7)
series_1.add_data_point(1.8, 3.2)
series_1.add_data_point(2.6, 0.8)

series_2 = chart_data.add_series('Model 2')
series_2.add_data_point(1.3, 3.7)
series_2.add_data_point(2.7, 2.3)
series_2.add_data_point(1.6, 1.8)

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
).chart
```

Creation of a bubble chart is very similar, having an additional value for each data point that specifies the bubble size:

```python
chart_data = BubbleChartData()

series_1 = chart_data.add_series('Series 1')
series_1.add_data_point(0.7, 2.7, 10)
series_1.add_data_point(1.8, 3.2, 4)
series_1.add_data_point(2.6, 0.8, 8)

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.BUBBLE, x, y, cx, cy, chart_data
).chart
```

## Axes

Let's change up the category and value axes a bit:

```python
from pptx.enum.chart import XL_TICK_MARK
from pptx.util import Pt

category_axis = chart.category_axis
category_axis.has_major_gridlines = True
category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
category_axis.tick_labels.font.italic = True
category_axis.tick_labels.font.size = Pt(24)

value_axis = chart.value_axis
value_axis.maximum_scale = 50.0
value_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
value_axis.has_minor_gridlines = True

tick_labels = value_axis.tick_labels
tick_labels.number_format = '0"%"'
tick_labels.font.bold = True
tick_labels.font.size = Pt(14)
```

Okay, that was probably going a bit too far. But it gives us an idea of the kinds of things we can do with the value and category axes. Let's undo this part and go back to the version we had before.

## Data Labels

Let's add some data labels so we can see exactly what the value for each bar is:

```python
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LABEL_POSITION

plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels

data_labels.font.size = Pt(13)
data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
data_labels.position = XL_LABEL_POSITION.INSIDE_END
```

Here we needed to access a Plot object to gain access to the data labels. A plot is like a sub-chart, containing one or more series and drawn as a particular chart type, like column or line. This distinction is needed for charts that combine more than one type, like a line chart appearing on top of a column chart. A chart like this would have two plot objects, one for the series appearing as columns and the other for the lines.

Most charts only have a single plot and python-pptx doesn't yet support creating multi-plot charts, but you can access multiple plots on a chart that already has them.

In the Microsoft API, the name ChartGroup is used for this object. I found that term confusing for a long time while I was learning about MS Office charts so I chose the name Plot for that object in python-pptx.

## Legend

A legend is often useful to have on a chart, to give a name to each series and help a reader tell which one is which:

```python
from pptx.enum.chart import XL_LEGEND_POSITION

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.RIGHT
chart.legend.include_in_layout = False
```

Nice! Okay, let's try some other chart types.

## Line Chart

A line chart is added pretty much the same way as a bar or column chart, the main difference being the chart type provided in the `add_chart()` call:

```python
chart_data = ChartData()
chart_data.categories = ['Q1 Sales', 'Q2 Sales', 'Q3 Sales']
chart_data.add_series('West', (32.2, 28.4, 34.7))
chart_data.add_series('East', (24.3, 30.6, 20.2))
chart_data.add_series('Midwest', (20.4, 18.3, 26.2))

x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data
).chart

chart.has_legend = True
chart.legend.include_in_layout = False
chart.series[0].smooth = True
```

I switched the categories and series data here to better suit a line chart. You can see the line for the "West" region is smoothed into a curve while the other two have their points connected with straight line segments.

## Pie Chart

A pie chart is a little special in that it only ever has a single series and doesn't have any axes:

```python
chart_data = ChartData()
chart_data.categories = ['West', 'East', 'North', 'South', 'Other']
chart_data.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data
).chart

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False

chart.plots[0].has_data_labels = True
data_labels = chart.plots[0].data_labels
data_labels.number_format = '0%'
data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
```

# Working with tables

PowerPoint allows text and numbers to be presented in tabular form (aligned rows and columns) in a reasonably flexible way. A PowerPoint table is not nearly as functional as an Excel spreadsheet, and is definitely less powerful than a table in Microsoft Word, but it serves well for most presentation purposes.

## Concepts

There are a few terms worth reviewing as a basis for understanding PowerPoint tables:

**table** - A table is a matrix of cells arranged in aligned rows and columns. This orderly arrangement allows a reader to more easily make sense of relatively large number of individual items. It is commonly used for displaying numbers, but can also be used for blocks of text.

**cell** - An individual content "container" within a table. A cell has a text-frame in which it holds that content. A PowerPoint table cell can only contain text. It cannot hold images, other shapes, or other tables. A cell has a background fill, borders, margins, and several other formatting settings that can be customized on a cell-by-cell basis.

**row** - A side-by-side sequence of cells running across the table, all sharing the same top and bottom boundary.

**column** - A vertical sequence of cells spanning the height of the table, all sharing the same left and right boundary.

**table grid, also cell grid** - The underlying cells in a PowerPoint table are strictly regular. In a three-by-three table there are nine grid cells, three in each row and three in each column. The presence of merged cells can obscure portions of the cell grid, but not change the number of cells in the grid. Access to a table cell in python-pptx is always via that cell's coordinates in the cell grid, which may not conform to its visual location (or lack thereof) in the table.

**merged cell** - A cell can be "merged" with adjacent cells, horizontally, vertically, or both, causing the resulting cell to look and behave like a single cell that spans the area formerly occupied by those individual cells.

**merge-origin cell** - The top-left grid-cell in a merged cell has certain special behaviors. The content of that cell is what appears on the slide; content of any "spanned" cells is hidden. In python-pptx a merge-origin cell can be identified with the `_Cell.is_merge_origin` property. Such a cell can report the size of the merged cell with its `span_height` and `span_width` properties, and can be "unmerged" back to its underlying grid cells using its `split()` method.

**spanned-cell** - A grid-cell other than the merge-origin cell that is "occupied" by a merged cell is called a spanned cell. Intuitively, the merge-origin cell "spans" the other grid cells within its area. A spanned cell can be identified with its `_Cell.is_spanned` property. A merge-origin cell is not itself a spanned cell.

## Adding a table

The following code adds a 3-by-3 table in a new presentation:

```python
>>> from pptx import Presentation
>>> from pptx.util import Inches

>>> # ---create presentation with 1 slide---
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[5])

>>> # ---add table to slide---
>>> x, y, cx, cy = Inches(2), Inches(2), Inches(4), Inches(1.5)
>>> shape = slide.shapes.add_table(3, 3, x, y, cx, cy)
>>> shape
<pptx.shapes.graphfrm.GraphicFrame object at 0x1022816d0>
>>> shape.has_table
True
>>> table = shape.table
>>> table
<pptx.table.Table object at 0x1096f8d90>
```

A couple things to note:

`SlideShapes.add_table()` returns a shape that contains the table, not the table itself. In PowerPoint, a table is contained in a graphic-frame shape, as is a chart or SmartArt. You can determine whether a shape contains a table using its `has_table` property and you access the table object using the shape's `table` property.

## Inserting a table into a table placeholder

A placeholder allows you to specify the position and size of a shape as part of the presentation "template", and to place a shape of your choosing into that placeholder when authoring a presentation based on that template. This can lead to a better looking presentation, with objects appearing in a consistent location from slide-to-slide.

Placeholders come in different types, one of which is a table placeholder. A table placeholder behaves like other placeholders except it can only accept insertion of a table. Other placeholder types accept text bullets or charts.

There is a subtle distinction between a layout placeholder and a slide placeholder. A layout placeholder appears in a slide layout, and defines the position and size of the placeholder "cloned" from it onto each slide created with that layout. As long as you don't adjust the position or size of the slide placeholder, it will inherit its position and size from the layout placeholder it derives from.

To insert a table into a table placeholder, you need a slide layout that includes a table placeholder, and you need to create a slide using that layout. These examples assume that the third slide layout in template.pptx includes a table placeholder:

```python
>>> prs = Presentation('template.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[2])
```

**Accessing the table placeholder.** Generally, the easiest way to access a placeholder shape is to know its position in the slide.shapes collection. If you always use the same template, it will always show up in the same position:

```python
>>> table_placeholder = slide.shapes[1]
```

**Inserting a table.** A table is inserted into the placeholder by calling its `insert_table()` method and providing the desired number of rows and columns:

```python
>>> shape = table_placeholder.insert_table(rows=3, cols=4)
```

The return value is a GraphicFrame shape containing the new table, not the table object itself. Use the `table` property of that shape to access the table object:

```python
>>> table = shape.table
```

The containing shape controls the position and size. Everything else, like accessing cells and their contents, is done from the table object.

## Accessing a cell

All content in a table is in a cell, so getting a reference to one of those is a good place to start:

```python
>>> cell = table.cell(0, 0)
>>> cell.text
''
>>> cell.text = 'Unladen Swallow'
```

The cell is specified by its row, column coordinates as zero-based offsets. The top-left cell is at row, column (0, 0).

Like an auto-shape, a cell has a text-frame and can contain arbitrary text divided into paragraphs and runs. Any desired character formatting can be applied individually to each run. Often however, cell text is just a simple string. For these cases the read/write `_Cell.text` property can be the quickest way to set cell contents.

## Merging cells

A merged cell is produced by specifying two diagonal cells. The merged cell will occupy all the grid cells in the rectangular region specified by that diagonal:

```python
>>> cell = table.cell(0, 0)
>>> other_cell = table.cell(1, 1)
>>> cell.is_merge_origin
False
>>> cell.merge(other_cell)
>>> cell.is_merge_origin
True
>>> cell.is_spanned
False
>>> other_cell.is_spanned
True
>>> table.cell(0, 1).is_spanned
True
```

A few things to observe:

- The merged cell appears as a single cell occupying the space formerly occupied by the other grid cells in the specified rectangular region.
- The formatting of the merged cell (background color, font etc.) is taken from the merge origin cell, the top-left cell of the table in this case.
- Content from the merged cells was migrated to the merge-origin cell. That content is no longer present in the spanned grid cells (although you can't see those at the moment).
- The content of each cell appears as a separate paragraph in the merged cell; it isn't concatenated into a single paragraph.
- Content is migrated in left-to-right, top-to-bottom order of the original cells.
- Calling `other_cell.merge(cell)` would have the exact same effect. The merge origin is always the top-left cell in the specified rectangular region. There are four distinct ways to specify a given rectangular region (two diagonals, each having two orderings).

## Un-merging a cell

A merged cell can be restored to its underlying grid cells by calling the `split()` method on its merge-origin cell. Calling `split()` on a cell that is not a merge-origin raises ValueError:

```python
>>> cell = table.cell(0, 0)
>>> cell.is_merge_origin
True
>>> cell.split()
>>> cell.is_merge_origin
False
>>> table.cell(0, 1).is_spanned
False
```

Note that the content migration performed as part of the `.merge()` operation was not reversed.

## A few snippets that might be handy

**Use Case: Interrogate table for merged cells:**

```python
def iter_merge_origins(table):
    """Generate each merge-origin cell in *table*.

    Cell objects are ordered by their position in the table,
    left-to-right, top-to-bottom.
    """
    return (cell for cell in table.iter_cells() if cell.is_merge_origin)

def merged_cell_report(cell):
    """Return str summarizing position and size of merged *cell*."""
    return (
        'merged cell at row %d, col %d, %d cells high and %d cells wide'
        % (cell.row_idx, cell.col_idx, cell.span_height, cell.span_width)
    )

# ---Print a summary line for each merged cell in *table*.---
for merge_origin_cell in iter_merge_origins(table):
    print(merged_cell_report(merge_origin_cell))
```

prints a report like:

```
merged cell at row 0, col 0, 2 cells high and 2 cells wide
merged cell at row 3, col 2, 1 cells high and 2 cells wide
merged cell at row 4, col 0, 2 cells high and 1 cells wide
```

**Use Case: Access only cells that display text (are not spanned):**

```python
def iter_visible_cells(table):
    return (cell for cell in table.iter_cells() if not cell.is_spanned)
```

**Use Case: Determine whether table contains merged cells:**

```python
def has_merged_cells(table):
    for cell in table.iter_cells():
        if cell.is_merge_origin:
            return True
    return False
```
