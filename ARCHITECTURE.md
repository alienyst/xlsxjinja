# xlsxjinja Architecture & Flow Documentation

This document provides detailed information about the internal architecture, data flow, and design decisions in xlsxjinja.

---

## Table of Contents

1. [Overview](#overview)
2. [Core Concepts](#core-concepts)
3. [Data Flow](#data-flow)
4. [Module Architecture](#module-architecture)
5. [Node Tree Structure](#node-tree-structure)
6. [Rendering Pipeline](#rendering-pipeline)
7. [Debug & Logging](#debug--logging)

---

## Overview

xlsxjinja is a template rendering engine that bridges Excel and Jinja2. It works by:

1. **Reading** Excel template files using openpyxl
2. **Parsing** cells to detect Jinja2 syntax
3. **Building** a hierarchical node tree representing the template structure
4. **Compiling** the tree into a Jinja2 template string
5. **Rendering** the template with user data
6. **Writing** results back to Excel with formatting preserved

---

## Core Concepts

### 1. Node Tree

The entire Excel template is represented as a tree of nodes:

```
Tree (Sheet)
├── Row (Row 1)
│   ├── Cell (A1)
│   ├── TagCell (B1) - contains Jinja2
│   └── Cell (C1)
├── Row (Row 2)
│   ├── Cell (A2)
│   └── RichTagCell (B2) - rich text with Jinja2
└── Row (Row 3)
    └── ...
```

Each node has a lifecycle: `enter()` → process children → `exit()`

### 2. Cell Types

- **Cell**: Regular cell with static value
- **TagCell**: Cell containing Jinja2 templates (plain text)
- **RichTagCell**: Cell with rich text (multiple fonts) containing Jinja2
- **XvCell**: Cell with expression variables (`{{xv ...}}`)
- **EmptyCell**: Placeholder for cells without content

### 3. Template Compilation

Cells with Jinja2 syntax are converted to node tags:

```
Excel: {{ item.name }}
↓
Node Tag: {% cell '1,2' %}{{ item.name }}{% endcell %}
↓
Jinja2 renders this and calls back to write the cell
```

---

## Data Flow

### High-Level Flow

```
┌─────────────────┐
│ Excel Template  │
│  (template.xlsx)│
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  BookWriter     │
│  .load()        │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Parse Sheets   │
│  Build Trees    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Compile to      │
│ Jinja2 Template │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Render with     │
│ User Data       │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Write to Excel  │
│  (output.xlsx)  │
└─────────────────┘
```

### Detailed Rendering Flow

1. **Initialization**
   ```python
   writer = BookWriter('template.xlsx', debug=True)
   ```
   - Load workbook with openpyxl
   - Create NodeMap for tracking nodes
   - Create JinjaEnv with custom extensions
   - Parse all sheets into SheetResource objects

2. **Tree Building** (for each sheet)
   ```python
   tree = writer.build(sheet, index, merger)
   ```
   - Iterate through all rows and columns
   - For each cell:
     - Check for Jinja2 syntax
     - Create appropriate cell node type
     - Parse cell comments for special tags
     - Add node to tree

3. **Template Compilation**
   ```python
   template_str = tree.to_tag()
   ```
   - Each node converts itself to Jinja2 tag
   - Tags include node keys for callback: `{% cell '1,2' %}`
   - Result is a single Jinja2 template string

4. **Rendering**
   ```python
   writer.render_book(payloads)
   ```
   - For each payload (data dict):
     - Create SheetWriter for output
     - Compile Jinja2 template
     - Render template with data
     - During rendering, Jinja2 calls back via extensions:
       - `{% cell %}` → write cell
       - `{% row %}` → start new row
       - `{% seg %}` → process text segment
       - etc.

5. **Writing**
   ```python
   writer.save('output.xlsx')
   ```
   - Clean up defined names
   - Save workbook
   - Clear caches

---

## Module Architecture

### Core Modules

#### `writer.py` - Entry Point
- **BookWriter**: Main interface class
- **SheetWriter**: Handles individual sheet writing
- Flow: Load → Build → Render → Save

#### `xlnode.py` - Node System
- **Node**: Base class for tree nodes
- **Cell, TagCell, RichTagCell**: Cell node types
- **Row**: Row container node
- **Tree**: Root sheet node
- **Segment**: Text fragment within cells
- Lifecycle: `enter()` → `exit()` pattern

#### `jinja.py` - Template Engine
- **JinjaEnv**: Custom Jinja2 environment
- Features:
  - Excel-specific extensions
  - Enhanced error reporting with cell addresses
  - Color-coded error output
  - Node map integration

#### `base.py` - Base Classes
- **SheetBase**: Sheet operations (copy settings, dimensions, cells)
- **BookBase**: Workbook operations (font caching)

#### `merger.py` - Merged Cells
- **MergeMixin**: Base for merge handling
- **ImageMerger**: Track image locations
- **Merger**: Combine multiple mergers

#### `richtexthandler.py` - Rich Text
- **RichTextHandler**: Handle multi-font text
- Detects and fixes Jinja2 tags split across font segments

### Extension Modules

#### `xlext.py` - Jinja2 Extensions
- **NodeExtension**: `{% row %}`, `{% cell %}`
- **SegmentExtension**: `{% seg %}`
- **XvExtension**: `{% xv %}`
- **ImageExtension**: `{% img %}`
- **OpExtension**: `{% op %}`

#### `ynext.py` - Yes/No Extension
- **YnExtension**: `{% yn %}` for checkmarks
- Uses Wingdings 2 font

### Utility Modules

#### `logger.py` - Logging System
- **ColoredFormatter**: ANSI color support
- **setup_logger()**: Configure logger with debug mode
- **get_logger()**: Retrieve logger instance

#### `utils.py` - Utilities
- Regex patterns for tag detection
- Tag parsing and fixing
- Cell coordinate conversion

#### `patch.py` - openpyxl Patches
- Fix image writing issues
- Fix geometry guide list namespace

---

## Node Tree Structure

### Tree Hierarchy

```
Tree (index=0, node_key='0')
├── Row (rowx=1, node_key='0,0')
│   ├── Cell (A1, node_key='0,1')
│   ├── TagCell (B1, node_key='0,2')
│   │   ├── Section
│   │   │   ├── Segment ("text1")
│   │   │   ├── BlockSegment ("{% for %}")
│   │   │   └── Segment ("text2")
│   └── Cell (C1, node_key='0,3')
├── Row (rowx=2, node_key='0,1')
│   └── ...
```

### Node Keys

Each node has a unique key based on its position in the tree:
- Tree: `"0"` (sheet index)
- First child: `"0,0"`
- Second child: `"0,1"`
- First grandchild: `"0,0,0"`

These keys are embedded in Jinja2 tags for callback during rendering.

### Node Lifecycle

```python
# 1. enter() - Initialize node state
node.enter()

# 2. Process children
for child in node._children:
    child.enter()
    # ... recursive processing
    child.exit()

# 3. exit() - Finalize and write
node.exit()
```

Example for TagCell:
```python
def enter(self):
    self.child_rvs = []  # Store child results
    self.richs = []      # Track rich text positions

def exit(self):
    rv = self.pack()     # Combine child results
    self.write(rv, self.cty)  # Write to Excel
```

---

## Rendering Pipeline

### Phase 1: Template Parsing

```
Excel Cell: "Hello {{ name }}, total: {{ total }}"
↓
TagCell created with value detection
↓
Unpacked into Segments:
- Segment("Hello ")
- BlockSegment("{{ name }}")
- Segment(", total: ")
- BlockSegment("{{ total }}")
```

### Phase 2: Tag Generation

```
TagCell.to_tag() generates:
{% cell '0,2' %}Hello {{ name }}, total: {{ total }}{% endcell %}
```

### Phase 3: Jinja2 Rendering

```
Template: {% cell '0,2' %}Hello {{ name }}, total: {{ total }}{% endcell %}
Data: {'name': 'John', 'total': 100}
↓
Jinja2 processes:
- Enters cell extension
- Renders: "Hello John, total: 100"
- Calls write_cell() via callback
```

### Phase 4: Excel Writing

```
write_cell() called with:
- cell_node: TagCell instance
- rv: "Hello John, total: 100"
- cty: data type
↓
CellContext created:
- Gets target position from Box
- Creates cell at position
- Sets value and formatting
- Applies filters if any
↓
Cell written to output Excel
```

---

## Debug & Logging

### Logger Hierarchy

```
xlsxjinja (root logger)
├── xlsxjinja.writer
├── xlsxjinja.jinja
├── xlsxjinja.patch
└── xlsxjinja.xlnode
```

### Debug Mode Features

1. **Colored Output**
   - DEBUG: Cyan
   - INFO: Green
   - WARNING: Yellow
   - ERROR: Red

2. **Cell Address Tracking**
   ```
   ERROR - Syntax Error in Cell B3
   ERROR - error message: unexpected 'end of template'
   ```

3. **Context Lines**
   Shows lines before/after error for context

4. **Tree Visualization**
   ```python
   tree.tag_tree()  # Prints tree structure with indentation
   ```

### Logging Usage

```python
from xlsxjinja.logger import get_logger

logger = get_logger(__name__)

logger.debug("Processing cell at %s", address)
logger.warning("Image %s does not have key attribute", img)
logger.error("Template syntax error at line %d", lineno)
```

---

## Performance Considerations

### 1. Font Caching
- Fonts are cached in `BookBase.font_map`
- Avoids recreating InlineFont objects

### 2. Row/Column Dimension Tracking
- Uses sets to track written rows/columns
- Prevents duplicate dimension copying

### 3. Lazy Cell Creation
- CellContext creates target cells only when accessed
- Via `@property target_cell`

### 4. Node Map
- Efficient lookup of nodes by key
- Used during error reporting

---

## Extension Points

### Custom Jinja2 Extensions

Add custom tags by creating an extension:

```python
from jinja2 import Extension

class MyExtension(Extension):
    tags = {'mytag'}
    
    def parse(self, parser):
        # Parse tag syntax
        ...
        return nodes.CallBlock(...)
```

Register in `JinjaEnv.__init__()`:
```python
super().__init__(
    extensions=[
        NodeExtension,
        MyExtension,  # Add here
        ...
    ]
)
```

### Custom Cell Filters

Add filters via cell tags:

```python
def my_filter(cell_context, arg1, arg2):
    cell = cell_context.target_cell
    # Modify cell...
    
# Register in cell node:
cell_node.add_filter(my_filter, (arg1, arg2))
```

---

## Testing

### Unit Testing

Test individual components:
```python
# Test logger
from xlsxjinja.logger import setup_logger
logger = setup_logger('test', debug=True)

# Test node creation
from xlsxjinja.xlnode import Cell
cell = Cell(sheet_cell, 1, 1, "value", "s")
```

### Integration Testing

Test full pipeline:
```python
from xlsxjinja import BookWriter

writer = BookWriter('template.xlsx', debug=True)
writer.render_book([{'data': 'value'}])
writer.save('output.xlsx')
```

---

## Error Handling

### Common Error Scenarios

1. **Template Syntax Error**
   - Detected during Jinja2 compilation
   - Shows exact cell location
   - Shows context lines

2. **Missing Variables**
   - Jinja2 throws UndefinedError
   - Can use `{{ var|default('') }}`

3. **Type Errors**
   - E.g., iterating over non-iterable
   - Check data types match template expectations

4. **Rich Text Tag Splits**
   - Automatically fixed by RichTextHandler
   - Merges segments to reconstruct complete tags

---

## Contributing

When adding features:

1. **Add docstrings** with Flow documentation
2. **Use logger** instead of print statements
3. **Add tests** for new functionality
4. **Update this document** for architectural changes

---

## References

- [openpyxl documentation](https://openpyxl.readthedocs.io/)
- [Jinja2 documentation](https://jinja.palletsprojects.com/)
- [Original xltpl project](https://github.com/zhangyu836/xltpl)
