# xlsxjinja Developer Guide

Quick reference guide for developers working on xlsxjinja.

---

## Quick Start

### Development Setup

```bash
# Clone repository
git clone <repo-url>
cd xlsxjinja

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -e .
pip install -r requirements-dev.txt  # If you have dev dependencies

# Run tests
python test_logging.py
```

---

## Code Structure

### Module Map

```
xlsxjinja/
├── __init__.py          # Public API exports
├── logger.py            # ⭐ Logging system
├── writer.py            # ⭐ Main entry point (BookWriter)
├── base.py              # Base classes for sheet/book operations
├── xlnode.py            # ⭐ Node tree structure
├── jinja.py             # ⭐ Custom Jinja2 environment
├── cellcontext.py       # Cell writing context
├── merger.py            # Merged cells + image insertion handling
├── richtexthandler.py   # Rich text processing
├── xlext.py             # Jinja2 extensions (row, cell, seg, xv, img, insert_img, op)
├── ynext.py             # Yes/no checkmark extension
├── patch.py             # openpyxl patches (image dedup, geometry ns, NestedDateTime tz, WebP registration)
├── utils.py             # Utility functions
├── config.py            # Configuration
├── celltag.py           # Cell tag parsing
├── nodemap.py           # Node lookup map
├── sheetresource.py     # Sheet resource management
├── writermixin.py       # Writer mixins
├── image.py             # Image handling
├── filters.py           # Custom filters
└── misc.py              # Miscellaneous utilities
```

⭐ = Core modules - start here

---

## Common Tasks

### Adding Debug Logging

```python
from .logger import get_logger

logger = get_logger(__name__)

# Usage
logger.debug("Detailed info: %s", value)
logger.info("General info: %s", status)
logger.warning("Warning: %s", issue)
logger.error("Error occurred: %s", error)
```

**When to use each level:**
- `DEBUG`: Detailed tracing, variable dumps
- `INFO`: Normal operations, milestones
- `WARNING`: Unexpected but handled situations
- `ERROR`: Errors that affect functionality

### Adding Docstrings

**Template:**
```python
def function_name(self, arg1, arg2):
    """
    Brief one-line description.
    
    Longer description if needed. Explain the purpose
    and any important context.
    
    Args:
        arg1: Description of arg1
        arg2: Description of arg2
        
    Returns:
        Description of return value
        
    Raises:
        ExceptionType: When this exception is raised
        
    Flow:
    1. First step description
    2. Second step description
    3. Final step description
    
    Example:
        >>> result = function_name(val1, val2)
        >>> print(result)
        Expected output
    """
```

**Required for:**
- ✅ All public classes
- ✅ All public methods
- ✅ All public functions
- ✅ Module-level docstrings

**Flow section:**
- Use numbered steps
- Describe the logical sequence
- Keep it concise but complete
- Use sub-steps (a, b, c) for details

### Adding a New Cell Type

1. **Create node class in `xlnode.py`:**
```python
class MyCell(Cell):
    """
    Custom cell type for special handling.
    
    Flow:
    1. Initialize with cell data
    2. Process during enter()
    3. Write during exit()
    """
    
    def enter(self):
        """Prepare cell processing."""
        super().enter()
        # Custom initialization
        
    def exit(self):
        """Write cell to output."""
        # Custom processing
        self.write(value, data_type)
```

2. **Update `create_cell()` in `xlnode.py`:**
```python
def create_cell(...):
    # ... existing checks
    elif my_special_test(value):
        cell = MyCell(...)
    # ... rest of code
```

3. **Add tests and documentation**

### Adding a Jinja2 Extension

1. **Create extension in `xlext.py` or new file:**
```python
from jinja2 import Extension, nodes

class MyExtension(Extension):
    """
    Custom Jinja2 extension for special functionality.
    
    Usage: {% mytag arg1 arg2 %}
    
    Flow:
    1. Parse tag syntax
    2. Generate callback nodes
    3. Execute callback during rendering
    """
    tags = {'mytag'}
    
    def parse(self, parser):
        """Parse tag from template."""
        lineno = next(parser.stream).lineno
        
        # Parse arguments
        args = [parser.parse_expression()]
        
        # Create callback
        call = self.call_method('_process', args)
        return nodes.Output([call], lineno=lineno)
    
    def _process(self, arg):
        """Process tag during rendering."""
        # Your logic here
        return f"Processed: {arg}"
```

2. **Register in `jinja.py`:**
```python
class JinjaEnv(Environment):
    def __init__(self, node_map):
        super().__init__(
            extensions=[
                NodeExtension,
                MyExtension,  # Add here
                # ... other extensions
            ]
        )
```

### Testing Your Changes

```python
# test_my_feature.py
from xlsxjinja import BookWriter
from xlsxjinja.logger import setup_logger

# Enable debug logging for tests
setup_logger('xlsxjinja', debug=True)

# Test your feature
writer = BookWriter('test_template.xlsx', debug=True)
writer.render_book([{'test': 'data'}])
writer.save('test_output.xlsx')

print("✓ Test passed")
```

---

## Debugging Tips

### Enable Full Debug Output

```python
from xlsxjinja import BookWriter

writer = BookWriter('template.xlsx', debug=True)
```

### Visualize Node Tree

```python
# After building tree
tree = writer.build(sheet, index, merger)
tree.tag_tree()  # Prints tree structure
```

### Check Compiled Template

```python
# Get template string before rendering
template_str = tree.to_tag()
print(template_str)
```

### Trace Cell Processing

Add debug logs in cell node's enter/exit:
```python
def enter(self):
    logger.debug("Entering cell at %s", self.get_coordinate(0))
    super().enter()
    
def exit(self):
    logger.debug("Exiting cell, writing value: %s", self.value)
    super().exit()
```

### Common Issues

**Issue: Import errors**
```bash
# Solution: Install in development mode
pip install -e .
```

**Issue: openpyxl not found**
```bash
# Solution: Install dependencies
pip install openpyxl jinja2
```

**Issue: No debug output**
```python
# Solution: Ensure debug=True
writer = BookWriter('template.xlsx', debug=True)
```

**Issue: Colors not showing**
- Windows: Install colorama
- Or: Use logger without colors

---

## Code Style Guidelines

### Python Style

- Follow PEP 8
- Use 4 spaces for indentation
- Max line length: 88 (Black formatter compatible)
- Use type hints where beneficial

### Docstring Style

- Follow Google style (with Flow section)
- One-line summary for simple functions
- Full docstring for complex functions
- Always include Flow for non-trivial logic

### Logging Style

```python
# Good
logger.debug("Processing cell at row=%d col=%d", row, col)

# Bad
logger.debug(f"Processing cell at row={row} col={col}")
# (f-strings evaluated even if not logged)

# Good
logger.error("Failed to process %s: %s", item, error)

# Bad
print(f"Error: {error}")  # Never use print
```

### Import Order

```python
# 1. Standard library
import copy
import sys

# 2. Third-party
from jinja2 import Environment
from openpyxl import load_workbook

# 3. Local
from .logger import get_logger
from .utils import tag_test
```

---

## Git Workflow

### Commit Messages

```
<type>: <subject>

<body>

<footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation only
- `refactor`: Code refactoring
- `test`: Adding tests
- `chore`: Maintenance

**Example:**
```
feat: add debug logging system

- Implement ColoredFormatter for console output
- Add setup_logger() and get_logger() functions
- Replace all print statements with logger calls
- Add debug mode to BookWriter

Closes #123
```

### Branch Naming

- `feature/logging-system`
- `fix/cell-merge-error`
- `docs/architecture-guide`
- `refactor/node-structure`

---

## Performance Tips

### 1. Font Caching
Fonts are already cached in `BookBase.font_map`. No need to optimize.

### 2. Cell Creation
Use lazy creation via `CellContext.target_cell` property.

### 3. Node Lookup
`NodeMap` provides O(1) lookup. Use it for node access.

### 4. Logging Performance
```python
# Good - lazy evaluation
logger.debug("Data: %s", expensive_operation())

# Better - conditional
if logger.isEnabledFor(logging.DEBUG):
    logger.debug("Data: %s", expensive_operation())
```

---

## Resources

### Documentation
- [ARCHITECTURE.md](ARCHITECTURE.md) - System architecture
- [README.md](README.md) - User guide
- [CHANGELOG.md](CHANGELOG.md) - Version history

### External Resources
- [openpyxl docs](https://openpyxl.readthedocs.io/)
- [Jinja2 docs](https://jinja.palletsprojects.com/)
- [Python logging](https://docs.python.org/3/library/logging.html)

### IDE Setup

**VSCode:**
```json
{
  "python.linting.enabled": true,
  "python.linting.pylintEnabled": true,
  "python.formatting.provider": "black",
  "editor.rulers": [88]
}
```

**PyCharm:**
- Enable docstring format: Google
- Set line length: 88
- Enable PEP 8 inspections

---

## FAQ

**Q: How do I add a new Jinja2 tag?**
A: Create an Extension class and register in JinjaEnv. See "Adding a Jinja2 Extension" above.

**Q: Where should I add logging?**
A: Add logger in any module using `get_logger(__name__)`. Use appropriate log levels.

**Q: How do I test template errors?**
A: Enable debug mode and check error output includes cell locations.

**Q: Can I use print() for debugging?**
A: No, always use logger. Print statements will be removed in code review.

**Q: How do I handle rich text?**
A: Use RichTextHandler. It automatically fixes split Jinja2 tags.

**Q: What's the difference between Cell and TagCell?**
A: Cell = static value, TagCell = contains Jinja2 templates.

---

## Contact & Support

- Issues: GitHub Issues
- Discussions: GitHub Discussions
- Email: [maintainer email]

---

**Happy Coding! 🚀**
