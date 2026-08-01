"""
xlsxjinja - Excel Template Library with Jinja2

Generate Excel (.xlsx) files from templates using Jinja2 syntax.

Features:
- Template-based Excel generation with Jinja2
- Support for loops, conditionals, and variables
- Rich text formatting preservation
- Image insertion
- Data validation preservation
- Merged cells handling
- Auto filter support

Example:
    >>> from xlsxjinja import BookWriter
    >>> writer = BookWriter('template.xlsx')
    >>> writer.render_book([{'name': 'John', 'items': [...]}])
    >>> writer.save('output.xlsx')

Supported formats:
- .xlsx (Excel 2007+)

Python version:
- Python 3.7+
"""

from .logger import get_logger, setup_logger
from .writer import BookWriter

__version__ = "1.0.0"
__author__ = "Your Name"  # Update dengan nama Anda
__all__ = ["BookWriter", "setup_logger", "get_logger"]
