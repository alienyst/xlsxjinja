"""Custom Jinja2 environment with Excel-specific extensions and error handling."""

import re
import sys

from jinja2 import Environment
from jinja2.exceptions import TemplateSyntaxError

from .logger import get_logger
from .xlext import (
    ImageExtension,
    NodeExtension,
    OpExtension,
    SegmentExtension,
    XvExtension,
)
from .ynext import YnExtension

logger = get_logger(__name__)


class JinjaEnv(Environment):
    """
    Custom Jinja2 environment for Excel template rendering.

    Features:
    - Excel-specific extensions (row, cell, seg, xv, img, yn, op)
    - Enhanced error reporting with cell location information
    - Node map integration for tree navigation

    Flow:
    1. Initialize with Excel-specific Jinja2 extensions
    2. Template is compiled with node references
    3. During rendering, errors are caught and enhanced with cell locations
    4. Error messages show exact Excel cell addresses for easy debugging
    """

    def __init__(self, node_map):
        """
        Initialize Jinja2 environment with custom extensions.

        Args:
            node_map: NodeMap instance for tree navigation

        Flow:
        1. Call parent Environment constructor
        2. Register all Excel-specific extensions
        3. Store node_map reference for error reporting
        4. Set offset for coordinate calculations
        """
        super().__init__(
            extensions=[
                NodeExtension,
                SegmentExtension,
                YnExtension,
                XvExtension,
                ImageExtension,
                OpExtension,
            ]
        )
        self.node_map = node_map
        self.offset = 0  # openpyxl uses 1-based indexing

    def handle_exception(self, *args, **kwargs):
        """
        Enhanced error handling with cell location info.

        When a Jinja2 template error occurs, this method provides
        helpful debugging information including which Excel cell
        contains the error.

        Flow:
        1. Capture exception information from sys
        2. Format error type and message with colors
        3. If syntax error, extract line number and source
        4. Log error with line numbers (log_lines)
        5. Log error with cell addresses (log_cells)
        6. Call parent exception handler

        Args:
            *args: Variable arguments from Jinja2
            **kwargs: Keyword arguments including 'source'
        """
        exc_type, exc_value, tb = sys.exc_info()
        self.red_fmt = "\033[31m%s\033[0m"
        self.blue_fmt = "\033[34m%s\033[0m"
        self.error_type = self.red_fmt % (f"error type:  {exc_type}")
        self.error_message = self.red_fmt % (f"error message:  {exc_value}")

        if exc_type is TemplateSyntaxError:
            lineno = exc_value.lineno
            source = kwargs["source"]
            src_lines = source.splitlines()
            self.log_lines(lineno, src_lines)
            self.log_cells(lineno, src_lines)

        super().handle_exception(*args, **kwargs)

    def get_debug_info(self, line):
        """
        Extract debug info from template line.

        Args:
            line: Line of template source

        Returns:
            DebugInfo object or None

        Flow:
        1. Search for node key pattern in line (format: 'row,col')
        2. Extract node key from match
        3. Look up node in node_map
        4. Get debug info from node (includes cell address)
        5. Return debug info or None
        """
        p = re.compile(r"'(\d*,\d*[,\d]*)'")
        m = p.findall(line)
        debug_info = None

        if len(m) > 0:
            key = m[0]
            node = self.node_map.get_tag_node(key)
            if node:
                debug_info = node.get_debug_info(self.offset)
            else:
                logger.warning("No node found for key: %s", key)

        return debug_info

    def log_cells(self, lineno, lines):
        """
        Log error with cell addresses.

        Args:
            lineno: Line number with error
            lines: All template source lines

        Flow:
        1. Iterate through all template source lines
        2. Get debug info for each line (cell address)
        3. If error line, highlight in red
        4. If adjacent line, highlight in blue
        5. Show cell address and value for context
        """
        for i, line in enumerate(lines):
            debug_info = self.get_debug_info(line)

            if not debug_info:
                if i + 1 == lineno:
                    logger.error(self.error_message)
                log_str = self.red_fmt % (line)
                logger.error(log_str)
                continue

            if debug_info.value and isinstance(debug_info.value, str):
                line_info = f"{debug_info.address} : {debug_info.value}"

                if i + 1 == lineno:
                    log_str = self.red_fmt % (line_info)
                    logger.error(
                        self.blue_fmt % ("Syntax Error in " + debug_info.address)
                    )
                    logger.error(self.error_message)
                elif i + 1 in [lineno - 1, lineno + 1]:
                    log_str = self.blue_fmt % (line_info)
                else:
                    log_str = line_info

                logger.error(log_str)

    def log_lines(self, lineno, lines):
        """
        Log error with line numbers.

        Args:
            lineno: Line number with error
            lines: All template source lines

        Flow:
        1. Iterate through all template source lines
        2. Get debug info for each line
        3. Format with line number and cell address
        4. Highlight error line in red
        5. Highlight adjacent lines in blue for context
        """
        for i, line in enumerate(lines):
            debug_info = self.get_debug_info(line)

            if not debug_info:
                if i + 1 == lineno:
                    logger.error(self.error_message)
                line_info = f"line {i + 1:4d} : {line}"
                logger.error(self.red_fmt % (line_info))
                continue

            address_line = "   <---   " + debug_info.address
            line_info = f"line {i + 1:4d} : {line} {address_line}"

            if i + 1 == lineno:
                log_str = self.red_fmt % (line_info)
                logger.error(self.blue_fmt % ("Syntax Error in " + debug_info.address))
                logger.error(self.error_message)
            elif i + 1 in [lineno - 1, lineno + 1]:
                log_str = self.blue_fmt % (line_info)
            else:
                log_str = line_info

            logger.error(log_str)
