# -*- coding: utf-8 -*-
"""
Logging configuration for xlsxjinja.

Provides debug logging capabilities that can be enabled/disabled.
Logs are only output when debug mode is active.
"""

import logging
import sys


class ColoredFormatter(logging.Formatter):
    """
    Custom formatter with color support for terminal output.

    Flow:
    1. Receives log record from logger
    2. Applies ANSI color codes based on log level
    3. Returns formatted colored message
    """

    # ANSI color codes
    COLORS = {
        "DEBUG": "\033[36m",  # Cyan
        "INFO": "\033[32m",  # Green
        "WARNING": "\033[33m",  # Yellow
        "ERROR": "\033[31m",  # Red
        "CRITICAL": "\033[35m",  # Magenta
        "RESET": "\033[0m",  # Reset
        "BLUE": "\033[34m",  # Blue
    }

    def format(self, record):
        """
        Format log record with colors.

        Args:
            record: LogRecord instance

        Returns:
            Formatted and colored log message

        Flow:
        1. Get color for log level
        2. Format base message
        3. Apply color codes
        4. Return colored message
        """
        log_color = self.COLORS.get(record.levelname, self.COLORS["RESET"])
        record.levelname = f"{log_color}{record.levelname}{self.COLORS['RESET']}"
        record.msg = f"{log_color}{record.msg}{self.COLORS['RESET']}"
        return super().format(record)


def setup_logger(name="xlsxjinja", debug=False):
    """
    Setup and configure logger for xlsxjinja.

    Args:
        name: Logger name
        debug: Enable debug mode (default: False)

    Returns:
        Configured logger instance

    Flow:
    1. Create logger with specified name
    2. Set logging level based on debug flag
    3. Create console handler with colored formatter
    4. Attach handler to logger
    5. Return configured logger
    """
    logger = logging.getLogger(name)

    # Clear existing handlers to avoid duplicates
    logger.handlers.clear()

    # Set level based on debug mode
    if debug:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.WARNING)  # Only show warnings and errors by default

    # Create console handler
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.DEBUG)

    # Create colored formatter
    formatter = ColoredFormatter("%(levelname)s - %(message)s")
    handler.setFormatter(formatter)

    # Add handler to logger
    logger.addHandler(handler)

    return logger


def get_logger(name="xlsxjinja"):
    """
    Get existing logger instance.

    Args:
        name: Logger name

    Returns:
        Logger instance

    Flow:
    1. Retrieve logger by name
    2. Return logger instance
    """
    return logging.getLogger(name)


# Module-level logger instance (initialized when first imported)
_default_logger = None


def get_default_logger():
    """
    Get or create default logger instance.

    Returns:
        Default logger instance

    Flow:
    1. Check if default logger exists
    2. Create if not exists
    3. Return logger instance
    """
    global _default_logger
    if _default_logger is None:
        _default_logger = setup_logger()
    return _default_logger
