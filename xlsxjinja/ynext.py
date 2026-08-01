"""Extension for yes/no checkbox rendering in Excel cells."""

import copy

from jinja2 import nodes
from jinja2.ext import Extension
from openpyxl.cell.rich_text import TextBlock


def yes(font):
    """
    Create a checkmark using Wingdings 2 font.

    Args:
        font: Source InlineFont object

    Returns:
        TextBlock with Wingdings 2 font and 'R' character (checkmark)
    """
    wfont = copy.copy(font)
    wfont.rFont = "Wingdings 2"
    return TextBlock(wfont, "R")


def no():
    """
    Create an empty checkbox.

    Returns:
        Unicode empty box character
    """
    return "□"


def yn(value, font):
    """
    Return checkmark or empty box based on value.

    Args:
        value: Boolean value
        font: Font for checkmark

    Returns:
        TextBlock (checkmark) or str (empty box)
    """
    if value:
        return yes(font)
    else:
        return no()


class YnExtension(Extension):
    """
    Jinja2 extension for yes/no checkbox rendering.

    Usage in template:
        {% yn is_approved %}           - Check if True
        {% yn is_rejected, True %}     - Inverted (check if False)
    """

    tags = {"yn"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]

        if parser.stream.skip_if("comma"):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(None))

        body = []
        return nodes.CallBlock(self.call_method("_yn", args), [], [], body).set_lineno(
            lineno
        )

    def _yn(self, arg0, arg1, caller):
        """
        Process yes/no checkbox rendering.

        Args:
            arg0: Boolean value
            arg1: Inversion flag (if not None, invert arg0)
            caller: Jinja2 caller

        Returns:
            Processed rich text value
        """
        segment = self.environment.node_map.current_node

        # Invert if arg1 is provided
        if arg1 is not None:
            arg0 = not arg0

        rv = yn(arg0, segment.font)
        rv = segment.process_rich_rv(rv)
        return rv
