"""Custom Jinja2 extensions for Excel template rendering."""

import os
from inspect import isfunction

from jinja2 import nodes
from jinja2.ext import Extension
from jinja2.runtime import Undefined

# Check for PIL availability
try:
    from PIL.ImageFile import ImageFile

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class NodeExtension(Extension):
    """Extension for marking node positions in template (row, cell, node tags)."""

    tags = {"row", "cell", "node", "extra"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = []
        return nodes.CallBlock(
            self.call_method("_node", args), [], [], body
        ).set_lineno(lineno)

    def _node(self, key, caller):
        """Get node from node map and trigger enter/exit lifecycle."""
        node = self.environment.node_map.get_node(key)
        return str(key)


class SegmentExtension(Extension):
    """Extension for wrapping text segments (seg/endseg tags)."""

    tags = {"seg"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        body = parser.parse_statements(["name:endseg"], drop_needle=True)
        return nodes.CallBlock(self.call_method("_seg", args), [], [], body).set_lineno(
            lineno
        )

    def _seg(self, key, caller):
        """Process segment content."""
        segment = self.environment.node_map.get_node(key)
        rv = caller()
        rv = segment.process_rv(rv)
        return rv


class XvExtension(Extension):
    """Extension for non-string variable values (xv tag)."""

    tags = {"xv"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]

        if parser.stream.skip_if("comma"):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(0))

        body = []
        return nodes.CallBlock(self.call_method("_xv", args), [], [], body).set_lineno(
            lineno
        )

    def _xv(self, xv, key, caller):
        """
        Handle variable value rendering.

        Stores the actual value (number, date, etc.) in XvCell
        and returns string representation for template continuity.
        """
        if key == 0:
            return str(xv)

        xvcell = self.environment.node_map.get_node(key)

        if xv is None or type(xv) is Undefined:
            xv = ""

        xvcell.rv = xv
        return str(xv)


class OpExtension(Extension):
    """Extension for post-render operations (op tag)."""

    tags = {"op"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]
        func_args = []

        while parser.stream.skip_if("comma"):
            func_args.append(parser.parse_expression())

        args.append(nodes.List(func_args))
        body = []

        return nodes.CallBlock(self.call_method("_op", args), [], [], body).set_lineno(
            lineno
        )

    def _op(self, func, func_args, caller):
        """Register operation to be executed after cell is written."""
        if isfunction(func):
            node = self.environment.node_map.current_node
            node.add_op((func, func_args))
        return str(func)


class ImageRef:
    """Reference to an image to be inserted into a cell."""

    def __init__(self, image, image_index):
        """
        Initialize image reference.

        Args:
            image: PIL Image object or path to image file
            image_index: Index for multiple images in same cell
        """
        self.image = image
        self.image_index = image_index
        self.rdrowx = -1
        self.rdcolx = -1
        self.wtrowx = -1
        self.wtcolx = -1

        # Validate image path if not PIL Image
        if PIL_AVAILABLE and not isinstance(image, ImageFile):
            fname = str(image)
            if not os.path.exists(fname):
                self.image = None

    @property
    def image_key(self):
        """Unique key for tracking image position."""
        return (self.rdrowx, self.rdcolx, self.image_index)

    @property
    def wt_top_left(self):
        """Output position (top-left corner)."""
        return (self.wtrowx, self.wtcolx)


class ImageExtension(Extension):
    """Extension for inserting images into cells (img tag)."""

    tags = {"img"}

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        args = [parser.parse_expression()]

        if parser.stream.skip_if("comma"):
            args.append(parser.parse_expression())
        else:
            args.append(nodes.Const(0))

        body = []
        return nodes.CallBlock(
            self.call_method("_image", args), [], [], body
        ).set_lineno(lineno)

    def _image(self, image, image_index, caller):
        """
        Handle image insertion.

        Args:
            image: PIL Image or path to image file
            image_index: Index for multiple images

        Returns:
            'image' placeholder string
        """
        if not PIL_AVAILABLE:
            return ""

        image_ref = ImageRef(image, image_index)

        if image_ref.image:
            node = self.environment.node_map.current_node
            node.set_image_ref(image_ref)

        return "image"
