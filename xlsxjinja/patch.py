# -*- coding: utf-8 -*-
"""
Patches for openpyxl to fix image handling issues.

This module applies runtime patches to openpyxl classes to fix:
1. Image deduplication during Excel file writing
2. Geometry guide list XML namespace issues
"""

from openpyxl.drawing.geometry import GeomGuideList
from openpyxl.writer.excel import ExcelWriter
from openpyxl.xml.constants import DRAWING_NS

from .logger import get_logger

logger = get_logger(__name__)


class ExWriter(ExcelWriter):
    """
    Patched ExcelWriter with improved image handling.

    Flow:
    1. Inherits from openpyxl.writer.excel.ExcelWriter
    2. Overrides _write_images method
    3. Adds image deduplication logic
    4. Prevents duplicate images in output file
    """

    def _write_images(self):
        """
        Write images to Excel file with deduplication.

        Flow:
        1. Initialize tracking dict for written images
        2. Iterate through all images
        3. Check if image has unique key
        4. Skip if already written (deduplication)
        5. Write image data to archive
        6. Mark image as written
        """
        _written = {}
        for img in self._images:
            if hasattr(img, "key"):
                key = img.key
                _w = _written.get(key)
                if _w:
                    continue
                else:
                    _written[key] = True
            else:
                logger.warning(
                    f"Image {img} does not have 'key' attribute - not an Img instance"
                )
            self._archive.writestr(img.path[1:], img._data())


# Apply the patch to ExcelWriter
ExcelWriter._write_images = ExWriter._write_images

# Fix geometry guide list XML namespace issue
GeomGuideList.tagname = "avLst"
GeomGuideList.namespace = DRAWING_NS

try:
    import datetime
    from openpyxl.packaging.core import NestedDateTime

    _original_to_tree = NestedDateTime.to_tree

    def _patched_to_tree(self, tagname=None, value=None, namespace=None):
        if value is not None and value.tzinfo is not None:
            value = value.astimezone(datetime.timezone.utc).replace(tzinfo=None)
        return _original_to_tree(self, tagname, value, namespace)

    NestedDateTime.to_tree = _patched_to_tree
except Exception as exc:
    logger.warning(f"Unable to patch openpyxl NestedDateTime: {exc}")
