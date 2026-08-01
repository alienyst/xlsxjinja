"""
Cell context for writing cells with formatting to Excel.

This module provides the CellContext class which acts as a bridge
between template cells and output cells, handling lazy cell creation,
value assignment, and filter application.
"""

import copy


class CellContext:
    """
    Context object for writing a cell to the output Excel file.

    Handles value writing, style copying, and filter application.

    Flow:
    1. Create context with cell node and value
    2. Access target_cell property when needed (lazy creation)
    3. Cell is created with value and formatting
    4. Apply filters if registered
    5. Finalize cell writing
    """

    def __init__(self, sheet_writer, cell_node, value, data_type):
        """
        Initialize cell context.

        Args:
            sheet_writer: SheetWriter instance
            cell_node: Source cell node from template
            value: Value to write
            data_type: Excel data type

        Flow:
        1. Store references to sheet writer and cell node
        2. Store value and data type
        3. Initialize target cell as None (lazy creation)
        """
        self.sheet_writer = sheet_writer
        self.cell_node = cell_node
        self.value = value
        self.data_type = data_type
        self._target_cell = None

    @property
    def rdsheet(self):
        """Source sheet from template."""
        return self.sheet_writer.rdsheet

    @property
    def wtsheet(self):
        """Target sheet in output."""
        return self.sheet_writer.wtsheet

    @property
    def source_cell(self):
        """Source cell from template."""
        return self.cell_node.sheet_cell

    @property
    def rdcolx(self):
        """Source column index."""
        return self.cell_node.colx

    @property
    def rdrowx(self):
        """Source row index."""
        return self.cell_node.rowx

    @property
    def wtcolx(self):
        """Target column index."""
        return self.sheet_writer.box.right

    @property
    def wtrowx(self):
        """Target row index."""
        return self.sheet_writer.box.bottom

    @property
    def target_cell(self):
        """
        Get or create target cell with value and formatting.

        Returns:
            openpyxl Cell object

        Flow:
        1. Return cached cell if already created
        2. If not created yet:
           a. Get target position from box
           b. Create cell at target position
           c. Set cell value (handle formulas, data types)
           d. Copy cell style from template
           e. Copy hyperlink if present
           f. Cache cell for future access
        3. Return target cell
        """
        if self._target_cell:
            return self._target_cell

        source = self.source_cell
        wtcolx = self.wtcolx
        wtrowx = self.wtrowx
        value = self.value
        data_type = self.data_type

        target = self.wtsheet.cell(column=wtcolx, row=wtrowx)

        # Set value
        if value is None:
            target._value = source._value
            target.data_type = source.data_type
        elif isinstance(value, (str, bytes)) and value.startswith("="):
            target.value = value
        elif data_type:
            target._value = value
            target.data_type = data_type
        else:
            target.value = value

        # Copy style
        if source.has_style:
            target._style = copy.copy(source._style)

        # Copy hyperlink
        if source.hyperlink:
            target.hyperlink = copy.copy(source.hyperlink)

        self._target_cell = target
        return self._target_cell

    def get_style(self):
        """Get cell style."""
        return self.target_cell.style

    def apply_filters(self):
        """Apply registered filters to cell."""
        if hasattr(self.cell_node, "filters") and self.cell_node.filters:
            for filter_func, args in self.cell_node.filters:
                filter_func(self, *args)
            self.cell_node.filters.clear()

    def finish(self):
        """Finalize cell writing by applying filters."""
        self.apply_filters()
