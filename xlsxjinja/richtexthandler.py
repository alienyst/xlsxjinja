"""
Handler for rich text (text with multiple fonts) in Excel cells.

Rich text in Excel can have Jinja2 tags split across multiple segments
with different fonts. This module handles:
1. Detecting split tags
2. Fixing broken tags by merging segments
3. Creating rich text output with proper formatting
"""

from copy import copy

from openpyxl.cell.rich_text import CellRichText, TextBlock

from .utils import fix_test, tag_fix


class RichTextHandler:
    """
    Handler for rich text with multiple fonts.

    Handles parsing, fixing broken Jinja2 tags in rich text,
    and reconstructing rich text after template rendering.

    Flow:
    1. Parse rich text into segments
    2. Detect if Jinja2 tags are split across segments
    3. Fix broken tags by merging text
    4. Render template with fixed tags
    5. Reconstruct rich text with original formatting
    """

    @classmethod
    def iter(cls, rich_text, font):
        """
        Iterate through rich text segments with tag fixing.

        Args:
            rich_text: CellRichText object
            font: Default font

        Yields:
            (text, font, segment) tuples

        Flow:
        1. Extract segment font (from TextBlock or use default)
        2. Get combined text from all segments
        3. Test if Jinja2 tags are broken across segments
        4. If broken tags detected:
           a. Fix tags by merging segments
           b. Yield fixed text with correct fonts
        5. If no broken tags:
           a. Yield segments as-is
        """

        def get_segment_font(segment):
            if isinstance(segment, TextBlock):
                return segment.font or font
            return font

        text_4_fix = cls.text_4_fix(rich_text)

        if fix_test(text_4_fix):
            fixed = tag_fix(text_4_fix)
            for i, segment in enumerate(rich_text):
                if i in fixed:
                    text = fixed[i]
                    if text == "":
                        continue
                    else:
                        segment_font = get_segment_font(segment)
                        yield text, segment_font, segment
        else:
            for segment in rich_text:
                segment_font = get_segment_font(segment)
                yield str(segment), segment_font, segment

    @classmethod
    def rich_segment(cls, text, font):
        """
        Create a rich text segment.

        Args:
            text: Text content
            font: InlineFont object

        Returns:
            TextBlock object
        """
        return TextBlock(font, text)

    @classmethod
    def text_content(cls, value):
        """
        Extract plain text from value.

        Args:
            value: Any value (str or CellRichText)

        Returns:
            Plain text string
        """
        return value

    @classmethod
    def text_4_fix(cls, rich_text):
        """
        Prepare text for tag fixing by adding markers.

        Args:
            rich_text: CellRichText object

        Returns:
            Text with segment markers
        """
        text = []
        fmt = "___%d___"
        for i, segment in enumerate(rich_text):
            text.append(fmt % i)
            text.append(str(segment))
        return "".join(text)

    @classmethod
    def rich_content(cls, value):
        """
        Create rich text content from list of segments.

        Args:
            value: List of TextBlock objects

        Returns:
            CellRichText object
        """
        return CellRichText(value)

    @classmethod
    def mid(cls, rich_text, head, tail):
        """
        Extract substring from rich text with font preservation.

        Args:
            rich_text: CellRichText object
            head: Start position
            tail: End position

        Returns:
            (CellRichText, plain_text) tuple
        """
        st = 0
        end = -1
        segments = []
        texts = []

        def get_segment_copy(segment, text):
            if isinstance(segment, TextBlock):
                segment_copy = copy(segment)
                segment_copy.text = text
                return segment_copy
            return text

        for index, segment in enumerate(rich_text):
            segment_text = str(segment)
            l_text = len(segment_text)
            st = end + 1
            end += l_text

            if end < head:
                continue
            elif st <= head <= end:
                if end < tail:
                    text_st = head - st
                    text = segment_text[text_st:]
                    segment_copy = get_segment_copy(segment, text)
                    segments.append(segment_copy)
                    texts.append(text)
                else:
                    text_st = head - st
                    text_end = tail - st
                    text = segment_text[text_st : text_end + 1]
                    segment_copy = get_segment_copy(segment, text)
                    segments.append(segment_copy)
                    texts.append(text)
                    break
            elif end < tail:
                segment_copy = copy(segment)
                text = segment.text
                segments.append(segment_copy)
                texts.append(text)
            else:
                text_end = tail - st
                text = segment_text[: text_end + 1]
                segment_copy = get_segment_copy(segment, text)
                segments.append(segment_copy)
                texts.append(text)
                break

        return CellRichText(segments), "".join(texts)


# Global instance
rich_handler = RichTextHandler()
