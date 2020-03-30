"""
Creates a header for a document
"""

from conventec_docx.parts.list import ListPoint
from conventec_docx.parts.paragraph import Paragraph
from conventec_docx.parts.table import Table, Cell

class Header():
    """
    A header of a document
    """
    
    def __init__(self, preset_styles, style_name = None, **kwargs):
        self.styles = preset_styles
        self.style_name = style_name
        
        self.even = SubHeader(preset_styles, style_name = style_name)
        self.default = SubHeader(preset_styles, style_name = style_name)
        self.first = SubHeader(preset_styles, style_name = style_name)
        self.type = 'header'
        
    def render(self, type, root):
        """
        Renders the content of the header
        """
        if type == 'even':
            subheader = self.even
        elif type == 'default':
            subheader = self.default
        elif type == 'first':
            subheader = self.first
            
        for part in subheader:
            part.render(root)
            

class SubHeader():
    """
    Represents any of the three header types even, default and first.
    """
    
    def __init__(self, preset_styles, style_name = None, **kwargs):
        
        self.styles = preset_styles
        self.style_name = style_name
        
        self.parts = []
        self.type = 'subheader'
        
    def add_paragraph(self, style_name = None, alignment=None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, indent = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Append a paragraph to the subheader
        """
        new_paragraph = Paragraph(preset_styles = self.styles, style_name = style_name, alignment = alignment, border_bottom = border_bottom, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, indent = indent, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_paragraph)
        return new_paragraph
    
    def add_list_point(self, list, style_name = None, level = 0, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a list template to the subheader
        """
        new_list_point = ListPoint(list, preset_styles = self.styles, style_name = style_name, level = level, alignment = alignment, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_list_point)
        return new_list_point
    
    def add_table(self, style_name = None, rows = 1, columns = 1, width = 1, alignment = "left", style = 'default', delete_empty = False, **kwargs):
        """
        Add a list template to the subheader
        """
        new_table = Table(preset_styles = self.styles, style_name = style_name, rows = rows, columns = columns, width = width, alignment = alignment, style=style, delete_empty = delete_empty)
        self.parts.append(new_table)
        return new_table
    
    def render(self, root):
        for part in self.parts:
            part.render(root)