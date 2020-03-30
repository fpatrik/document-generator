"""
Allows defining default styles in for the document
"""

class Style():
    """
    A predefined style for the document
    """
    
    def __init__(self, document, reference_name, level = 0, alignment=None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, indent = None, bold=None, italics=None, underlined=None, small_caps=None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        self.reference_name = reference_name
        
        if self.reference_name == 'default':
            self.alignment = alignment
            self.border_bottom = border_bottom
            self.keep_next = keep_next
            self.spacing_before = spacing_before
            self.spacing_after = spacing_after
            self.spacing_line = spacing_line
            self.indent = indent
            self.bold = bold
            self.italics = italics
            self.underlined = underlined
            self.small_caps = small_caps
            self.font_type = font_type
            self.font_size = font_size
            self.text_color = text_color
            self.highlight_color = highlight_color
            self.vertical_align = vertical_align
            
        else:            
            style = document.styles.get('conventec_default')
            
            if alignment == None:
                self.alignment = style.alignment
            else:
                self.alignment = alignment
                
            if border_bottom == None:
                self.border_bottom = style.border_bottom
            else:
                self.border_bottom = border_bottom
                
            if keep_next == None:
                self.keep_next = style.keep_next
            else:
                self.keep_next = keep_next
                
            if spacing_before == None:
                self.spacing_before = style.spacing_before
            else:
                self.spacing_before = spacing_before
                
            if spacing_after == None:
                self.spacing_after = style.spacing_after
            else:
                self.spacing_after = spacing_after
                
            if spacing_line == None:
                self.spacing_line = style.spacing_line
            else:
                self.spacing_line = spacing_line
            
            if indent == None:
                self.indent = style.indent
            else:
                self.indent = indent
            
            if bold == None:
                self.bold = style.bold
            else:
                self.bold = bold
                
            if italics == None:
                self.italics = style.italics
            else:
                self.italics = italics
                
            if underlined == None:
                self.underlined = style.underlined
            else:
                self.underlined = underlined
                
            if small_caps == None:
                self.small_caps = style.small_caps
            else:
                self.small_caps = small_caps
                
            if font_type == None:
                self.font_type = style.font_type
            else:
                self.font_type = font_type
                
            if font_size == None:
                self.font_size = style.font_size
            else:
                self.font_size = font_size
            
            if text_color == None:
                self.text_color = style.text_color
            else:
                self.text_color = text_color
                
            if highlight_color == None:
                self.highlight_color = style.highlight_color
            else:
                self.highlight_color = highlight_color
                
            if vertical_align == None:
                self.vertical_align = style.vertical_align
            else:
                self.vertical_align = vertical_align