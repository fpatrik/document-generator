"""
Represents a paragraph within a document
"""
from lxml import etree
from conventec_docx.parts.text import Text
from conventec_docx.parts.breaks import LineBreak, PageBreak
from conventec_docx.parts.simplefield import SimpleField
from conventec_docx.parts.numbering import Reference

class Paragraph():
    """
    Represents a paragraph within a document
    """
    
    def __init__(self, preset_styles, style_name = None, alignment = None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line=None, indent = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Initialise the parts of the paragraph
        """
        self.parts = []
        
        self.styles = preset_styles
        
        if style_name == None:
            self.style_name = 'conventec_default'
        else:
            self.style_name = style_name
            
        style = self.styles.get(self.style_name, 'conventec_default')
        
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
        
        
        self.type = 'paragraph'
    
    def add_text(self, text, style_name = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size = None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Adds a text node
        """
        preset_styles = self.styles
        
        if style_name == None:
            style_name = self.style_name
        
        new_text = Text(preset_styles, text = text, style_name = style_name, bold = bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size = font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_text)
        return new_text
    
    def add_line_break(self, style_name = None, n=1, **kwargs):
        """
        Adds a line break
        """
        preset_styles = self.styles
        
        if style_name == None:
            style_name = self.style_name
        
        new_line_break = LineBreak(preset_styles, style_name = style_name, n = n, bold = self.bold, italics=self.italics, underlined=self.underlined, small_caps = self.small_caps, font_type = self.font_type, font_size=self.font_size, text_color = self.text_color, highlight_color = self.highlight_color, vertical_align = self.vertical_align)
        self.parts.append(new_line_break)
        return new_line_break
    
    def add_page_break(self, style_name = None, n=1, **kwargs):
        """
        Adds a page break
        """
        preset_styles = self.styles
        
        if style_name == None:
            style_name = self.style_name
        
        new_page_break = PageBreak(preset_styles, style_name = style_name, n = n, bold = self.bold, italics=self.italics, underlined=self.underlined, small_caps = self.small_caps, font_type = self.font_type, font_size=self.font_size, text_color = self.text_color, highlight_color = self.highlight_color, vertical_align = self.vertical_align)
        self.parts.append(new_page_break)
        return new_page_break
    
    def use_image(self, image, **kwargs):
        self.parts.append(image)
        return image
    
    def add_simplefield(self, style_name = None, content = "page", bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size = None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Adds a simplefield to the paragraph
        """
        preset_styles = self.styles
        if style_name == None:
            style_name = self.style_name
            
        new_simplefield = SimpleField(preset_styles, style_name = style_name, content = content, bold = bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_simplefield)
        return new_simplefield
    
    def add_reference(self, title, style_name = None, reference = 'title', bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a reference to a title
        """
        preset_styles = self.styles
        
        if style_name == None:
            style_name = self.style_name
        
        new_reference = Reference(title, preset_styles, style_name = style_name,  reference = reference, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_reference)
        return new_reference
    
    def render(self, root):
        """
        Adds the paragraph to a given root node
        """
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        new_root = etree.SubElement(root, '{%s}p' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        ppr_root = etree.SubElement(new_root, '{%s}pPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        #Alignment
        jc_root = etree.SubElement(ppr_root, '{%s}jc' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        jc_root.set('{%s}val' % CURRENT_NAMESPACES['w'], self.alignment)
        
        #Border Bottom
        if self.border_bottom:
            pbdr_node = etree.SubElement(ppr_root, '{%s}pBdr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            bottom_node = etree.SubElement(pbdr_node, '{%s}bottom' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            bottom_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "single")
            bottom_node.set('{%s}sz' % CURRENT_NAMESPACES['w'], "4")
            bottom_node.set('{%s}space' % CURRENT_NAMESPACES['w'], "1")
            bottom_node.set('{%s}color' % CURRENT_NAMESPACES['w'], "auto")
        
        #Keep next
        if self.keep_next:
            etree.SubElement(ppr_root, '{%s}keepNext' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            
        #Spacing
        spacing_root = etree.SubElement(ppr_root, '{%s}spacing' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        spacing_root.set('{%s}before' % CURRENT_NAMESPACES['w'], str(self.spacing_before))
        spacing_root.set('{%s}after' % CURRENT_NAMESPACES['w'], str(self.spacing_after))
        spacing_root.set('{%s}line' % CURRENT_NAMESPACES['w'], str(int(240*float(self.spacing_line))))
        spacing_root.set('{%s}lineRule' % CURRENT_NAMESPACES['w'], 'auto')
        
        #Indent
        indent_root = etree.SubElement(ppr_root, '{%s}ind' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        indent_root.set('{%s}left' % CURRENT_NAMESPACES['w'], str(int(self.indent) * 720))
        
        rpr_root = etree.SubElement(ppr_root, '{%s}rPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        if self.bold:
            etree.SubElement(rpr_root, '{%s}b' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        if self.italics:
            etree.SubElement(rpr_root, '{%s}i' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        if self.underlined:
            etree.SubElement(rpr_root, '{%s}u' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        if self.small_caps:
            etree.SubElement(rpr_root, '{%s}smallCaps' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            
        #Font Type
        rfonts_node = etree.SubElement(rpr_root, '{%s}rFonts' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        rfonts_node.set('{%s}ascii' % CURRENT_NAMESPACES['w'], self.font_type)
        rfonts_node.set('{%s}hAnsi' % CURRENT_NAMESPACES['w'], self.font_type)
        rfonts_node.set('{%s}cs' % CURRENT_NAMESPACES['w'], self.font_type)
            
        #Set font size 
        sz_node = etree.SubElement(rpr_root, '{%s}sz' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        sz_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(2*int(self.font_size)))
        
        szcs_node = etree.SubElement(rpr_root, '{%s}szCs' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        szcs_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(2*int(self.font_size)))
        
        #Vertical alignment
        if self.vertical_align != None:
            vert_node = etree.SubElement(rpr_node, '{%s}vertAlign' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            vert_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.vertical_align)
        
        for part in self.parts:
            part.render(new_root)
            