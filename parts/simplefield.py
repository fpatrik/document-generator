"""
Contains simple fields to automatically display page numbers etc.
"""
from lxml import etree


class SimpleField():
    """
    A Simple Field showing the page number etc.
    """
    def __init__(self, preset_styles, style_name = None, content = 'page', bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        
    
        self.content = content
        
        self.styles = preset_styles
        
        if style_name == None:
            style_name = 'conventec_default'
            
        style = self.styles.get(style_name, 'conventec_default')
        
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
            self.small_caps= small_caps
            
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
        
        
        self.type = 'simplefield'
        
    def render(self, root):
        """
        Adds a simplefield node to a given root node
        """
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'xml' : 'http://www.w3.org/XML/1998/namespace'}
            
        r_node = etree.SubElement(root, '{%s}r' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        #Set styles
        rpr_node = etree.SubElement(r_node, '{%s}rPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        if(self.bold or self.italics or self.underlined or self.small_caps):
            
            if self.bold:
                etree.SubElement(rpr_node, '{%s}b' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            if self.italics:
                etree.SubElement(rpr_node, '{%s}i' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            if self.underlined:
                underlined_node = etree.SubElement(rpr_node, '{%s}u' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                underlined_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "single") 
            if self.small_caps:
                etree.SubElement(rpr_node, '{%s}smallCaps' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        #Set font type
        rfonts_node = etree.SubElement(rpr_node, '{%s}rFonts' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        rfonts_node.set('{%s}ascii' % CURRENT_NAMESPACES['w'], self.font_type)
        rfonts_node.set('{%s}hAnsi' % CURRENT_NAMESPACES['w'], self.font_type)
        rfonts_node.set('{%s}cs' % CURRENT_NAMESPACES['w'], self.font_type)
        
        #Colors
        if self.text_color != None:
            color_node = etree.SubElement(rpr_node, '{%s}color' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            color_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.text_color)
            
        if self.highlight_color != None:
            highlight_node = etree.SubElement(rpr_node, '{%s}highlight' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            highlight_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.highlight_color)
        
        #Set font size 
        sz_node = etree.SubElement(rpr_node, '{%s}sz' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        sz_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(2*int(self.font_size)))
        
        szcs_node = etree.SubElement(rpr_node, '{%s}szCs' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        szcs_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(2*int(self.font_size)))
        
        #Vertical alignment
        if self.vertical_align != None:
            vert_node = etree.SubElement(rpr_node, '{%s}vertAlign' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            vert_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.vertical_align)
        
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "begin")
        instrtext_node = etree.SubElement(r_node, '{%s}instrText' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        instrtext_node.text = "PAGE"
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "separate")
        text_node = etree.SubElement(r_node, '{%s}t' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        text_node.text = "1"
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "end")
        
        