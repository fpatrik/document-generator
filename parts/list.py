"""
Represents an abstract list template
"""
from lxml import etree
from conventec_docx.parts.text import Text
from conventec_docx.parts.breaks import LineBreak, PageBreak

class ListTemplate():
    """
    A template for a list
    """
    
    def __init__(self, id, indent = "1", type = "numbering", **kwargs):
        self.id = id
        self.indent = indent
        self.type = type
        self.numlinks = []
        
    def render(self):
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'}
        abstractnum_root = etree.Element('{%s}abstractNum' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        abstractnum_root.set('{%s}abstractNumId' % CURRENT_NAMESPACES['w'], str(self.id))
        abstractnum_root.set('{%s}restartNumberingAfterBreak' % CURRENT_NAMESPACES['w15'], "0")
            
        multilevel_node = etree.SubElement(abstractnum_root, '{%s}multiLevelType' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        multilevel_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "hybridMultilevel")
        
        
        if self.type == "numbering":
            numfmt_list = ["decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman"]
            lvltext_list = ["%1.", "%2.","%3.",'%4.','%5.','%6.','%7.','%8.', '%9.']
            
        elif self.type == "bullet":
            numfmt_list = ["bullet"] * 9
            lvltext_list = [u"\u25CF", u"\u25A0",u"\u25B6",u"\u25C7",u"\u25CF",u"\u25A0",u"\u25B6",u"\u25C7", u"\u25CF"]
        
        elif self.type == "list":
            numfmt_list = ["bullet"] * 9
            lvltext_list = [u"\u2013", u"\u25CF", u"\u25A0",u"\u25B6",u"\u25C7",u"\u25CF",u"\u25A0",u"\u25B6",u"\u25C7"]
        
        elif self.type == "roman":
            numfmt_list = ["upperRoman", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman"]
            lvltext_list = ["%1.", "%2.","%3.",'%4.','%5.','%6.','%7.','%8.', '%9.']
            
        elif self.type == 'letter':
            numfmt_list = ["lowerLetter", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman", "decimal", "lowerLetter", "lowerRoman"]
            lvltext_list = ["%1)", "%2.","%3.",'%4.','%5.','%6.','%7.','%8.', '%9.']

            
        for level in range(len(numfmt_list)):
            lvl_node = etree.SubElement(abstractnum_root, '{%s}lvl' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvl_node.set('{%s}ilvl' % CURRENT_NAMESPACES['w'], str(level))
            if level > 0:
                lvl_node.set('{%s}tentative' % CURRENT_NAMESPACES['w'], "1")
            start_node = etree.SubElement(lvl_node, '{%s}start' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            start_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "1")
            numfmt_node = etree.SubElement(lvl_node, '{%s}numFmt' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            numfmt_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(numfmt_list[level]))
            lvltext_node = etree.SubElement(lvl_node, '{%s}lvlText' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvltext_node.set('{%s}val' % CURRENT_NAMESPACES['w'], lvltext_list[level])
            lvljc_node = etree.SubElement(lvl_node, '{%s}lvlJc' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvljc_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "left")
            ppr_node = etree.SubElement(lvl_node, '{%s}pPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            ind_node = etree.SubElement(ppr_node, '{%s}ind' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            ind_node.set('{%s}left' % CURRENT_NAMESPACES['w'], str((int(self.indent) + level) * 720))
            ind_node.set('{%s}hanging' % CURRENT_NAMESPACES['w'], "360")
        
        return abstractnum_root
    
class List():
    """
    Is an actual list in the document
    """
    def __init__(self, list_template, n_of_lists, style_name = None, **kwargs):
        self.numid = n_of_lists
        self.style_name = style_name
        
        self.templateid = list_template.id
        
        list_template.numlinks.append(self.render_num_link())
    
    def render_num_link(self, **kwargs):
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        num_root = etree.Element('{%s}num' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        num_root.set('{%s}numId' % CURRENT_NAMESPACES['w'], str(self.numid))
        abstractnumid_node = etree.SubElement(num_root, '{%s}abstractNumId' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        abstractnumid_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(self.templateid))
        
        return num_root
        
class ListPoint():
    """
    Is A list point of a given list
    """
    def __init__(self, list, preset_styles, style_name = None, level = 0, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Initialise the parts of the paragraph
        """
        self.parts = []
        self.numid = list.numid
        self.level = level
        
        self.styles = preset_styles
        
        if style_name == None:
            if list.style_name != None:
                self.style_name = list.style_name
            else:
                self.style_name = 'conventec_default'
        else:
            self.style_name = style_name
            
        style = self.styles.get(self.style_name, 'conventec_default')
        
        if alignment == None:
            self.alignment = style.alignment
        else:
            self.alignment = alignment
            
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
        
        
        self.type = 'listpoint'
    
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
        """
        Use an Image
        """
        self.parts.append(image)
        return image
    
    def add_reference(self, style_name = None, title = None, reference = 'title', bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a reference to a title
        """
        preset_styles = self.styles
        
        new_reference = Reference(title, preset_styles, style_name = style_name,  reference = reference, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_reference)
        return new_reference
    
    def render(self, root):
        """
        Renders the list point
        """
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        new_root = etree.SubElement(root, '{%s}p' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        ppr_root = etree.SubElement(new_root, '{%s}pPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        pstyle_root = etree.SubElement(ppr_root, '{%s}pStyle' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        pstyle_root.set('{%s}val' % CURRENT_NAMESPACES['w'], "ListParagraph")
        
        #List options
        numpr_root = etree.SubElement(ppr_root, '{%s}numPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        ilvl_node = etree.SubElement(numpr_root, '{%s}ilvl' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        ilvl_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(self.level))
        numid_node = etree.SubElement(numpr_root, '{%s}numId' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        numid_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(self.numid))
        
        #Alignment
        jc_root = etree.SubElement(ppr_root, '{%s}jc' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        jc_root.set('{%s}val' % CURRENT_NAMESPACES['w'], self.alignment)
        
        #Keep next
        if self.keep_next:
            etree.SubElement(ppr_root, '{%s}keepNext' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            
        #Spacing
        spacing_root = etree.SubElement(ppr_root, '{%s}spacing' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        spacing_root.set('{%s}before' % CURRENT_NAMESPACES['w'], str(self.spacing_before))
        spacing_root.set('{%s}after' % CURRENT_NAMESPACES['w'], str(self.spacing_after))
        spacing_root.set('{%s}line' % CURRENT_NAMESPACES['w'], str(int(240*float(self.spacing_line))))
        spacing_root.set('{%s}lineRule' % CURRENT_NAMESPACES['w'], 'auto')
        
        contextualspacing_root = etree.SubElement(ppr_root, '{%s}contextualSpacing' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        contextualspacing_root.set('{%s}val' % CURRENT_NAMESPACES['w'], "0")
        
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
        
        rpr_root = etree.SubElement(ppr_root, '{%s}rPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        
        for part in self.parts:
            part.render(new_root)