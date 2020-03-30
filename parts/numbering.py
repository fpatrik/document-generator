"""
Contains Numbered Titles with their references
"""

from lxml import etree
import roman
from conventec_docx.parts.text import Text
from conventec_docx.parts.breaks import LineBreak, PageBreak

class NumberedTitleTemplate():
    """
    A template for numbered titles
    """
    
    def __init__(self, id, text = '', style = 'numbering', separator = '.', **kwargs):
        self.id = id
        self.text = text
        self.style = style
        self.separator = separator
        self.numlinks = []
        
    def render(self):
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'}
        abstractnum_root = etree.Element('{%s}abstractNum' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        abstractnum_root.set('{%s}abstractNumId' % CURRENT_NAMESPACES['w'], str(self.id))
        abstractnum_root.set('{%s}restartNumberingAfterBreak' % CURRENT_NAMESPACES['w15'], "0")
            
        multilevel_node = etree.SubElement(abstractnum_root, '{%s}multiLevelType' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        multilevel_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "hybridMultilevel")
        
        
        if self.style == "numbering":
            numfmt_list = ["decimal"]
            lvltext_list = [self.text + '%1' + self.separator]
            
        elif self.style == "roman":
            numfmt_list = ["upperRoman"]
            lvltext_list = [self.text + '%1' + self.separator]

            
        for level in range(len(numfmt_list)):
            lvl_node = etree.SubElement(abstractnum_root, '{%s}lvl' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvl_node.set('{%s}ilvl' % CURRENT_NAMESPACES['w'], str(level))
            if level > 0:
                lvl_node.set('{%s}tentative' % CURRENT_NAMESPACES['w'], "1")
            start_node = etree.SubElement(lvl_node, '{%s}start' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            start_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "1")
            numfmt_node = etree.SubElement(lvl_node, '{%s}numFmt' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            numfmt_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(numfmt_list[level]))
            suff_node = etree.SubElement(lvl_node, '{%s}suff' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            suff_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "space")
            lvltext_node = etree.SubElement(lvl_node, '{%s}lvlText' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvltext_node.set('{%s}val' % CURRENT_NAMESPACES['w'], lvltext_list[level])
            lvljc_node = etree.SubElement(lvl_node, '{%s}lvlJc' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            lvljc_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "left")
            ppr_node = etree.SubElement(lvl_node, '{%s}pPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            ind_node = etree.SubElement(ppr_node, '{%s}ind' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            ind_node.set('{%s}left' % CURRENT_NAMESPACES['w'], "360")
            ind_node.set('{%s}hanging' % CURRENT_NAMESPACES['w'], "360")
            
        
        return abstractnum_root
    

class NumberedTitle():
    """
    Numbered titles in the document
    """
    
    def __init__(self, title_template, n_of_lists, preset_styles, style_name = None, **kwargs):
        self.numid = n_of_lists
        self.templateid = title_template.id
        title_template.numlinks.append(self.render_num_link())
        
        self.style_name = style_name
        self.styles = preset_styles
        
        self.text = title_template.text
        self.style = title_template.style
        self.separator = title_template.separator
        
        self.current_number = 1
    
    def render_num_link(self):
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        num_root = etree.Element('{%s}num' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        num_root.set('{%s}numId' % CURRENT_NAMESPACES['w'], str(self.numid))
        abstractnumid_node = etree.SubElement(num_root, '{%s}abstractNumId' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        abstractnumid_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(self.templateid))
        
        return num_root
    
class Title():
    """
    Is a numbered title in the document
    """
    def __init__(self, list, preset_styles, style_name = None, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Initialise title
        """
        self.parts = []
        self.numid = list.numid
        self.number = list.current_number
        list.current_number += 1
        
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
        
    
        self.text = list.text
        self.style = list.style
        self.separator = list.separator
        
        self.type = 'title'
    
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
        
        new_line_break = LineBreak(preset_styles, style_name = style_name, n = n, bold = self.bold, italics=self.italics, underlined=self.underlined, small_caps = self.small_caps, font_type = self.font_type, font_size=self.font_size, text_color = self.text_color, highlight_color = self.highlight_color, vertical_align = self.vertical_align)
        self.parts.append(new_line_break)
        return new_line_break
    
    def add_page_break(self, style_name = None, n=1, **kwargs):
        """
        Adds a page break
        """
        preset_styles = self.styles
        
        new_page_break = PageBreak(preset_styles, style_name = style_name, n = n, bold = self.bold, italics=self.italics, underlined=self.underlined, small_caps = self.small_caps, font_type = self.font_type, font_size=self.font_size, text_color = self.text_color, highlight_color = self.highlight_color, vertical_align = self.vertical_align)
        self.parts.append(new_page_break)
        return new_page_break
    
    def use_image(self, image, **kwargs):
        """
        Use an Image
        """
        self.parts.append(image)
        return image
    
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
        ilvl_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "0")
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
        
        bookmark_node = etree.SubElement(new_root, '{%s}bookmarkStart' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        bookmark_node.set('{%s}id' % CURRENT_NAMESPACES['w'], "0")
        bookmark_node.set('{%s}name' % CURRENT_NAMESPACES['w'], "_Ref" + str(self.numid) + '-' + str(self.number))
        bookmark_node = etree.SubElement(new_root, '{%s}bookmarkEnd' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        bookmark_node.set('{%s}id' % CURRENT_NAMESPACES['w'], "0")
        
        for part in self.parts:
            part.render(new_root)
            
class Reference():
    """
    A Reference to a title
    """

    def __init__(self, title, preset_styles, style_name = None, reference = 'title', bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Initialises the text node with default empty default text
        """
        self.title = title
        self.reference = reference
        
        self.styles = preset_styles
        
        if style_name == None:
            style_name = 'conventec_default'
            
        style = self.styles.get(style_name, self.styles.get('conventec_default'))
        
        
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
            
        self.type = 'reference'
        
        
    def render(self, root):
        """
        Adds a text node to a given root node
        """
        
        if self.reference == 'title':
            reference_string = " REF _Ref" + str(self.title.numid) + '-' + str(self.title.number) + " \r \h "
        
        elif self.reference == 'page':
            reference_string = 'PAGE'
            
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
        
        #Vertical alignment
        if self.vertical_align != None:
            vert_node = etree.SubElement(rpr_node, '{%s}vertAlign' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            vert_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.vertical_align)
        
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "begin")
        instrtext_node = etree.SubElement(r_node, '{%s}instrText' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        instrtext_node.set('{%s}space' % CURRENT_NAMESPACES['xml'], "preserve")
        instrtext_node.text = reference_string
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "separate")
        text_node = etree.SubElement(r_node, '{%s}t' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        if self.title.style == 'roman':
            number = roman.toRoman(self.title.number)
        elif self.title.style == 'numbering':
            number = str(self.title.number)
            
        if self.reference == 'title':
            text = self.title.text + number + str(self.title.separator)
        
        elif self.reference == 'page':
            text = 'UPDATE FIELDS PLEASE'
            
        text_node.text = text
        fldchar_node = etree.SubElement(r_node, '{%s}fldChar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        fldchar_node.set('{%s}fldCharType' % CURRENT_NAMESPACES['w'], "end")
