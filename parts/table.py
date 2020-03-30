"""
Cotains tables of a document
"""
from lxml import etree
from conventec_docx.parts.paragraph import Paragraph
from conventec_docx.parts.list import ListPoint

class Table():
    """
    Table in a document
    """
    
    def __init__(self, preset_styles, style_name = None, rows = 1, columns = 1, width = 1, alignment = "left", style = 'default', delete_empty = False, **kwargs):        
        self.styles = preset_styles
        self.style_name = style_name
        
        
        self.rows = int(rows)
        self.columns = int(columns)
        self.style = style
        self.cells = []
        self.type = "table"
        self.width = float(width)
        self.alignment = alignment
        self.column_widths = [int(11900*width / self.columns)] * self.columns
        self.row_heights = [None] * self.rows
        self.delete_empty = delete_empty
        
        for row in range(int(self.rows)):
            self.cells.append([])
            for column in range(int(self.columns)):
                new_cell = Cell(self.styles, style_name = self.style_name, width = self.column_widths[column])
                self.cells[-1].append(new_cell)
                
    def add_row(self, **kwargs):
        """
        Adds a row to the table
        """
        self.rows += 1
        self.cells.append([])
        for column in range(int(self.columns)):
            new_cell = Cell(self.styles, style_name = self.style_name, width = self.column_widths[column])
            self.cells[-1].append(new_cell)
        
        self.row_heights.append(None)
                
    def render(self, root):
        #Update Cell widths
        for row in range(int(self.rows)):
            for column in range(int(self.columns)):
                self.cells[row][column].width = int(self.column_widths[column])
                
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tbl_root = etree.SubElement(root, '{%s}tbl' % CURRENT_NAMESPACES['w'])
        tblPr_node = etree.SubElement(tbl_root, '{%s}tblPr' % CURRENT_NAMESPACES['w'])
        
        tblstyle_node = etree.SubElement(tblPr_node, '{%s}tblStyle' % CURRENT_NAMESPACES['w'])
        tblstyle_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "TableGrid")
        
        if self.style == 'default':
            pass
            
        elif self.style == 'borderless':
            tblborders_node = etree.SubElement(tblPr_node, '{%s}tblBorders' % CURRENT_NAMESPACES['w'])
            for name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                current_node = etree.SubElement(tblborders_node, ('{%s}' + name) % CURRENT_NAMESPACES['w'])
                current_node.set('{%s}val' % CURRENT_NAMESPACES['w'], "none")
                current_node.set('{%s}sz' % CURRENT_NAMESPACES['w'], "0")
                current_node.set('{%s}space' % CURRENT_NAMESPACES['w'], "0")
                current_node.set('{%s}color' % CURRENT_NAMESPACES['w'], "auto")
        
        tblw_node = etree.SubElement(tblPr_node, '{%s}tblW' % CURRENT_NAMESPACES['w'])
        tblw_node.set('{%s}w' % CURRENT_NAMESPACES['w'], "0")
        tblw_node.set('{%s}type' % CURRENT_NAMESPACES['w'], "auto")
        jc_node = etree.SubElement(tblPr_node, '{%s}jc' % CURRENT_NAMESPACES['w'])
        jc_node.set('{%s}val' % CURRENT_NAMESPACES['w'], self.alignment)
        tbl_grid =  etree.SubElement(tbl_root, '{%s}tblGrid' % CURRENT_NAMESPACES['w'])
        
        for width in self.column_widths:
            gridcol_node =  etree.SubElement(tbl_grid, '{%s}gridCol' % CURRENT_NAMESPACES['w'])
            gridcol_node.set('{%s}w' % CURRENT_NAMESPACES['w'], str(width))
            
        for i in range(len(self.cells)):
            #If delete_empty, only render non empty rows
            if not self.delete_empty or len(self.cells[i][0].parts) > 0:
                tr_node = etree.SubElement(tbl_root, '{%s}tr' % CURRENT_NAMESPACES['w'])
                
                if self.row_heights[i] != None:
                    trpr_node = etree.SubElement(tr_node, '{%s}trPr' % CURRENT_NAMESPACES['w'])
                    trheight_node = etree.SubElement(trpr_node, '{%s}trHeight' % CURRENT_NAMESPACES['w'])
                    trheight_node.set('{%s}val' % CURRENT_NAMESPACES['w'], str(self.row_heights[i]))
                    
                for cell in self.cells[i]:
                    cell.render(tr_node)
            

class Cell():
    """
    A cell in a table
    """
    
    def __init__(self, preset_styles, style_name = None,  width = 0, fill = False, **kwargs):
        self.style_name = style_name
        self.styles = preset_styles
        self.parts = []
        self.width = int(width)
        self.fill = fill
        
    def add_paragraph(self, style_name = None, alignment=None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, indent = None, bold=None, italics=None, underlined=None, small_caps=None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Append a paragraph to the cell
        """
        if style_name == None:
            style_name = self.style_name
        
        new_paragraph = Paragraph(preset_styles = self.styles, style_name = style_name, alignment = alignment, border_bottom = border_bottom, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, indent = indent, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_paragraph)
        return new_paragraph
    
    def add_list_point(self, list, style_name = None, level = 0, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a list template to the cell
        """
        if style_name == None:
            style_name = self.style_name
            
        new_list_point = ListPoint(list, preset_styles = self.styles, style_name = style_name, level = level, alignment = alignment, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_list_point)
        return new_list_point
    
    def add_table(self, style_name = None, rows = 1, columns = 1, width = 1, alignment = "left", style = 'default', **kwargs):
        """
        Add a list template to the document
        """
        new_table = Table(preset_styles = self.styles, style_name = style_name, rows = rows, columns = columns, width = width, alignment = alignment, style=style)
        self.parts.append(new_table)
        return new_table
    
    def render(self, root):
        if len(self.parts) < 1 or (self.parts[-1].type != 'paragraph' and self.parts[-1].type != 'listpoint'):
            self.add_paragraph()
            
        CURRENT_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tc_root = etree.SubElement(root, '{%s}tc' % CURRENT_NAMESPACES['w'])
        tcpr_node = etree.SubElement(tc_root, '{%s}tcPr' % CURRENT_NAMESPACES['w'])
        tcw_node = etree.SubElement(tcpr_node, '{%s}tcW' % CURRENT_NAMESPACES['w'])
        tcw_node.set('{%s}w' % CURRENT_NAMESPACES['w'], str(self.width))
        tcw_node.set('{%s}type' % CURRENT_NAMESPACES['w'],"dxa")
        
        if self.fill:
            shd_node = etree.SubElement(tcpr_node, '{%s}shd' % CURRENT_NAMESPACES['w'])
            shd_node.set('{%s}val' % CURRENT_NAMESPACES['w'], 'clear')
            shd_node.set('{%s}col' % CURRENT_NAMESPACES['w'], 'auto')
            shd_node.set('{%s}fill' % CURRENT_NAMESPACES['w'], self.fill)
        
        for part in self.parts:
            part.render(tc_root)
        
        