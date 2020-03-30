"""
Contains the Document class which is an entire document
"""

from lxml import etree
from conventec_docx.writing.docx import DocxFile
from conventec_docx.parts.list import ListTemplate, List, ListPoint
from conventec_docx.parts.paragraph import Paragraph
from conventec_docx.parts.table import Table, Cell
from conventec_docx.parts.image import Image
from conventec_docx.parts.header import Header
from conventec_docx.parts.footer import Footer
from conventec_docx.parts.numbering import NumberedTitleTemplate, NumberedTitle, Title
from conventec_docx.parts.style import Style

class Document():
    """
    Contains the entire docx document
    """
    def __init__(self, title_page = False, landscape = False, auto_hyphenate = True, **kwargs):
        """
        Initialises the part of the document
        """
        self.parts = []
        self.header = None
        self.footer = None
        self.list_templates = []
        self.images = []
        self.styles = {'default' : Style(self, 'default', alignment="left", border_bottom = False, keep_next = False, spacing_before = "0", spacing_after = "0", spacing_line = "1", indent = "0", bold=False, italics=False, underlined=False, small_caps=False, font_type = "Arial", font_size="12", text_color = None, highlight_color = None, vertical_align = None)}
        self.set_default_style(**kwargs)
        self.styles['conventec_default'] = self.styles['default']
        self.n_of_lists = 0
        self.title_page = title_page
        self.landscape = landscape
        self.auto_hyphenate = auto_hyphenate
    
    def set_default_style(self, **kwargs):
        """
        Sets default style of document
        """
        for key, value in kwargs.items():
            setattr(self.styles['default'], key, value)
    
    def add_paragraph(self, style_name = None, alignment=None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, indent = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Append a paragraph to the document
        """
        new_paragraph = Paragraph(preset_styles = self.styles, style_name = style_name, alignment = alignment, border_bottom = border_bottom, keep_next = keep_next, indent = indent, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_paragraph)
        return new_paragraph
    
    def add_list_template(self, indent = "1", type = "numbering", **kwargs):
        """
        Add a list template to the document
        """
        new_list = ListTemplate(len(self.list_templates), indent = indent, type = type)
        self.list_templates.append(new_list)
        return new_list
    
    def add_list(self, list_template, style_name = None, **kwargs):
        """
        Add a list template to the document
        """
        self.n_of_lists += 1
        return List(list_template, self.n_of_lists, style_name = style_name)
    
    
    def add_list_point(self, list, style_name = None, level = 0, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a list template to the document
        """
        new_list_point = ListPoint(list, preset_styles = self.styles, style_name = style_name, level = level, alignment = alignment, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_list_point)
        return new_list_point
    
    def add_numbered_title_template(self, indent = "1", text = '', style = 'numbering', separator = '.', **kwargs):
        """
        Add a numbered title template
        """
        new_title = NumberedTitleTemplate(len(self.list_templates), text = text, style = style, separator = separator)
        self.list_templates.append(new_title)
        return new_title
    
    def add_numbered_title(self, numbered_title_template, style_name = None, **kwargs):
        """
        Add a list template to the document
        """
        self.n_of_lists += 1
        return NumberedTitle(numbered_title_template, self.n_of_lists, self.styles, style_name = style_name)
    
    
    def add_title(self, numbered_title, style_name = None, alignment = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type = None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Add a new title to the document
        """
        new_title = Title(numbered_title, self.styles, style_name = style_name, alignment = alignment, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.parts.append(new_title)
        return new_title
    
    def add_table(self, style_name = None, rows = 1, columns = 1, width = 1, alignment = "left", style = 'default', delete_empty = False, **kwargs):
        """
        Add a list template to the document
        """
        new_table = Table(self.styles, style_name = style_name, rows = rows, columns = columns, width = width, alignment = alignment, style=style, delete_empty = delete_empty)
        self.parts.append(new_table)
        return new_table
    
    def import_image(self, path, **kwargs):
        """
        Import an image to the document
        """
        new_image = Image(path, str(20 + len(self.images)))
        self.images.append(new_image)
        return new_image
    
    def add_header(self, style_name = None, **kwargs):
        """
        Adds a header to the document
        """
        new_header = Header(preset_styles = self.styles, style_name = style_name)
        self.header = new_header
        return new_header
    
    def add_footer(self, style_name = None, style = 'default', **kwargs):
        """
        Adds a footer to the document
        """
        new_footer = Footer(preset_styles = self.styles, style_name = style_name)
        self.footer = new_footer
        return new_footer
    
    def add_style(self, reference_name, alignment=None, border_bottom = None, keep_next = None, spacing_before = None, spacing_after = None, spacing_line = None, bold=None, italics=None, underlined=None, small_caps = None, font_type =None, font_size=None, text_color = None, highlight_color = None, vertical_align = None, **kwargs):
        """
        Adds a new style to the document
        """
        new_style = Style(self, reference_name, alignment=alignment, border_bottom = border_bottom, keep_next = keep_next, spacing_before = spacing_before, spacing_after = spacing_after, spacing_line = spacing_line, bold=bold, italics=italics, underlined=underlined, small_caps = small_caps, font_type = font_type, font_size=font_size, text_color = text_color, highlight_color = highlight_color, vertical_align = vertical_align)
        self.styles[reference_name] = new_style
        return new_style
    
    def save(self, path, **kwargs):
        """
        Saves the document to a given path
        """
        output = DocxFile()
        output.write('[Content_Types].xml', self.render_content_types())
        output.write('_rels/.rels', self.render_main_rels())
        output.write('docProps/app.xml', self.render_app())
        output.write('docProps/core.xml', self.render_core())
        output.write('word/_rels/document.xml.rels', self.render_word_rels())
        output.write('word/fontTable.xml', self.render_font_table())
        output.write('word/settings.xml', self.render_settings())
        output.write('word/styles.xml', self.render_styles())
        output.write('word/webSettings.xml', self.render_web_settings())
        
        if len(self.list_templates) > 0:
            output.write('word/numbering.xml', self.render_numbering())
        if self.header is not None:
            if len(self.header.even.parts) > 0:
                output.write('word/_rels/header1.xml.rels', self.render_header1_rels())
                output.write('word/header1.xml', self.render_header1())
            if len(self.header.default.parts) > 0:
                output.write('word/_rels/header2.xml.rels', self.render_header2_rels())
                output.write('word/header2.xml', self.render_header2())
            if len(self.header.first.parts) > 0:
                output.write('word/_rels/header3.xml.rels', self.render_header3_rels())
                output.write('word/header3.xml', self.render_header3())
                
        if self.footer is not None:
            if len(self.footer.even.parts) > 0:
                output.write('word/_rels/footer1.xml.rels', self.render_footer1_rels())
                output.write('word/footer1.xml', self.render_footer1())
            if len(self.footer.default.parts) > 0:
                output.write('word/_rels/footer2.xml.rels', self.render_footer2_rels())
                output.write('word/footer2.xml', self.render_footer2())
            if len(self.footer.first.parts) > 0:
                output.write('word/_rels/footer3.xml.rels', self.render_footer3_rels())
                output.write('word/footer3.xml', self.render_footer3())
        
        for image in self.images:
            output.write('word/' + image.target ,image.bytesimage.getvalue())
            
        output.write('word/document.xml', self.render_document())
        
        output.save(path)
        output.close()
        
    
    def download(self, **kwargs):
        """
        Returns the document for download
        """
        output = DocxFile()
        output.write('[Content_Types].xml', self.render_content_types())
        output.write('_rels/.rels', self.render_main_rels())
        output.write('docProps/app.xml', self.render_app())
        output.write('docProps/core.xml', self.render_core())
        output.write('word/_rels/document.xml.rels', self.render_word_rels())
        output.write('word/fontTable.xml', self.render_font_table())
        output.write('word/settings.xml', self.render_settings())
        output.write('word/styles.xml', self.render_styles())
        output.write('word/webSettings.xml', self.render_web_settings())
        
        if len(self.list_templates) > 0:
            output.write('word/numbering.xml', self.render_numbering())
        if self.header is not None:
            if len(self.header.even.parts) > 0:
                output.write('word/_rels/header1.xml.rels', self.render_header1_rels())
                output.write('word/header1.xml', self.render_header1())
            if len(self.header.default.parts) > 0:
                output.write('word/_rels/header2.xml.rels', self.render_header2_rels())
                output.write('word/header2.xml', self.render_header2())
            if len(self.header.first.parts) > 0:
                output.write('word/_rels/header3.xml.rels', self.render_header3_rels())
                output.write('word/header3.xml', self.render_header3())
                
        if self.footer is not None:
            if len(self.footer.even.parts) > 0:
                output.write('word/_rels/footer1.xml.rels', self.render_footer1_rels())
                output.write('word/footer1.xml', self.render_footer1())
            if len(self.footer.default.parts) > 0:
                output.write('word/_rels/footer2.xml.rels', self.render_footer2_rels())
                output.write('word/footer2.xml', self.render_footer2())
            if len(self.footer.first.parts) > 0:
                output.write('word/_rels/footer3.xml.rels', self.render_footer3_rels())
                output.write('word/footer3.xml', self.render_footer3())
        
        for image in self.images:
            output.write('word/' + image.target ,image.bytesimage.getvalue())
            
        output.write('word/document.xml', self.render_document())
        
        return output.download()
        
        
    def render_document(self):
        """
        Renders 'word/document.xml'
        """
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        document_root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:document>""".encode('utf-8'))
        
        body_root = etree.SubElement(document_root, '{%s}body' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        body_root.text=''
        for part in self.parts:
            part.render(body_root)
        
        sectpr_root = etree.SubElement(body_root, '{%s}sectPr' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        if self.header is not None:
            if len(self.header.even.parts) > 0:
                headerreference = etree.SubElement(sectpr_root, '{%s}headerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                headerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'even')
                headerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId6")
            if len(self.header.default.parts) > 0:
                headerreference = etree.SubElement(sectpr_root, '{%s}headerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                headerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'default')
                headerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId7")
            if len(self.header.first.parts) > 0:
                headerreference = etree.SubElement(sectpr_root, '{%s}headerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                headerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'first')
                headerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId8")
                
        if self.footer is not None:
            if len(self.footer.even.parts) > 0:
                footerreference = etree.SubElement(sectpr_root, '{%s}footerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                footerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'even')
                footerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId9")
            if len(self.footer.default.parts) > 0:
                footerreference = etree.SubElement(sectpr_root, '{%s}footerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                footerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'default')
                footerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId10")
            if len(self.footer.first.parts) > 0:
                footerreference = etree.SubElement(sectpr_root, '{%s}footerReference' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                footerreference.set('{%s}type' % CURRENT_NAMESPACES['w'], 'first')
                footerreference.set('{%s}id' % CURRENT_NAMESPACES['r'],"rId11")
                
        pgsz = etree.SubElement(sectpr_root, '{%s}pgSz' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        if self.landscape:
            pgsz.set('{%s}orient' % CURRENT_NAMESPACES['w'],"landscape")
            pgsz.set('{%s}w' % CURRENT_NAMESPACES['w'],"16838")
            pgsz.set('{%s}h' % CURRENT_NAMESPACES['w'],"11906")
        else:
            pgsz.set('{%s}w' % CURRENT_NAMESPACES['w'],"11906")
            pgsz.set('{%s}h' % CURRENT_NAMESPACES['w'],"16838")
        pgmar = etree.SubElement(sectpr_root, '{%s}pgMar' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        pgmar.set('{%s}top' % CURRENT_NAMESPACES['w'],"1417")
        pgmar.set('{%s}bottom' % CURRENT_NAMESPACES['w'],"1134")
        pgmar.set('{%s}right' % CURRENT_NAMESPACES['w'],"1417")
        pgmar.set('{%s}left' % CURRENT_NAMESPACES['w'],"1417")
        pgmar.set('{%s}header' % CURRENT_NAMESPACES['w'],"708")
        pgmar.set('{%s}footer' % CURRENT_NAMESPACES['w'],"708")
        pgmar.set('{%s}gutter' % CURRENT_NAMESPACES['w'],"0")
        cols = etree.SubElement(sectpr_root, '{%s}cols' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        cols.set('{%s}space' % CURRENT_NAMESPACES['w'],"708")
        
        if self.title_page == True:
            etree.SubElement(sectpr_root, '{%s}titlePg' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        
        docgrid = etree.SubElement(sectpr_root, '{%s}docGrid' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        docgrid.set('{%s}linePitch' % CURRENT_NAMESPACES['w'],"360")
        
        return etree.tostring(document_root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_content_types(self):
        """
        Renders '[Content_types].xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>""".encode('utf-8'))
        etree.SubElement(root, "Default", Extension = 'rels', ContentType="application/vnd.openxmlformats-package.relationships+xml")
        etree.SubElement(root, "Default", Extension = 'xml', ContentType="application/xml")
        etree.SubElement(root, "Default", Extension = 'jpeg', ContentType="image/jpeg")
        etree.SubElement(root, "Override", PartName="/word/document.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
        etree.SubElement(root, "Override", PartName="/word/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
        etree.SubElement(root, "Override", PartName="/word/settings.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
        etree.SubElement(root, "Override", PartName="/word/webSettings.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml")
        etree.SubElement(root, "Override", PartName="/word/fontTable.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")
        etree.SubElement(root, "Override", PartName="/docProps/core.xml", ContentType="application/vnd.openxmlformats-package.core-properties+xml")
        etree.SubElement(root, "Override", PartName="/docProps/app.xml", ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml")
        etree.SubElement(root, "Override", PartName="/word/numbering.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
        if self.header != None:
            if len(self.header.even.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/header1.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
            if len(self.header.default.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/header2.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
            if len(self.header.first.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/header3.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
        
        if self.footer != None:
            if len(self.footer.even.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/footer1.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
            if len(self.footer.default.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/footer2.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
            if len(self.footer.first.parts) > 0:
                etree.SubElement(root, "Override", PartName="/word/footer3.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
                
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_main_rels(self):
        """
        Renders '_rels/.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        etree.SubElement(root, "Relationship", Id="rId3", Target="docProps/app.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
        etree.SubElement(root, "Relationship", Id="rId2", Target="docProps/core.xml", Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
        etree.SubElement(root, "Relationship", Id="rId1", Target="word/document.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        

                
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_app(self):
        """
        Renders 'docProps/app.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>""".encode('utf-8'))
        etree.SubElement(root, "Template").text = "Normal.dotm"
        etree.SubElement(root, "TotalTime").text = "0"
        etree.SubElement(root, "Pages").text = "18"
        etree.SubElement(root, "Words").text = "0"
        etree.SubElement(root, "Characters").text = "0"
        etree.SubElement(root, "Application").text = "Microsoft Office Word"
        etree.SubElement(root, "DocSecurity").text = "0"
        etree.SubElement(root, "Lines").text = "0"
        etree.SubElement(root, "Paragraphs").text = "0"
        etree.SubElement(root, "ScaleCrop").text = "false"
        etree.SubElement(root, "Company").text = ""
        etree.SubElement(root, "LinksUpToDate").text = "false"
        etree.SubElement(root, "CharactersWithSpaces").text = "0"
        etree.SubElement(root, "SharedDoc").text = "false"
        etree.SubElement(root, "HyperlinksChanged").text = "false"
        etree.SubElement(root, "AppVersion").text = "16.0000"
    
        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")

    def render_core(self):
        """
        Renders 'docProps/core.xml'
        """
        
        CURRENT_NAMESPACES = {'cp' : 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'dc' : 'http://purl.org/dc/elements/1.1/', 'dcterms' : 'http://purl.org/dc/terms/', 'dcmitype' : 'http://purl.org/dc/dcmitype/', 'xsi' : 'http://www.w3.org/2001/XMLSchema-instance'}
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>""".encode('utf-8'))
        etree.SubElement(root, '{%s}title' % CURRENT_NAMESPACES['dc'], nsmap=CURRENT_NAMESPACES).text=''
        etree.SubElement(root, '{%s}subject' % CURRENT_NAMESPACES['dc'], nsmap=CURRENT_NAMESPACES).text=''
        etree.SubElement(root, '{%s}creator' % CURRENT_NAMESPACES['dc'], nsmap=CURRENT_NAMESPACES).text = "Conventec"
        etree.SubElement(root, '{%s}keywords' % CURRENT_NAMESPACES['cp'], nsmap=CURRENT_NAMESPACES).text=''
        etree.SubElement(root, '{%s}description' % CURRENT_NAMESPACES['dc'], nsmap=CURRENT_NAMESPACES).text=''
        etree.SubElement(root, '{%s}lastModifiedBy' % CURRENT_NAMESPACES['cp'], nsmap=CURRENT_NAMESPACES).text='Conventec'
        etree.SubElement(root, '{%s}revision' % CURRENT_NAMESPACES['cp'], nsmap=CURRENT_NAMESPACES).text='1'
        created = etree.SubElement(root, '{%s}created' % CURRENT_NAMESPACES['dcterms'], nsmap=CURRENT_NAMESPACES)
        created.set('{%s}type' % CURRENT_NAMESPACES['xsi'],"dcterms:W3CDTF")
        created.text = "2017-08-28T11:54:00Z"
        modified = etree.SubElement(root, '{%s}modified' % CURRENT_NAMESPACES['dcterms'], nsmap=CURRENT_NAMESPACES)
        modified.set('{%s}type' % CURRENT_NAMESPACES['xsi'],"dcterms:W3CDTF")
        modified.text = "2017-08-28T11:54:00Z"
    
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_word_rels(self):
        """
        Renders 'word/rels/document.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        etree.SubElement(root, "Relationship", Id="rId3", Target="webSettings.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings")
        etree.SubElement(root, "Relationship", Id="rId2", Target="settings.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings")
        etree.SubElement(root, "Relationship", Id="rId1", Target="styles.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        etree.SubElement(root, "Relationship", Id="rId4", Target="fontTable.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable")
        etree.SubElement(root, "Relationship", Id="rId5", Target="numbering.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
        
        if self.header is not None:
            if len(self.header.even.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId6", Target="header1.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
            if len(self.header.default.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId7", Target="header2.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
            if len(self.header.first.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId8", Target="header3.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
                
        if self.footer is not None:
            if len(self.footer.even.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId9", Target="footer1.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
            if len(self.footer.default.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId10", Target="footer2.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
            if len(self.footer.first.parts) > 0:
                etree.SubElement(root, "Relationship", Id="rId11", Target="footer3.xml", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
            
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header1_rels(self):
        """
        Renders 'word/rels/header1.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header2_rels(self):
        """
        Renders 'word/rels/header2.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header3_rels(self):
        """
        Renders 'word/rels/header3.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    
    def render_footer1_rels(self):
        """
        Renders 'word/rels/footer1.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_footer2_rels(self):
        """
        Renders 'word/rels/footer2.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_footer3_rels(self):
        """
        Renders 'word/rels/footer3.xml.rels'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""".encode('utf-8'))
        for image in self.images:
            etree.SubElement(root, "Relationship", Id = image.id, Target = image.target , Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header1(self):
        """
        Renders 'word/header1.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:hdr>""".encode('utf-8'))
        if self.header is not None:
            self.header.even.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header2(self):
        """
        Renders 'word/header2.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:hdr>""".encode('utf-8'))
        if self.header is not None:
            self.header.default.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_header3(self):
        """
        Renders 'word/header3.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:hdr>""".encode('utf-8'))
        if self.header is not None:
            self.header.first.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    
    def render_footer1(self):
        """
        Renders 'word/footer1.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:ftr>""".encode('utf-8'))
        if self.footer is not None:
            self.footer.even.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_footer2(self):
        """
        Renders 'word/footer2.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:ftr>""".encode('utf-8'))
        if self.footer is not None:
            self.footer.default.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_footer3(self):
        """
        Renders 'word/footer3.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:ftr>""".encode('utf-8'))
        if self.footer is not None:
            self.footer.first.render(root)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_font_table(self):
        """
        Renders 'word/fontTable.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid"></w:fonts>""".encode('utf-8'))
        root.text=''
        
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_settings(self):
        """
        Renders 'word/settings.xml'
        """
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15 w16se w16cid"></w:settings>""".encode('utf-8'))
        if self.header != None:
            if len(self.header.even.parts) > 0:
                etree.SubElement(root, '{%s}evenAndOddHeaders' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                
        if self.footer != None:
            if len(self.footer.even.parts) > 0:
                etree.SubElement(root, '{%s}evenAndOddFooters' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
                
        if self.auto_hyphenate:
            etree.SubElement(root, '{%s}autoHyphenation' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            hyphenation_zone = etree.SubElement(root, '{%s}hyphenationZone' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
            hyphenation_zone.set('{%s}val' % CURRENT_NAMESPACES['w'], "425")
                
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_styles(self):
        """
        Renders 'word/styles.xml'
        """
  
        root = etree.fromstring(("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid">
        <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph" />
    <w:basedOn w:val="Normal" />
    <w:uiPriority w:val="34" />
    <w:qFormat />
    <w:rsid w:val="007A5E92" />
    <w:pPr>
      <w:ind w:left="720" />
      <w:contextualSpacing />
    </w:pPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table" />
    <w:uiPriority w:val="99" />
    <w:semiHidden />
    <w:unhideWhenUsed />
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa" />
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa" />
        <w:left w:w="108" w:type="dxa" />
        <w:bottom w:w="0" w:type="dxa" />
        <w:right w:w="108" w:type="dxa" />
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid" />
    <w:basedOn w:val="TableNormal" />
    <w:uiPriority w:val="39" />
    <w:rsid w:val="00B60D5C" />
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto" />
    </w:pPr>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto" />
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto" />
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto" />
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto" />
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto" />
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto" />
      </w:tblBorders>
    </w:tblPr>
  </w:style>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal" />
    <w:qFormat />
  </w:style>
    <w:style w:type="paragraph" w:styleId="Header">
    <w:name w:val="header" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="HeaderChar" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="000C2F70" />
    <w:pPr>
      <w:tabs>
        <w:tab w:val="center" w:pos="4536" />
        <w:tab w:val="right" w:pos="9072" />
      </w:tabs>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto" />
    </w:pPr>
  </w:style>
  </w:styles>""").encode('utf-8'))
        root.text=''
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
    
    def render_web_settings(self):
        """
        Renders 'word/webSettings.xml'
        """
        
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid"></w:webSettings>""".encode('utf-8'))
        root.text=''
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
        
        
    def render_numbering(self):
        """
        Render 'word/numbering.xml'
        """   
        root = etree.XML("""<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"></w:numbering>""".encode('utf-8'))
        for template in self.list_templates:
            root.append(template.render())
            
        for template in self.list_templates:
            for link in template.numlinks:
                root.append(link)
            
        return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes").decode("utf-8")
        
    
        