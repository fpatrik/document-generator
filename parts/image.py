"""
Image support
"""
from lxml import etree
import os
from PIL import Image as Img
from io import BytesIO

class Image():
    """
    Contains data for an image
    """
    
    def __init__(self, image_path, image_no, width = 0.1, **kwargs):
        self.path = image_path
        
        extension = os.path.splitext(self.path)[1][1:].upper()
        if extension == 'JPG' or extension == 'JPEG':
            self.format = 'JPEG'
            
        self.target = 'media/image_'+ image_no + os.path.splitext(self.path)[1]
        self.id = 'rId' + image_no
        self.image = Img.open(self.path)
        self.bytesimage = BytesIO()
        self.image.save(self.bytesimage, format = self.format)
        self.type = 'image'
        
 
        self.width = str(int(7556500 * width))
        self.height = str(int(7556500 * width * self.image.size[1] / self.image.size[0]))
            
    
    def set_width(self, width = 0.1, **kwargs):
        self.width = str(int(7556500 * width))
        self.height = str(int(7556500 * width * self.image.size[1] / self.image.size[0]))
        

    def render(self, root):
        """
        Renders the content of header1.xml
        """
        
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'wp' : 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        r_node = etree.SubElement(root, '{%s}r' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        drawing_node = etree.SubElement(r_node, '{%s}drawing' % CURRENT_NAMESPACES['w'], nsmap=CURRENT_NAMESPACES)
        inline_node = etree.SubElement(drawing_node, '{%s}inline' % CURRENT_NAMESPACES['wp'], nsmap=CURRENT_NAMESPACES)
        inline_node.set('distT', "0")
        inline_node.set('distB', "0")
        inline_node.set('distL', "0")
        inline_node.set('distR', "0")
        extent_node = etree.SubElement(inline_node, '{%s}extent' % CURRENT_NAMESPACES['wp'], nsmap=CURRENT_NAMESPACES)
        extent_node.set('cx', str(self.width))
        extent_node.set('cy', str(self.height))
        docpr_node = etree.SubElement(inline_node, '{%s}docPr' % CURRENT_NAMESPACES['wp'], nsmap=CURRENT_NAMESPACES)
        docpr_node.set('id', "1")
        docpr_node.set('name', "Picture" + self.id)
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'wp' : 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'a' : 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        graphic_node = etree.SubElement(inline_node, '{%s}graphic' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        graphicdata_node = etree.SubElement(graphic_node, '{%s}graphicData' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        graphicdata_node.set('uri', "http://schemas.openxmlformats.org/drawingml/2006/picture")
        CURRENT_NAMESPACES = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'wp' : 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'a' : 'http://schemas.openxmlformats.org/drawingml/2006/main', 'pic' : 'http://schemas.openxmlformats.org/drawingml/2006/picture', 'r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        pic_node = etree.SubElement(graphicdata_node, '{%s}pic' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        nvpicpr_node = etree.SubElement(pic_node, '{%s}nvPicPr' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        cnvpr_node = etree.SubElement(nvpicpr_node, '{%s}cNvPr' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        cnvpr_node.set('id', "0")
        cnvpr_node.set('name', "Picture" + self.id)
        cnvpicpr_node = etree.SubElement(nvpicpr_node, '{%s}cNvPicPr' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        piclocks_node = etree.SubElement(cnvpicpr_node, '{%s}picLocks' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        piclocks_node.set('noChangeAspect', "1")
        piclocks_node.set('noChangeArrowheads', "1")
        blipfill_node = etree.SubElement(pic_node, '{%s}blipFill' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        blip_node = etree.SubElement(blipfill_node, '{%s}blip' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        blip_node.set('{%s}embed' % CURRENT_NAMESPACES['r'], self.id)
        stretch_node = etree.SubElement(blipfill_node, '{%s}stretch' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        fillrect_node = etree.SubElement(stretch_node, '{%s}fillRect' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        sppr_node = etree.SubElement(pic_node, '{%s}spPr' % CURRENT_NAMESPACES['pic'], nsmap=CURRENT_NAMESPACES)
        xfrm_node = etree.SubElement(sppr_node, '{%s}xfrm' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        off_node = etree.SubElement(xfrm_node, '{%s}off' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        off_node.set('x' , "0")
        off_node.set('y' , "0")
        ext_node = etree.SubElement(xfrm_node, '{%s}ext' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        ext_node.set('cx' , str(self.width))
        ext_node.set('cy' , str(self.height))
        prstgeom_node = etree.SubElement(sppr_node, '{%s}prstGeom' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        prstgeom_node.set('prst' , "rect")
        avlst = etree.SubElement(prstgeom_node, '{%s}avLst' % CURRENT_NAMESPACES['a'], nsmap=CURRENT_NAMESPACES)
        
    