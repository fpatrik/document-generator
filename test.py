from conventec_docx.document import Document
def test():
    doc = Document()
    doc.add_style('style 1', font_size = '100')
    
    p = doc.add_paragraph()
    p.add_text('Some test ', bold=True, font_size = "48", font_type = "NSimSun")
    p.add_text('Another test', italics=True)
    
    p = doc.add_paragraph(alignment='center', spacing_line ="2.0")
    p.add_text('Ja ', font_size = "128")
    p.add_text('Sali!', underlined=True)
    
    
    p = doc.add_paragraph(alignment='right', spacing_line="3.5")
    p.add_text('TEST', font_size = "56")
    
    p = doc.add_paragraph(alignment='both')
    p.add_text('Das ist ein weiterer sehr wichtiger Test, der Aufschluss über viele Dinge geben wird. Und es ist genau dieser Aufschluss, der dringend benötigt wird.', font_size = "56")
    
    p = doc.add_paragraph(style_name = 'style 1')
    p.add_text('It works!')
    
    p = doc.add_paragraph()
    p.add_text('small caps!', small_caps = True)
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
def test2():
    doc = Document()
    
    doc.add_style('style 1', font_size = '20', italics = True)
    
    p = doc.add_paragraph(spacing_after="600")
    p.add_text("Sali!")
    list_template_1 = doc.add_list_template(type = "roman")
    list_1 = doc.add_list(list_template_1, style_name = 'style 1')
    point_1 = doc.add_list_point(list_1, spacing_after="500")
    point_1.add_text('Sali!', font_size = "12")
    point_1.add_text('Sali!', font_size = "12", bold= True)
    point_2 = doc.add_list_point(list_1, style_name = 'style 1')
    point_2.add_text('Du!', font_size = "12")
    point_2 = doc.add_list_point(list_1, level = 1)
    point_2.add_text('Level up!', font_size = "24", bold = True)
    point_2.add_text('Test')
    
    doc.add_paragraph(alignment='center')
    p = doc.add_paragraph(alignment='left')
    p.add_text('Normal text')
    doc.add_paragraph(alignment='center')
    
    list_template_2 = doc.add_list_template(type = "bullet")
    list_2 = doc.add_list(list_template_2)
    point_1 = doc.add_list_point(list_2)
    point_1.add_text('Morge!', font_size = "56")
    point_2 = doc.add_list_point(list_2)
    point_2.add_text('Mitenand!', font_size = "56")
    
    doc.add_paragraph(alignment='center')
    doc.add_paragraph(alignment='center')
    
    normal_list_template = doc.add_list_template(type = "numbering", indent = 1)
    normal_list = doc.add_list(normal_list_template)
    
    list_in_list_template = doc.add_list_template(type = "bullet", indent = 2)
    list_in_list = doc.add_list(list_in_list_template)
    
    p1 = doc.add_list_point(normal_list)
    p1.add_text('Bla', font_size = "24", text_color = 'red')
    p2 = doc.add_list_point(normal_list)
    p2.add_text('Bli', font_size = "24", highlight_color = 'blue')
    p3 = doc.add_list_point(list_in_list)
    p3.add_text('Blu', font_size = "24")
    p4 = doc.add_list_point(list_in_list)
    p4.add_text('Ble', font_size = "24")
    p5 = doc.add_list_point(normal_list)
    p5.add_text('Blabli', font_size = "24")

    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
    
def test3():
    doc = Document(landscape = True)
    
    doc.add_style('style 1', font_size = '20', italics = True)
    
    table = doc.add_table(rows = 2, columns = 2, width = 0.5, alignment = "center", style_name = 'style 1')
    
    table.cells[0][0].fill = "F2F2F2"
    
    list_template_1 = doc.add_list_template(type = "roman")
    list_1 = doc.add_list(list_template_1, style_name = 'style 1')
    
    p1 = table.cells[0][0].add_list_point(list_1, spacing_after="500")
    p1.add_text('Sali!', font_size = "12")
    
    p1 = table.cells[0][1].add_paragraph()
    p1.add_text('Bla vla bla, hier ist voll viel Text.', vertical_align="superscript")
    
    doc.add_paragraph()
    
    table2 = doc.add_table(rows = 2, columns = 2, width = 0.75, alignment = "center", style = "borderless")
    
    p1 = table2.cells[0][0].add_paragraph()
    p1.add_text('Table 2', font_size = "24")
    
    list_template = doc.add_list_template(type = "roman")
    list_template.type = 'list'
    list_1 = doc.add_list(list_template)
    
    point = doc.add_list_point(list_1, level = 0)
    point.add_text('This point also belongs to list 1, but it is indented!')
    point = doc.add_list_point(list_1, level = 1)
    point.add_text('This point also belongs to list 1, but it is indented!')
    point = doc.add_list_point(list_1, level = 2)
    point.add_text('This point also belongs to list 1, but it is indented!')
    
    
    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
def test4():
    doc = Document()
    
    image1 = doc.import_image('c:\\Users\\User\\Desktop\\conventec.jpeg')
    header = doc.add_header()
    p1 = header.default.add_paragraph(alignment = 'right')
    p1.use_image(image1)
    image1.set_width(width = 0.1)
    
    p3 = header.first.add_paragraph(alignment = 'left')
    p3.add_text('Some text', font_size = "24")
    
    footer = doc.add_footer()
    p4 = footer.default.add_paragraph(alignment = 'center')
    p4.add_simplefield('page', font_size = "24", font_type = "Comic Sans MS")
    
    doc.title_page = False
    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
def test5():
    doc = Document()
    p1 = doc.add_paragraph(keep_next = True)
    p1.add_text('Title', font_size = "32", bold = True)
    
    p2 = doc.add_paragraph()
    p2.add_text('Text', font_size = "24")
    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
    
def test6():
    doc = Document()
    p1 = doc.add_paragraph()
    p1.add_line_break()
    p1.add_text('Test', font_size = "32")
    p1.add_line_break(n=7)
    p1.add_text('This Page', font_size = "32")
    p1.add_page_break()
    p1.add_text('Next Page', font_size = "32")
    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
def test7():
    doc = Document()
    
    title_template_1 = doc.add_numbered_title_template(text = '', style = 'roman' , separator = '.')
    titles_1 = doc.add_numbered_title(title_template_1)
    
    title_template_2 = doc.add_numbered_title_template(text = 'Artikel ', style = 'numbering' , separator = '.')
    titles_2 = doc.add_numbered_title(title_template_2)
    
    title1 = doc.add_title(titles_1, alignment = "center", font_size = "72")
    title1.add_text("Allgemeines")
    
    title2 = doc.add_title(titles_2, bold = True, font_size = "72")
    title2.add_text("Bla Bla Bla")
    
    p1 = doc.add_paragraph()
    p1.add_text("Ein bisschen Text. Das Spezifische ist unter ")
    reference1 = p1.add_reference()
    p1.add_text(" zu finden.")
    
    title3 = doc.add_title(titles_2)
    title3.add_text("Bli Bla Blu")
    
    p2 = doc.add_paragraph()
    p2.add_text("Und noch ein bisschen Text! ")
    p2.add_reference(title2)
    p2.add_text(" finde ich der beste!")
    
    title4 = doc.add_title(titles_1, alignment = "center")
    title4.add_text("Spezifisches")
    reference1.title = title4
    
    
    
    title5 = doc.add_title(titles_2)
    title5.add_text("Details")
    
    p3 = doc.add_paragraph()
    p3.add_text("Das sind die Details...")
    
    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
    
def test8():
    doc = Document()
    p1 = doc.add_paragraph(border_bottom = True)
    p1.add_text('Horizontal Line!', font_size = "32")

    
    doc.save('c:\\Users\\User\\Desktop\\test.docx')
