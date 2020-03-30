"""
CONVENTEC DOCX TUTORIAL

Conventec Docx is a Python package that can generate docx-files.
This file conatins examples demonstrating the capabilities of Conventec Docx.
"""

"""
EXAMPLE 1 - CREATING DOCUMENTS

The first step for using Conventec Docx is importing the Document class from the package.
Then an instance of this class can be created.
Finally, the document can be saved using the save method.
"""

# Import Document class
from conventec_docx.document import Document

# Set path for saving documents:
save_path = 'c:\\Users\\User\\Desktop\\'

def example_1():
    """
    First an instance of the Document class is created, it contains all data of the document.
    Document takes the following arguments:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    title_page        If the document should contain a title   False            Optional
                      page. Shows different headers and
                      footers in case they are set.
                      
    landscape         If the document should be in landscape   False            Optional
                      format.
                      
    """
    doc = Document(title_page = False, landscape = False)
    
    """
    The document can now be saved with the SAVE METHOD. For the moment the document is 
    saved to the hard drive. Later the document could be served directly from the memory
    to the client. The save method takes these arguments:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    path              If the document should contain a title                    Required
                      page. Shows different headers and
                      footers in case they are set.
    
    
    """
    doc.save(save_path + 'example_1.docx')
    
    """
    This should create an empty document.
    """


"""
EXAMPLE 2 - ADDING PARAGRAPHS AND TEXT

In order to add text to document, you have to add a paragraph first which can then be 
populated with text.
"""

def example_2():
    """
    First we create an empty document:
    """
    
    doc = Document()
    
    """
    A paragraph can be added with the ADD_PARAGRAPH METHOD of the Document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    alignment         The alignment of the paragraph.          'left'           Optional
                      'left', 'right', 'center' and
                      'both' is supported.
                      
    border_bottom     If the Paragraph should have a solid     False            Optional
                      border at the bottom. Useful for
                      creating horizontal lines or 
                      signature fields.
                      
    keep_next         If the paragraph following the           False            Optional
                      current should be on the same
                      page. Useful for preventing page
                      breaks between title and text.
                      
    spacing_before    Adds a padding before the paragraph      None             Optional
                      120 might be an appropriate value.
                      
    spacing_after     Adds a padding after the paragraph       None             Optional
                      120 might be an appropriate value.
                      
    spacing_line      Spacing between lines. Takes a float.    '1.0'            Optional
                     
    indent            The indentation of the paragraph         '0'              Optional
    
    bold              If the paragraph style should be         False            Optional
                      bold.
        
    italics           If the paragraph style should be         False            Optional
                      italic.
    
    underlined        If the paragraph style should be         False            Optional
                      underlined.
                      
    small_caps        If the paragraph style should be         False            Optional
                      in small caps.
                      
    font_type         Paragraph font. These are probably       'Arial'          Optional
                      supported:
                      https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
                      
    font_size         Paragraph font size.                     '24'             Optional
                      Use double the desired size.
                      e.g. (12pt = '24')
                      
    text_color        Paragraph text color.                    None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      RGB support could be implemented.
                      
    highlight_color   Paragraph highlight color.               None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      
    vertical_align    Vertical alignment of text.              None             Optional
                      e.g. 'superscript'

    """
    
    p = doc.add_paragraph(bold = False)
    
    """
    All above paragraph properties can still be changed later:

    """
    
    p.bold = True
    p.alignment = 'center'
    
    """
    Now Text can be added to the paragraph with the ADD_TEXT method:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    text              The text to be added to the                               Required
                      paragraph.
                      
    bold              If the text style should be              False            Optional
                      bold.                                   
        
    italics           If the text style should be              False            Optional
                      italic.
    
    underlined        If the text style should be              False            Optional
                      underlined.
                      
    small_caps        If the paragraph style should be         False            Optional
                      in small caps.
                      
    font_type         Text font. These are probably            'Arial'          Optional
                      supported:
                      https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
                      
    font_size         Text font size. (12pt = '24')            '24'             Optional
                      Use double the desired size.
                      e.g. (12pt = '24')
                      
    text_color        Text color.                              None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      RGB support could be implemented.
                      
    highlight_color   Text highlight color.                    None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      
    vertical_align    Vertical alignment of text.              None             Optional
                      e.g. 'superscript'
    """
    
    p.add_text('This is a text.')
    
    """
    Multiple texts can be added to a paragraph:
    """
    
    p.add_text(' This text is highlighted!', highlight_color = 'yellow')
    
    """
    Text attributes can also all be changed after creation.
    """
    
    t = p.add_text(' This text is red.', text_color = 'red')
    
    t.text_color = 'blue'
    t.text = ' But now it is blue!!!'
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_2.docx')
    
    """
    But why did we set the paragraph style if the text has its own style anyway?
    The paragraph style is used for line breaks and also appears at the end
    of the paragraph.
    """
    
    
"""
EXAMPLE 3 - CREATING LISTS

Lists are technically a paragraph within a document. They can be created in three steps:
1. Create a list template determining the style of the list
2. Create a list of that template
3. Create paragraphs (listpoints) of that list
"""

def example_3():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    Now we can use the ADD_LIST_TEMPLATE method of the document to create a list template:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    indent            Determines the indentation level         '1'              Optional
                      of the list.
                      
    type              Determines the the type of the list.     'numbering'      Optional
                      Currently 'numbering', 'bullet',
                      'roman', 'list' and 'letter' 
                      are supported.
    """
    
    list_template = doc.add_list_template(type = "roman")
    
    """
    Again, we can change the attributes later if we want:
    """
    
    list_template.type = 'numbering'
    
    """
    To create a list, we use the ADD_LIST method of the document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    list_template     The list templat that should be                           Required
                      used for the list.
    """
    
    list_1 = doc.add_list(list_template)
    
    """
    The template used for the list can no longer be changed later. Also, the template can
    ONLY BE USED FOR ONE LIST. This is beacuse Word has a meltdown otherwise.
    can add points of the lists to the document using the ADD_LIST_POINT method of the document. 
    Remember that list points are technically paragraphs, therefore the arguments are 
    almost the same.
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    list              The list the point belongs to.                            Reqiured
    
    level             The indentation level of the point       0                Optional
                      within the list.
    
    alignment         The alignment of the paragraph.          'left'           Optional
                      'left', 'right', 'center' and
                      'both' is supported.
                      
    keep_next         If the paragraph following the           False            Optional
                      current should be on the same
                      page. Useful for preventing page
                      breaks between title and text.
                
    spacing_before    Adds a padding before the paragraph      None             Optional
                      120 might be an appropriate value.
                      
    spacing_after     Adds a padding after the paragraph       None             Optional
                      120 might be an appropriate value.
                      
    spacing_line      Spacing between lines. Takes a float.    '1.0'            Optional
                      
    indent            The indentation of the paragraph         '0'              Optional
    
    bold              If the paragraph style should be         False            Optional
                      bold.
        
    italics           If the paragraph style should be         False            Optional
                      italic.
    
    underlined        If the paragraph style should be         False            Optional
                      underlined.
                      
    small_caps        If the paragraph style should be         False            Optional
                      in small caps.
                      
    font_type         Paragraph font. These are probably       'Arial'          Optional
                      supported:
                      https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
                      
    font_size         Paragraph font size.                     '24'             Optional
                      Use double the desired size.
                      e.g. (12pt = '24')
                      
    text_color        Paragraph text color.                    None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      RGB support could be implemented.
                      
    highlight_color   Paragraph highlight color.               None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      
    vertical_align    Vertical alignment of text.              None             Optional
                      e.g. 'superscript'
    """
    
    point = doc.add_list_point(list_1)
    
    """
    Since list points are paragraphs, the ADD_TEXT method of list points is identical to the
    one of paragraphs.
    """
    
    point.add_text('This point belongs to list 1.')
    
    point = doc.add_list_point(list_1)
    point.add_text('This point also belongs to list 1.')
    
    point = doc.add_list_point(list_1, level = 1)
    point.add_text('This point also belongs to list 1, but it is indented!')
    
    point = doc.add_list_point(list_1, level = 2)
    point.add_text('This point also belongs to list 1 and it is even more indented!!!')
    
    """
    Let us try to do a nested list with numbering on the outer level and bullet points on 
    the inner. 
    """
    
    #Add paragraph for visibility
    doc.add_paragraph()
    
    #Create the two templates
    outer_template = doc.add_list_template(type = "numbering", indent = '1')
    inner_template = doc.add_list_template(type = "bullet", indent = '2')
    
    #Create lists
    outer_list = doc.add_list(outer_template)
    inner_list = doc.add_list(inner_template)
    
    #Add points
    point = doc.add_list_point(outer_list)
    point.add_text('This is the first point.')
    
    point = doc.add_list_point(inner_list)
    point.add_text('This is a bullet.')
    
    point = doc.add_list_point(inner_list)
    point.add_text('And this as well.')
    
    point = doc.add_list_point(outer_list)
    point.add_text('This is the second point.')
    
    point = doc.add_list_point(inner_list)
    point.add_text('This is again a bullet.')
    
    
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_3.docx')
    

"""
EXAMPLE 4 - CREATING TABLES

Tables can be contained within the document as well as within other tables. They themselves
can contain paragraphs, lists and more tables. 
"""

def example_4():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    A table is created with the ADD_TABLE method of the document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    rows              The number of rows.                      1                Optional
                      Can be changed later.
                      
    columns           The number of columns.                   1                Optional
                      Can not yet be changed later.
                      
    width             The number of columns.                   1                Optional
                      This is a float. 1 means roughly
                      page width.
                      
    alignment         The alignment of the table. 'left',      'left'           Optional
                      'right', 'center' and 'both' are 
                      available.
                      
    style             The style of the table. Currently        'default'        Optional
                      'default' and 'borderless' are 
                      available.
    """
    
    table = doc.add_table(columns = 3, rows = 1)
    
    """
    The cells of the table can be accessed via the cells array of the table. Every cell has
    the add_paragraph, add_listpoint and add_table methods as presented before. Let us add
    a paragraph with text to the first cell in the first row:
    """
    
    p = table.cells[0][0].add_paragraph()
    p.add_text('This is the first cell.')
    
    """
    Let us add a list to the next cell.
    """
    
    list_template = doc.add_list_template(type = "roman")
    list_1 = doc.add_list(list_template)
    point = table.cells[0][1].add_list_point(list_1)
    point.add_text('First point of a list in a table!')
    point = table.cells[0][1].add_list_point(list_1)
    point.add_text('And this is another point...')
    
    """
    In the last cell we can insert another table.
    """
    
    nested_table = table.cells[0][2].add_table(columns = 2, rows = 2)
    
    """
    The ADD_ROW method of the table lets us add an additional row to the table.
    The method takes no arguments.
    """
    
    table.add_row()
    p = table.cells[-1][0].add_paragraph()
    p.add_text('We have added this row!')
    
    """
    Finally we can control the width of columns and the height of tables.
    To do this, the column_widths and row_heights attributes of the table can be manipulated.
    The system uses a weird unit where 11900 seems to be about the width of a page.
    """
    
    #Add paragraph for visibility
    doc.add_paragraph()
    
    #Add new table
    new_table = doc.add_table(rows = 3, columns = 3)
    
    #Set widths and heights
    new_table.column_widths = [1000, 2000, 3000]
    new_table.row_heights = [3000, 2000, 1000]
    
    """
    Another way to control the height and style of cells is to use spacing on the paragraphs
    they contain.
    """
    
    #Add paragraph for visibility
    doc.add_paragraph()
    
    #Add new table
    another_new_table = doc.add_table(rows = 3, columns = 3)
    
    #Add content
    p = another_new_table.cells[0][0].add_paragraph()
    p.add_text('This cell has no spacing')
    
    p = another_new_table.cells[1][0].add_paragraph(spacing_before = "120", spacing_after = "120")
    p.add_text('This cell has spacing of 120')
    
    p = another_new_table.cells[2][0].add_paragraph(spacing_before = "240", spacing_after = "240")
    p.add_text('This cell has spacing of 240')
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_4.docx')
    
    
"""
EXAMPLE 5 - USING IMAGES

Like text, images are added to paragraphs. There are two steps to working with images.
First, they have to be added to the document, then they can be used in a paragraph.
"""

def example_5():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    The image is added to the document with the IMPORT_IMAGE method of the document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    path              The path of the image to be imported.                     Required
                      Later images could be served with Python
                      or from an url.
                      
    """
    
    image_1 = doc.import_image('c:\\Users\\User\\Desktop\\conventec.jpeg')
    
    """
    For the moment only '.jpeg' images are supported, but more formats could be integrated.
    Now we can add a paragraph to the document and insert the image with the USE_IMAGE
    method of the paragraph:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    image             The image object previously imported.                     Required
    
    Listpoints have an equivalent USE_IMAGE method.
    """
    
    p = doc.add_paragraph(alignment = 'center')
    p.use_image(image_1)
    
    """
    If we want we can scale the image with the SET_WIDTH method of the image:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    width             A float between 0 and 1.                 0.1              Optional
                      1 is roughly the width of a page.
                      
    """
    
    image_1.set_width(width = 0.5)
    
    """
    Let us import another image and add it to the document twice:
    """
    
    image_2 = doc.import_image('c:\\Users\\User\\Desktop\\sponsors.jpeg')
    p = doc.add_paragraph()
    p.use_image(image_2)
    p.use_image(image_2)
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_5.docx')
    
"""
EXAMPLE 6 - HEADERS, FOOTERS AND PAGE NUMBERS

Headers and footers work identically. They both consist of three parts which are all optional.
'default', 'even' and 'first'. If set, 'even' is displayed on even pages, 'first' on the first
page and 'default' on all odd pages. To each part we can add paragraphs, list points and tables.
"""

def example_6():
    """
    Create the empty document:
    """
    
    doc = Document(title_page = True)
    
    """
    We set title_page to True in order to display a different header on the first page.
    Let us create a header which displays a centered conventec logo on the first page and a
    conventec logo on the top right on all odd pages. First the header is added to the document
    with the ADD_HEADER method, which takes no argument:
    """
    
    header = doc.add_header()
    
    #Import the conventec logo
    image_1 = doc.import_image('c:\\Users\\User\\Desktop\\conventec.jpeg')
    
    #Centered Logo on first page
    p = header.first.add_paragraph(alignment = 'center')
    p.use_image(image_1)
    
    #Logo in the top right as default
    p = header.default.add_paragraph(alignment = 'right')
    p.use_image(image_1)
    
    #Empty paragraph on even pages
    p = header.even.add_paragraph()
    
    """
    As a footer we might want to use the page number on all but the first page. To do this
    we can use the ADD_SIMPLEFIELD method of the paragaph:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    content           What should be displayed.                'page'           Optional
                      Currently only 'page' is available.
                      
    bold              If the text style should be              False            Optional
                      bold.                                   
        
    italics           If the text style should be              False            Optional
                      italic.
    
    underlined        If the text style should be              False            Optional
                      underlined.
                      
    small_caps        If the paragraph style should be         False            Optional
                      in small caps.
                      
    font_type         Text font. These are probably            'Arial'          Optional
                      supported:
                      https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
                      
    font_size         Text font size. (12pt = '24')            '24'             Optional
                      Use double the desired size.
                      e.g. (12pt = '24')
                      
    text_color        Text color.                              None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      RGB support could be implemented.
                      
    highlight_color   Text highlight color.                    None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...

    vertical_align    Vertical alignment of text.              None             Optional
                      e.g. 'superscript'
    """
    
    #Add footer to the document
    footer = doc.add_footer()
    
    #Page Number as default and even
    p = footer.default.add_paragraph(alignment = 'center')
    p.add_text('- ')
    p.add_simplefield()
    p.add_text(' -')
    
    p = footer.even.add_paragraph(alignment = 'center')
    p.add_text('- ')
    p.add_simplefield()
    p.add_text(' -')
    
    #Empty paragraph on first page
    footer.first.add_paragraph()
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_6.docx')
   

"""
EXAMPLE 7 - LINE AND PAGE BREAKS

Unfortunately line breaks are handled in many different ways by Word. An easy way to jump
to a new line is to add an additional paragraph. But the standard also supports explicit line
breaks. These work perfectly fine, but are apperently not used by Word and not supported
by Google Docs.

"""


def example_7():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    Line breaks are added with the ADD_LINE_BREAK method of paragraphs:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    n                 The number of lines to break.            1                Optional

    """
    
    p = doc.add_paragraph()
    p.add_text('This is some text')
    
    p.add_line_break(n = 1)
    
    p.add_text('And this is text after a line break.')
    
    """
    Line breaks also work within list points. But unfortunately not within tables:
    """
    
    table = doc.add_table(rows = 1, columns = 2)
    p = table.cells[0][0].add_paragraph()
    p.add_text('Here is no ')
    p.add_line_break()
    p.add_text('line break.')
    
    p = table.cells[0][1].add_paragraph()
    p.add_text('Here is a ')
    p = table.cells[0][1].add_paragraph()
    p.add_text('line break.')
    
    """
    The ADD_PAGE_BREAK method of paragraphs works similarly:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    n                 The number of page breaks.               1                Optional
    
    """
    
    p = doc.add_paragraph()
    p.add_text('This is before the page break.')
    p.add_page_break(n=2)
    p.add_text('And this is two pages later.')
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_7.docx')
    

"""
EXAMPLE 8 - NUMBERED TITLES

When documents are created dynamically, it is usefull to have automatically numbered titles.
You first have to create a numbered title template which can then be used for numbered titles.
From there you can add individual titles. Numbered titles can then also be referenced in the text.

"""


def example_8():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    We create two numbered title templates using the ADD_NUMBERED_TITLE_TEMPLATE method of the
    document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    indent            The indentation level of the title       '1'              Optional
    
    text              The text displayed before the number     ''               Optional
    
    style             The style of the numbering, 'numbering'  'numbering'      Optional
                      and 'roman' are supported.
                      
    separator         The separator between the number and     '.'              Optional
                      following text.
    """
    
    template_1 = doc.add_numbered_title_template(style = 'roman', separator = ' ')
    template_2 = doc.add_numbered_title_template(text = 'Article ', separator = '. ')
    
    """
    Now we can add titles with the ADD_NUMBERED_TITLE method of the document. 
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    template          The template the title belongs to.                        Required
    
    """
    
    titles_1 = doc.add_numbered_title(template_1)
    titles_2 = doc.add_numbered_title(template_2)
    
    """
    Now we are set up to create the actual title in the document with the ADD_TITLE method
    of the document. A title is basically a paragraph and takes the same arguments:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    titles            The titles the title belongs to.                          Required
    
    The arguments of a paragraph...
    """
    
    title_1 = doc.add_title(titles_1, alignment = "center", bold = True)
    title_1.add_text("General Part", bold = True)
    p = doc.add_paragraph()
    
    """
    Similarly we can add an article to the document:
    """
    
    article_1 = doc.add_title(titles_2, bold = True)
    article_1.add_text("Introduction", bold = True)
    
    p = doc.add_paragraph()
    p.add_text('Hello and welcome to this contract. The point of this document is to show numbered titles as well as references as supported by Conventec Docx.')
    
    """
    Titles can be referenced using the ADD_REFERENCE method of paragraphs:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    title             The title to be referenced.                                Required
    
    reference         What should be referenced. 'title'       'title'           Optional
                      and 'page' are supported.
                     
    The arguments of a text...
    """
    article_2 = doc.add_title(titles_2, bold = True)
    article_2.add_text("Goals", bold = True)
    
    p = doc.add_paragraph()
    p.add_text('In this article, we would like to elaborate some more on what we already said in ')
    p.add_reference(article_1)
    p.add_text(' on page ')
    p.add_reference(article_1, reference = 'page')
    p.add_text('.')
    
    """
    Page references have to be manually updated by the user in Word so they migh not be very
    convenient in practice.
    
    It is also possible to create a reference to a title that has not yet been created:
    """
    #Create new part
    p = doc.add_paragraph()
    title_2 = doc.add_title(titles_1, alignment = "center", bold = True)
    title_2.add_text("Specific Part", bold = True)
    p = doc.add_paragraph()
    
    #Add article that references a future article
    article_2 = doc.add_title(titles_2, bold = True)
    article_2.add_text("References to Future Articles", bold = True)
    
    p = doc.add_paragraph()
    p.add_text('In this article we add a reference to a future article, namely: ')
    
    #We create the reference even though the title has not yet been created
    future_reference = p.add_reference(title = None)
    
    article_3 = doc.add_title(titles_2, bold = True)
    article_3.add_text("An Article that has to be Refernced Before", bold = True)
    p = doc.add_paragraph()
    p.add_text('We have to reference this article in a previous one!')
    
    #Now we update the referenced title
    future_reference.title = article_3
    
    
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_8.docx')
    
    
"""
EXAMPLE 9 - HORIZONTAL LINES

Conventec Docx supports bottom borders for paragraphs which can both be used for styling the 
document as well as for creating signature fields

"""


def example_9():
    """
    Create the empty document:
    """
    
    doc = Document()
    
    """
    We can create a documente title with a horizontal line like this:
    """
    
    p = doc.add_paragraph(alignment = 'center', bold = True, font_size = '48')
    p.add_text('Document Title' , bold = True, font_size = '48')
    p = doc.add_paragraph(alignment = 'center', bold = True, font_size = '48')
    p = doc.add_paragraph(alignment = 'center', bold = True, font_size = '24')
    p.add_text('Document Subtitle' , bold = True, font_size = '24')
    p = doc.add_paragraph(alignment = 'center', border_bottom = True)
    
    """
    Horizontal lines are also useful for creating lines for signatures.
    For example like this:
    """
    
    table = doc.add_table(rows = 1, columns = 2, width = 1, alignment = "left", style= "borderless")
    
    p0 = table.cells[0][0].add_paragraph(border_bottom = True)
    p0.add_line_break()
    p1 = table.cells[0][0].add_paragraph()
    p1.add_line_break()
    p1.add_text('Mister A', font_size = "24")
    p2 = table.cells[0][0].add_paragraph()
    p2.add_text('CEO', font_size = "20", italics = True)
    
    p0 = table.cells[0][1].add_paragraph(border_bottom = True)
    p0.add_line_break()
    p1 = table.cells[0][1].add_paragraph()
    p1.add_line_break()
    p1.add_text('Mister B', font_size = "24")
    p2 = table.cells[0][1].add_paragraph()
    p2.add_text('CTO', font_size = "20", italics = True)
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_9.docx')
    

"""
EXAMPLE 10 - STYLES

In order to make working with styles more convenient, custom styles can be added to a document.
The styles can then be added to tables, headers, footers, lists, numbered titles, paragraphs,
listpoints, titles, texts, simplefields and breaks. Styles are inherited such that if for example
a paragraph is given a certain style, all text inside have the same default style. If no style is
given, a default style template is used. It is recommended that styles are added at the start of
the document.

"""


def example_10():
    """
    Create the empty document:
    """
    
    doc = Document(font_type='Comic Sans MS')
    
    """
    Custom styles can pe predefined for the entire document using the ADD_STYLE 
    method of the document:
    
    ARGUMENT          DESCRIPTION                              DEFAULT          Type
    
    reference_name    The name the style is referenced with                     Required 
    
    
    alignment         The alignment of paragraphs.             'left'           Optional
                      'left', 'right' and 'center' and
                      'both' is supported.
                      
    border_bottom     If paragraphs should have a solid        False            Optional
                      border at the bottom. Useful for
                      creating horizontal lines or 
                      signature fields.
                      
    keep_next         If paragraphs following the              False            Optional
                      current should be on the same
                      page. Useful for preventing page
                      breaks between title and text.
                      
    spacing_before    Adds a padding before the paragraph      None             Optional
                      120 might be an appropriate value.
                      
    spacing_after     Adds a padding after the paragraph       None             Optional
                      120 might be an appropriate value.
                      
    spacing_line      Spacing between lines. Takes a float.    '1.0'            Optional
                      
    indent            The indentation of the paragraph         '0'              Optional
    
    bold              If the text style should be              False            Optional
                      bold.
        
    italics           If the text style should be              False            Optional
                      italic.
    
    underlined        If the text style should be              False            Optional
                      underlined.
                     
    small_caps        If the paragraph style should be         False            Optional
                      in small caps.
                      
    font_type         Text font. These are probably            'Arial'          Optional
                      supported:
                      https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
                      
    font_size         Text font size.                          '24'             Optional
                      Use double the desired size.
                      e.g. (12pt = '24')
                      
    text_color        Text color.                              None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      RGB support could be implemented.
                       
    highlight_color   Text highlight color.                    None             Optional
                      Common colors are available.
                      e.g. red, yellow, blue...
                      
    vertical_align    Vertical alignment of text.              None             Optional
                      e.g. 'superscript'
        
    """
    
    doc.add_style('style 1', bold = True)
    doc.add_style('style 2', italics = True)
    
    """
    Lets create a list of style 1 and see how the style is inherited.
    
    """
    
    list_template = doc.add_list_template(type = "numbering")
    list1 = doc.add_list(list_template, style_name = 'style 1')
    
    point = doc.add_list_point(list1)
    point.add_text('This text is bold, because style 1 is inherited all the way down!')
    
    point = doc.add_list_point(list1, style_name = 'style 2')
    point.add_text('This text is italics, because the list point was given style 2.')
    
    point = doc.add_list_point(list1, style_name = 'style 2')
    point.add_text('This text is bold, because the text was given style 1.', style_name = 'style 1')
    
    point = doc.add_list_point(list1)
    point.add_text('This text is not bold, because we manually override the style!', bold = False)
    
    """
    Save the document.
    """
    
    doc.save(save_path + 'example_10.docx')
    
    