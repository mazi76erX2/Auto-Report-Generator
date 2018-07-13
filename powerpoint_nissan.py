#python-ppt-doc
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# load a presentation
prs = Presentation('Nissan.pptx')


# get reference to second slide
sl = prs.slides[4]

table = sl.shapes[3].table

#product, headline, publication, date, tone, reach, ave, link


def set_table_values(table, s):
    """Set string values from Series to the table
    """
    table.cell(1,2).text_frame.text = s.Product                 #Product
    table.cell(2,2).text_frame.text = s.Headline                #Headline
    table.cell(3,2).text_frame.text = "Star Power****"          #Publication
    table.cell(4,2).text_frame.text = "22-02-3000"              #Date
    table.cell(5,2).text_frame.text = "Balanced"                #Tone
    table.cell(6,2).text_frame.text = "91111"                   #Reach
    table.cell(7,2).text_frame.text = "R99 999,87"              #AVE
    table.cell(8,2).text_frame.text = "http://mazi/didthis"     #Link
    table.cell(9,2).text_frame.text = s.Extract                 #Extract

    return table


def set_table_styles(table):
    """styling from Series to the table as a list
    """
    for cell in table:
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
        cell.text_frame.paragraphs[0].font.name = 'Arial'
        cell.text_frame.paragraphs[0].font.bold = True

        ##  Not working
        ##  http://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-hyperlink.html
        #  Styling for link
        table.cell(8,2).text_frame.paragraphs[0].add_run()
        table.cell(8,2).text_frame.paragraphs[0].hyperlink =
                                                'https://www.google.com'

        #  Styling for extract
        table.cell(9,2).text_frame.paragraphs[0].font.bold = False
        table.cell(9,2).text_frame.paragraphs[0].font.italics = True
        table.cell(9,2).text_frame.paragraphs[0]\
                                            .font.color.rgb = RGBColor(89,89,89)

    return table
    

 

def populate_table_to_list(table):
    """Populate list from the Table (pptx.shapes.table object) values into
        a list of Cells (pptx.shapes.table.cell objects)
    """
    cell_list = []

    for i in range(1,10):
        cell_list.append(table.cell(i,2))

    return cell_list





prs.save('test.pptx')
print("Done!")
