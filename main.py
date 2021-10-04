import docx
from docx import Document
from docx import table
from openpyxl import load_workbook
import fileinput



# Copy range of cells as a nested list
# Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    # Loops through selected Rows
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            rowSelected.append(sheet.cell(i, j).value)
        # Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected

def get_para_data(output_doc_name, paragraph, style, append_last_run_with=""):
    """
    Write the run to the new file and then set its font, bold, alignment, color etc. data.
    """

    output_para = output_doc_name.add_paragraph(style=style)
    for run in paragraph.runs:
        output_run = output_para.add_run(run.text)
        # Run's bold data
        output_run.bold = run.bold
        # Run's italic data
        output_run.italic = run.italic
        # Run's underline data
        output_run.underline = run.underline
        # Run's color data
        output_run.font.size = run.font.size
        print("fnt", run.font.size)
        output_run.font.name = run.font.name
        # Run's font data
        #output_run.style.name = run.style.name

    if append_last_run_with != "":
        output_para.runs[-1].text += append_last_run_with
        print("OUTPUT", output_para.text)

    # Paragraph's alignment data
    output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
    output_para.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
    output_para.paragraph_format.space_after = 0
    print("spc",paragraph.paragraph_format.space_after)



wb = load_workbook(filename = fileinput.filename()) # '30 august - 3 september 2021a.xlsx')
sheet = wb["30 august - 3 september 2021a"]

wb2 = load_workbook("lastehoiud ja kuupäevad.xlsx")
hoiud = wb2.active
#hoiud = wb["Leht1"]

dok = Document("data.saatelehed.docx")
dok2 = Document()
#changing the page margins
sections = dok2.sections
for section in sections:
    section.top_margin = dok.sections[0].top_margin
    section.bottom_margin = dok.sections[0].bottom_margin
    section.left_margin = dok.sections[0].left_margin
    section.right_margin = dok.sections[0].right_margin

#s = []
#stiil =[]

for hoid in range(3):
    for päev in range(5):
        i = 0
        table_range = copyRange(päev*3+1, 5, päev*3+3, 13, sheet)
        for r in dok.paragraphs:
            #for run in r.runs:
            #    s.append(run.text)
            #stiil.append(r.style)
            #print(r.style)
            #dok2.add_paragraph(s[i], stiil[i])

            if i==3:
                table1 = dok2.add_table(1, 2)
                #table1.style = "TableGrid"
                hdr_cells = table1.rows[0].cells
                hdr_cells[0].text = 'Toidu väljastamise kuupäev:'
                hdr_cells[1].text = str(hoiud.cell(row=hoid+2, column=päev+2).value)
                hdr_cells[1].paragraphs[0].runs[0].bold = True
                row_cells = table1.add_row().cells
                row_cells[0].text = "Toidu väljastamise kellaaeg: "
                row_cells[1].text = "Toit säilib kuni (kellaaeg): "

            if i==6:
                table = dok2.add_table(1,5)
                table.style = "TableGrid"
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Jrk. nr'
                hdr_cells[1].text = 'Toidu nimetus'
                hdr_cells[2].text = 'Portsu kaal  keskmiselt'
                hdr_cells[3].text = 'Kcal keskmiselt'
                hdr_cells[4].text = 'Portsjonite arv'
                hdr_cells[0].paragraphs[0].runs[0].bold = True
                hdr_cells[1].paragraphs[0].runs[0].bold = True
                hdr_cells[2].paragraphs[0].runs[0].bold = True
                hdr_cells[3].paragraphs[0].runs[0].bold = True
                hdr_cells[4].paragraphs[0].runs[0].bold = True
                j=0
                for toit, kaal, kcal in table_range:
                    j+=1
                    if toit != "" and toit != None and kcal != None and kaal != None:
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(j)
                        row_cells[1].text = toit
                        row_cells[2].text = str(kaal)+" g"
                        row_cells[3].text = str(round(kcal,2))+" kcal"
            if i == 1:
                print(r.text)
                get_para_data(dok2, r, r.style, str(hoiud.cell(row=hoid + 2, column=1).value))
            else:
                get_para_data(dok2, r, r.style)
            if i == 2:
                u = dok2.add_paragraph()
                u.paragraph_format.space_after = 0
            print(i)
            print(r.text)
            if i==10:
                print("lõpp")
            i+=1
        if päev%2!=0:
            dok2.add_page_break()
        else:
            a = dok2.add_paragraph()
            a.paragraph_format.space_after = 0
            b = dok2.add_paragraph()
            b.paragraph_format.space_after = 0
    dok2.add_page_break()
#print('\n'.join(s))
dok2.save('saatelehed.docx')
