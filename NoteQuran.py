surah = 2
start = 60
stop = 101

import requests
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
import time


url = 'http://api.alquran.cloud/ayah/'

document = Document()
section = document.sections[0]
section.left_margin = Inches(.25)
section.right_margin = Inches(.25)
section.top_margin = Inches(.17)
section.bottom_margin = Inches(.17)

# rtlstyle = document.styles.add_style('rtlstyle', WD_STYLE_TYPE.PARAGRAPH)
# rtlstyle.base_style = document.styles['Normal']
# rtlstyle.font.rtl = True
# rtlstyle.font.cs_size = Pt(18)
# rtlstyle.font.name = 'Arial'
# rtlstyle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

table = document.add_table(rows=0, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for x in range(start, stop+1):
  print(x)
  while(True):
    try:
      ayah = requests.get(url + str(x)).json()['data']['text']
      tran = requests.get(url + str(x) + '/en.sahih').json()['data']['text']
      break
    except:
      print("except")
      time.sleep(.3) #Request fails so pause needed before retry.

  row_cells = table.add_row().cells
  row_cells[0].text = str(x) + ". " + tran
  row_cells[1].text = ayah
  # row_cells[1].paragraphs[0].style = rtlstyle
  # row_cells[1].paragraphs[0].runs[0].font.rtl = True
  row_cells[1].paragraphs[0].runs[0].font.name = 'Cambria'
  row_cells[1].paragraphs[0].runs[0].font.size = Pt(18)
  row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
  table.add_row()



document.save('test.docx')

print('done')