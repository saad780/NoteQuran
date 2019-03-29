surah = 1
start = 1
stop = 3

import requests
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt


url = 'http://api.alquran.cloud/ayah/'

# rayah = requests.get(url + '1')
# rtran = requests.get(url + '1/en.sahih')

# tran = rtran.json()['data']['text']
# ayah = rayah.json()['data']['text']

document = Document()
section = document.sections[0]
section.left_margin = Inches(.1)
section.right_margin = Inches(.1)
section.top_margin = Inches(.1)
section.bottom_margin = Inches(.1)

table = document.add_table(rows=0, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for x in range(start, stop+1):
  ayah = requests.get(url + str(x)).json()['data']['text']
  tran = requests.get(url + str(x) + '/en.sahih').json()['data']['text']
  row_cells = table.add_row().cells
  row_cells[0].text = tran
  row_cells[1].text = ayah
  row_cells[1].paragraphs[0].runs[0].font.rtl = True
  row_cells[1].paragraphs[0].runs[0].font.name = 'Cambria'
  row_cells[1].paragraphs[0].runs[0].font.size = Pt(18)
  row_cells[1].text = ayah



document.save('test.docx')

# print(requests.get(url + '1').json()['data']['text'])
# print(ayah)
# print(tran)
# print('\n\n')
print('done')