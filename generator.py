from docx import Document
import pandas as pd

dataFile = pd.read_excel('data.xlsx')
data = dataFile.to_dict('records')

for item in data:
    template = Document('template.docx')
    item['nameCap'] = item['name'].upper()
    item['document'] = str(item['document'])

    for para in template.paragraphs:
        for key, value in item.items():
            para.text = para.text.replace('{'+key+'}', value)

    template.save('output/' + str(item['document']) + '.docx')
