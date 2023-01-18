from docx import Document
import pandas as pd

df = pd.read_excel('data.xlsx')
data = df.to_dict('records')

for item in data:
    template = Document('template.docx')
    item['nameCap'] = item['name'].upper()

    for para in template.paragraphs:
        if '{name}' in para.text:
            para.text = para.text.replace('{name}', item['name'])
        if '{nameCap}' in para.text:
            para.text = para.text.replace('{nameCap}', item['nameCap'])
        if '{document}' in para.text:
            para.text = para.text.replace('{document}', str(item['document']))
        if '{marital}' in para.text:
            para.text = para.text.replace('{marital}', item['marital'])
        if '{address}' in para.text:
            para.text = para.text.replace('{address}', item['address'])

    template.save('output/' + str(item['document']) + '.docx')
