import docx
from docx2pdf import convert
import os

doc = docx.Document('input/example.docx')
os.mkdir("certificates")
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        if run.text == "??":
            names = open('input/names.txt', encoding='utf-8')
            run = run.clear()
            counter = 0
            for name in names:
                name = name.strip()
                print(name)
                run.add_text(name)
                doc.save(f'certificates/{name}.docx')
                convert(f"certificates/{name}.docx")
                os.remove(f'certificates/{name}.docx')
                counter += 1
            print(f"Созданы {counter} грамоты")
            names.close()