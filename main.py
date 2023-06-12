import shutil

import docx
from docx2pdf import convert
import os


def create_file(document, doc_name):
    doc.save(f'certificates/{doc_name}.docx')
    convert(f"certificates/{doc_name}.docx")
    os.remove(f'certificates/{doc_name}.docx')


def check_tables():
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text == "??":
                    names = open('input/names.txt', encoding='utf-8')
                    counter = 0
                    for name in names:
                        name = name.strip()
                        cell.text = name
                        print(name)
                        create_file(doc, name)
                        counter += 1
                        print(f"Созданы {counter} грамоты")
                    names.close()


def check_runs():
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.text == "??":
                names = open('input/names.txt', encoding='utf-8')
                counter = 0
                for name in names:
                    name = name.strip()
                    run = run.clear()
                    print(name)
                    run.add_text(name)
                    create_file(doc, name)
                    counter += 1
                print(f"Созданы {counter} грамоты")
                names.close()


doc = docx.Document('input/example.docx')
if os.path.exists("certificates"):
    shutil.rmtree("certificates")
os.mkdir("certificates")
check_tables()
check_runs()
