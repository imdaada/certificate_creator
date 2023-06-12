import docx
from docx2pdf import convert

doc = docx.Document('example.docx')

for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        if run.text == "??":
            names = open('names.txt', encoding='utf-8')
            run = run.clear()
            counter = 0
            for name in names:
                name = name.strip()
                print(name)
                run.add_text(name)
                doc.save(f'{name}.docx')
                convert(f"{name}.docx")
                counter += 1
            print(f"Созданы {counter} грамоты")
            names.close()