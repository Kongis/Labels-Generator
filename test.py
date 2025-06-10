from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
import io

# Pomocná funkce pro načtení zpracované šablony jako Document objekt
def render_to_document(template_path, context):
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    buffer = io.BytesIO()
    tpl.save(buffer)
    buffer.seek(0)
    return Document(buffer)

# Hlavní šablona
template_path = "test.docx"

# Seznam různých kontextů
contexts = [
    {"jmeno": "Anna", "vek": 25},
    {"jmeno": "Petr", "vek": 32},
    {"jmeno": "Lucie", "vek": 29},
]

# Vygeneruj první dokument a použij ho jako základ
main_doc = render_to_document(template_path, contexts[0])
composer = Composer(main_doc)

# Přidej zbylé
for context in contexts[1:]:
    sub_doc = render_to_document(template_path, context)
    composer.append(sub_doc)

# Uložení výsledku
composer.save("vystup.docx")