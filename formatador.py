#feito por https://github.com/Vitor-Izidoro
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Abrir um documento Word existente
document = Document('seu_arquivo.docx')

# Configurar margens
sections = document.sections
for section in sections:
    section.left_margin = Pt(85)    # 3 cm à esquerda
    section.top_margin = Pt(85)     # 3 cm superior
    section.right_margin = Pt(57)   # 2 cm à direita
    section.bottom_margin = Pt(57)  # 2 cm inferior

# Iterar por cada parágrafo para ajustar formatação
for paragraph in document.paragraphs:
    # Configurar fonte e tamanho
    for run in paragraph.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(12)

    # Configurar espaçamento entre linhas e alinhamento justificado
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# Salvar o documento formatado
document.save('seu_arquivo_formatado.docx')
