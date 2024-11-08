from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement

# Criar um novo documento Word
document = Document()

# Definir as margens
sections = document.sections
for section in sections:
    section.left_margin = Pt(85)    # 3 cm à esquerda
    section.top_margin = Pt(85)     # 3 cm superior
    section.right_margin = Pt(57)   # 2 cm à direita
    section.bottom_margin = Pt(57)  # 2 cm inferior

# Adicionar parágrafo de exemplo
p = document.add_paragraph("Este é um parágrafo de exemplo com as formatações especificadas.")

# Configurar a fonte e tamanho do texto
run = p.runs[0]
run.font.name = 'Arial'
run.font.size = Pt(12)

# Configurar o espaçamento entre linhas
p.paragraph_format.line_spacing = 1.5

# Justificar o texto nas margens esquerda e direita
p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# Salvar o documento
document.save("documento_formatado.docx")
