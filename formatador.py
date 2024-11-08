#feito por https://github.com/Vitor-Izidoro
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Abrir o documento existente
document = Document('seu_arquivo.docx')

# Configurar margens
sections = document.sections
for section in sections:
    section.left_margin = Pt(85)    # 3 cm à esquerda
    section.top_margin = Pt(85)     # 3 cm superior
    section.right_margin = Pt(57)   # 2 cm à direita
    section.bottom_margin = Pt(57)  # 2 cm inferior

# Aplicar formatações a cada parágrafo
for paragraph in document.paragraphs:
    # Configurar fonte Arial ou Times New Roman, tamanho 12, cor preta
    for run in paragraph.runs:
        run.font.name = 'Arial'  # Pode trocar para 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)  # Preto

    # Alinhamento justificado
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # Espaçamento de 1,5 entre linhas
    paragraph.paragraph_format.line_spacing = 1.5

# Salvar o documento com as formatações aplicadas
document.save('seu_arquivo_formatado_abnt.docx')
