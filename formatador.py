#feito por https://github.com/Vitor-Izidoro
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement

# Abrir um documento Word existente
document = Document('seu_arquivo.docx')

# Configurar margens conforme ABNT
sections = document.sections
for section in sections:
    section.left_margin = Pt(85)    # 3 cm à esquerda
    section.top_margin = Pt(85)     # 3 cm superior
    section.right_margin = Pt(57)   # 2 cm à direita
    section.bottom_margin = Pt(57)  # 2 cm inferior

# Iterar por cada parágrafo para ajustar formatação ABNT
for paragraph in document.paragraphs:
    # Verificar se o parágrafo é uma citação longa (ajuste manual ou reconhecimento de contexto necessário)
    if len(paragraph.text) > 40:  # Exemplo para detectar citação longa (40 caracteres ou mais)
        for run in paragraph.runs:
            run.font.size = Pt(10)  # Fonte tamanho 10 para citação longa
        paragraph.paragraph_format.line_spacing = 1.0  # Espaçamento simples
        paragraph.paragraph_format.left_indent = Pt(40)  # Recuo de 4 cm
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    else:
        # Configurar fonte e tamanho para corpo do texto
        for run in paragraph.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(12)

        # Espaçamento entre linhas 1,5 e alinhamento justificado
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Pt(12.5)  # Recuo de 1,25 cm

# Salvar o documento com a formatação ABNT
document.save('seu_arquivo_formatado_abnt.docx')
