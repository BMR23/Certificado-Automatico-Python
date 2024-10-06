import openpyxl
from PIL import ImageFont, ImageDraw, Image

workbook_alunos = openpyxl.load_workbook("planilha_alunos.xlsx")
sheet_alunos = workbook_alunos["Sheet1"]

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    fonte_nome = ImageFont.truetype('./fonts/tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./fonts/tahoma.ttf', 80)

    image = Image.open('./certificados/padrao/certificado_padrao.jpg')
    desenhar = ImageDraw. Draw(image)

    desenhar.text((1020, 825), nome_participante,
                  fill='black', font=fonte_nome)
    desenhar.text((1070, 954), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1453, 1070), tipo_participacao,
                  fill='black', font=fonte_geral)
    desenhar.text((1490, 1187), f'{carga_horaria}', fill='black', font=fonte_geral)

    desenhar.text((675, 1755), data_inicio, fill='black', font=fonte_geral)
    desenhar.text((675, 1910), data_final, fill='black', font=fonte_geral)

    desenhar.text((2165, 1910), data_emissao, fill='black', font=fonte_geral)
    
    image.save(f"./certificados/{indice} {nome_participante} certificado.png")
