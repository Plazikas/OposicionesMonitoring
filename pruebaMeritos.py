import PyPDF2
import re
from openpyxl import Workbook

lista_aprobados = []

with open('Escritorio/Oposiciones/PDFs/meritosTribunal4.pdf', 'rb') as file:
    reader = PyPDF2.PdfReader(file)
    n = 0
    for page in reader.pages:
        text = page.extract_text()
        text_list = text.split('\n')

        for l in text_list:
            if '***' in l:
                l = l.split()
                v_meritos = l[len(l) - 11]
                nombre = ''
                for e in l:
                    if e[0].lower().isalpha():
                        nombre += e + ' '
                lista = [nombre, v_meritos]
                lista_aprobados.append(lista)

lista_ordenada = sorted(lista_aprobados, key= lambda x: x[1], reverse=True)

for opositor in lista_ordenada:
    print(opositor)


