import PyPDF2
import re
from openpyxl import Workbook

patron = r"^(\d+,\d+)"

dict_aprobados_examen = {}

for i in range(1,13):
    lista_aprobados = []
    # Abre el archivo PDF en modo de lectura binaria
    with open('Escritorio/Oposiciones/PDFs/tribunal'+ str(i) +'.pdf', 'rb') as file:
        # Crea un objeto de lectura de PDF
        reader = PyPDF2.PdfReader(file)

        # Itera sobre cada página del PDF
        for page in reader.pages:
            # Extrae el texto de la página actual
            text = page.extract_text()
            text_list = text.split('\n')

            for e in text_list:
                resultado = re.search(patron, e)
                if resultado:
                    e = e.split()
                    valor_numerico = e[0]
                    valor_numerico = valor_numerico.split('*')
                    valor_numerico = valor_numerico[0]
                    valor_numerico = valor_numerico.replace(',','.')
                    nombre = ''
                    nif = ''
                    cont = 0
                    if e[1] != '-':
                        nif = e[1]
                    else:
                        print(e)
                        cont+=1
                    for l in e:
                        if l[0].lower().isalpha():
                            nombre += l + ' '
                    lista = [float(valor_numerico), nombre, nif]
                    lista_aprobados.append(lista)
        dict_aprobados_examen['Tribunal' + str(i)] = lista_aprobados
        print(cont)
dict_aprobados_final = {}
for i in range (1, 13):
    lista_meritos = []
    try: 
        with open('Escritorio/Oposiciones/PDFs/meritosTribunal'+ str(i) +'.pdf', 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text = page.extract_text()
                text_list = text.split('\n')

                for l in text_list:
                    if '***' in l:
                        l = l.split()
                        v_meritos = l[len(l) - 11]
                        v_meritos = v_meritos.replace(',','.')
                        nif = l[len(l) - 13]
                        nombre = ''
                        for e in l:
                            if e[0].lower().isalpha():
                                nombre += e + ' '
                        lista = [nombre, float(v_meritos), nif]
                        lista_meritos.append(lista)
        lista_opositores = dict_aprobados_examen['Tribunal'+str(i)]
        
        lista_opositores_examen_baremo = []
        for opositor_baremo in lista_meritos: 
            for opositor_examen in lista_opositores:
                if opositor_baremo[2] != '' and opositor_baremo[2] == opositor_examen[2]:
                    opositor_examen.append(opositor_baremo[1])
                    lista_opositores_examen_baremo.append(opositor_examen)

        print('Tribunal'+str(i), len(lista_opositores), len(lista_meritos))
        dict_aprobados_final['Tribunal'+ str(i)] = lista_opositores_examen_baremo

    except:
        print('No está el tribunal ', i)


workbook = Workbook()
hoja = workbook.active

hoja['A1'] = 'TRIBUNAL'
hoja['B1'] = 'NOMBRE'
hoja['C1'] = 'Nota Examen (sobre 10)'
hoja['D1'] = 'Nota Examen Ponderada (60%)'
hoja['E1'] = 'Méritos (sobre 10)'
hoja['F1'] = 'Méritos Ponderada (40%)'
hoja['G1'] = 'NOTA FINAL'
n=2
for tribunal in dict_aprobados_final.keys():
    
    for opositor in dict_aprobados_final[tribunal]:
        hoja['A'+str(n)] = tribunal
        hoja['B'+str(n)] = opositor[1]
        hoja['C'+str(n)] = opositor[0]
        hoja['E'+str(n)] = opositor[3]
        n+=1

workbook.save('Escritorio/Oposiciones/oposiciones.xlsx')
