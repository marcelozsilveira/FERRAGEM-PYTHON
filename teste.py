import xml.etree.ElementTree as ET
import os

lista_ferragens = os.listdir('ferragens')

total = []
completo = {}
full = []

for arquivo in lista_ferragens:

    with open(f'ferragens/{arquivo}', 'rb') as ferragem:

        tree = ET.parse(ferragem)
        root = tree.getroot()
        for item in root.iter('ITEM'):
            total.append(item.attrib['DESCRICAO'])

    for i in total:
        completo[i] = total.count(i)
        for j in total:
            if j == i:
                total.remove(i)


for i in completo.keys():
    print(i)
#for i in completo.keys():
    #if 'Fita de Borda 22mm Rolo' in i or 'Fita de Borda 29mm Rolo' in i or 'Fita de Borda 45mm Rolo' in i:
        #print(i[8:18], i[24:])
#for i in full:  
    #print(i.count('Parafuso 3,5 x 16 un'))


        #print(f'---> {arquivo_xml}\n')
        #for k, v in completo.items():
            #print(k, ' -> ', v)


