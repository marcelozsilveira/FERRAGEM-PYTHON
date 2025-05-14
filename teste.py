import xml.etree.ElementTree as ET
import os

lista_ferragens = os.listdir('xml_s')

total = []
item = ()

for arquivo in lista_ferragens:

    with open(f'xml_s/{arquivo}', 'rb') as ferragem:

        tree = ET.parse(ferragem)
        root = tree.getroot()
        
        for i in root.iter('ITEM'):
            
            item = i.attrib['CAMINHOITEMCATALOG']
            if item not in total:
                total.append(item)
            del item           
for i in total:
    print(i)
