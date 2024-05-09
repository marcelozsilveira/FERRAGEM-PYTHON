import xml.etree.ElementTree as ET
import os


def ler_xml(arquivo_xml):

    with open(f'ferragens/{arquivo_xml}', 'rb') as ferragem:

        tree = ET.parse(ferragem)
        root = tree.getroot()
        total = []
        completo = {}

        for item in root.iter('ITEM'):
            total.append(item.attrib['DESCRICAO'])

        for i in total:
            completo[i] = total.count(i)

        print(f'---> {arquivo_xml}\n')
        for k, v in completo.items():
            print(k, ' -> ', v)
        print()
        return completo


lista_ferragens = os.listdir('ferragens')

for ferragem in lista_ferragens:
    ler_xml(ferragem)