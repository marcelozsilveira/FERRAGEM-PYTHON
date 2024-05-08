import xml.etree.ElementTree as ET

tree = ET.parse('ELISABETE.xml')
root = tree.getroot()
total = []
completo = {}
for item in root.iter('ITEM'):
    total.append(item.attrib['DESCRICAO'])
for i in total:
    completo[i] = total.count(i)
for k, v in completo.items():
    print(k, ' -> ', v)
