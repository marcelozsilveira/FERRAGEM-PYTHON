import xml.etree.ElementTree as ET

tree = ET.parse('ELISABETE.xml')
root = tree.getroot()
total = []
for item in root.iter('ITEM'):
    total.append(item.attrib['DESCRICAO'])
for i in total:
    print(i, total.count(i))
    while i in total:
        total.remove(i)
