import xml.etree.ElementTree as ET
import os
import xlsxwriter

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

        workbook = xlsxwriter.Workbook(f'{arquivo_xml}.xlsx')

        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 32)
        worksheet.set_column('D:D', 10)

        formato_oc = workbook.add_format(
            {
            'font' : 'Arial',
            'font_size' : 16
            }
        )

        formato = workbook.add_format(
            {
            'border' : 1,
            'border_color' : '#000000',
            'font' : 'Arial',
            'font_size' : 14
            }
        )

        formato_c = workbook.add_format(
            {
            'border' : 1,
            'border_color' : '#000000',
            'font' : 'Arial',
            'font_size' : 14,
            'align' : 'center'
            }
        )


        #CABEÇALHO


        worksheet.write('A1', f'{arquivo_xml[:-4]}', formato_oc)
        worksheet.write('C1', 'Data:', formato_oc)
        worksheet.write('A2', 'Item:', formato_c)
        worksheet.write('B2', 'Qtds', formato_c)
        worksheet.write('C2', 'Item', formato_c)
        worksheet.write('D2', 'Qtds', formato_c)


        #COLUNA 1


        worksheet.write('A3', 'Dobradiça Reta', formato)
        worksheet.write('A4', 'Dobradiça Curva', formato)
        worksheet.write('A5', 'Dobradiça 45°', formato)
        worksheet.write('A6', 'Dobradiça Canto Cego', formato)
        worksheet.write('A7', 'Dobradiça Extra Alta', formato)
        worksheet.write('A8', 'Dobradiça 165°', formato)
        worksheet.write('A9', 'Dobradiça', formato)
        worksheet.write('A10', 'Minifix', formato)
        worksheet.write('A11', 'Atarraxante', formato)
        worksheet.write('A12', 'Passante', formato)
        worksheet.write('A13', 'Tambor', formato)
        worksheet.write('A14', 'Sup. de Prateleira', formato)
        worksheet.write('A15', 'Sup. de Cabide', formato)
        worksheet.write('A16', 'Sup. de Canto 90°', formato)
        worksheet.write('A17', 'Cantoneira 13x13', formato)
        worksheet.write('A18', '3,5x16', formato)
        worksheet.write('A19', '4,0x16', formato)
        worksheet.write('A20', '4,0x20', formato)
        worksheet.write('A21', '4,0x25', formato)
        worksheet.write('A22', '4,0x30', formato)
        worksheet.write('A23', '4,0x35', formato)
        worksheet.write('A24', '4,0x40', formato)
        worksheet.write('A25', '4,0x50', formato)
        worksheet.write('A26', '3,9x9,5', formato)
        worksheet.write('A27', '3,5x40', formato)
        worksheet.write('A28', '3,5x20 Flange', formato)
        worksheet.write('A29', '6,0x70', formato)
        worksheet.write('A30', 'Bucha 8mm', formato)
        worksheet.write('A31', '6,0x60 Flange', formato)
        worksheet.write('A32', '4,0x22', formato)
        worksheet.write('A33', 'Cavilha 8', formato)
        worksheet.write('A34', 'Cavilha 6', formato)
        worksheet.write('A35', 'Cola', formato)
        worksheet.write('A36', '4,2x13 Ponta Broca', formato)
        worksheet.write('A37', 'Marca', formato)
        worksheet.write('A38', '', formato)
        worksheet.write('A39', '', formato)
        worksheet.write('A40', '', formato)
        worksheet.write('A41', '', formato)


        #COLUNA 2


        if 'Dobradiça Reta Toque' in completo.keys():
            worksheet.write('B3', f'{completo["Dobradiça Reta Toque"]}', formato)
        if 'Dobradiça Curva Slow' in completo.keys():
            worksheet.write('B4', f'slow {completo["Dobradiça Curva Slow"]}', formato)
        worksheet.write('B5', '0', formato)
        worksheet.write('B6', '0', formato)
        worksheet.write('B7', '0', formato)
        worksheet.write('B8', '0', formato)
        worksheet.write('B9', '0', formato)
        worksheet.write('B10', '0', formato)
        if 'Atarraxante' in completo.keys():
            worksheet.write('B11', f'{completo["Atarraxante"]}', formato)
        worksheet.write('B12', '0', formato)
        if 'Tambor' in completo.keys():
            worksheet.write('B13', f'{completo["Tambor"]}', formato)
        if 'Suporte Pino p/ Prateleira' in completo.keys():
            worksheet.write('B14', f'{completo["Suporte Pino p/ Prateleira"]}', formato)
        if 'Suporte Cabide (conj. 2x peças)' in completo.keys():
            worksheet.write('B15', f'{completo["Suporte Cabide (conj. 2x peças)"]*2}', formato)
        worksheet.write('B16', '0', formato)
        if 'Cantoneira 13 x 13' in completo.keys():
            worksheet.write('B17', f'{completo["Cantoneira 13 x 13"]}', formato)
        if 'Parafuso 3,5 x 16 un' in completo.keys():
            worksheet.write('B18', f'{completo["Parafuso 3,5 x 16 un"]}', formato)
        if 'Parafuso 4 x 16 un' in completo.keys():
            worksheet.write('B19', f'{completo["Parafuso 4 x 16 un"]}', formato)
        if 'Parafuso 4 x 20 un' in completo.keys():
            worksheet.write('B20', f'{completo["Parafuso 4 x 20 un"]}', formato)
        worksheet.write('B21', '0', formato)
        if 'Parafuso 4 x 30 un' in completo.keys():
            worksheet.write('B22', f'{completo["Parafuso 4 x 30 un"]}', formato)
        worksheet.write('B23', '0', formato)
        if 'Parafuso 4 x 40 un' in completo.keys():
            worksheet.write('B24', f'{completo["Parafuso 4 x 40 un"]}', formato)
        worksheet.write('B25', '0', formato)
        worksheet.write('B26', '0', formato)
        if 'Parafuso 3,5 x 40 un' in completo.keys():
            worksheet.write('B27', f'{completo["Parafuso 3,5 x 40 un"]}', formato)
        if 'Parafuso 3,5 x 20 un Flangeado' in completo.keys():
            worksheet.write('B28', f'{completo["Parafuso 3,5 x 20 un Flangeado"]}', formato)
        if 'Parafuso 6 X 70' in completo.keys():
            worksheet.write('B29', f'{completo["Parafuso 6 X 70"]}', formato)
        if 'Bucha 8mm Comum' in completo.keys():
            worksheet.write('B30', f'{completo["Bucha 8mm Comum"]}', formato)
        if 'Parafuso 6 x 60 un Flangeado' in completo.keys():
            worksheet.write('B31', f'{completo["Parafuso 6 x 60 un Flangeado"]}', formato)
        worksheet.write('B32', '0', formato)
        if 'Cavilha' in completo.keys():
            worksheet.write('B33', f'{completo["Cavilha"]}', formato)
        if 'Cavilha 6X30' in completo.keys():
            worksheet.write('B34', f'{completo["Cavilha 6X30"]}', formato)
        worksheet.write('B35', '0', formato)
        worksheet.write('B36', '0', formato)
        worksheet.write('B37', '0', formato)
        worksheet.write('B38', '0', formato)
        worksheet.write('B39', '0', formato)
        worksheet.write('B40', '0', formato)
        worksheet.write('B41', '0', formato)


        #COLUNA 3


        worksheet.write('C3', 'Kit porta de correr', formato)
        worksheet.write('C4', 'Macho e Femea', formato)
        worksheet.write('C5', 'Fecho toque magnético', formato)
        worksheet.write('C6', 'Kit Cama', formato)
        worksheet.write('C7', 'Rod. Silicone', formato)
        worksheet.write('C8', 'Rod. Silicone', formato)
        worksheet.write('C9', 'Toalheiro', formato)
        worksheet.write('C10', 'Pistão', formato)
        worksheet.write('C11', 'Pistão', formato)
        worksheet.write('C12', 'Pistão inverso', formato)
        worksheet.write('C13', 'Sapatas', formato)
        worksheet.write('C14', 'Chapa união', formato)
        worksheet.write('C15', 'Pino Inglês', formato)
        worksheet.write('C16', 'Pux.', formato)
        worksheet.write('C17', 'Pux.', formato)
        worksheet.write('C18', 'Prego s/ Cabeça', formato)
        worksheet.write('C19', 'Prego c/ Cabeça', formato)
        worksheet.write('C20', 'Tampa de Tambor', formato)
        worksheet.write('C21', 'Tapa furo 15', formato)
        worksheet.write('C22', 'Tapa furo 15', formato)
        worksheet.write('C23', 'Tapa furo 15', formato)
        worksheet.write('C24', 'Cant. c/ Capa', formato)
        worksheet.write('C25', 'Cant. c/ Capa', formato)
        worksheet.write('C26', 'Adesivos', formato)
        worksheet.write('C27', 'Adesivos', formato)
        worksheet.write('C28', 'Adesivos', formato)
        worksheet.write('C29', 'Adesivos', formato)
        worksheet.write('C30', 'Batente Silicone', formato)
        worksheet.write('C31', 'Passa fio', formato)
        worksheet.write('C32', 'Passa fio', formato)
        worksheet.write('C33', 'Silicone', formato)
        worksheet.write('C34', 'Silicone', formato)
        worksheet.write('C35', 'Borda', formato)
        worksheet.write('C36', 'Borda', formato)
        worksheet.write('C37', 'Fita Aveludada', formato)
        worksheet.write('C38', 'Pé Plastico', formato)
        worksheet.write('C39', '', formato)
        worksheet.write('C40', '', formato)
        worksheet.write('C41', '', formato)


        #COLUNA 4


        worksheet.write('D3', '0', formato)
        worksheet.write('D4', '0', formato)
        worksheet.write('D5', '0', formato)
        worksheet.write('D6', '0', formato)
        worksheet.write('D7', '0', formato)
        worksheet.write('D8', '0', formato)
        worksheet.write('D9', '0', formato)
        worksheet.write('D10', '0', formato)
        worksheet.write('D11', '0', formato)
        worksheet.write('D12', '0', formato)
        worksheet.write('D13', '0', formato)
        worksheet.write('D14', '0', formato)
        worksheet.write('D15', '0', formato)
        worksheet.write('D16', '0', formato)
        worksheet.write('D17', '0', formato)
        worksheet.write('D18', '0', formato)
        worksheet.write('D19', '0', formato)
        worksheet.write('D20', '0', formato)
        worksheet.write('D21', '0', formato)
        worksheet.write('D22', '0', formato)
        worksheet.write('D23', '0', formato)
        worksheet.write('D24', '0', formato)
        worksheet.write('D25', '0', formato)
        worksheet.write('D26', '0', formato)
        worksheet.write('D27', '0', formato)
        worksheet.write('D28', '0', formato)
        worksheet.write('D29', '0', formato)
        worksheet.write('D30', '0', formato)
        worksheet.write('D31', '0', formato)
        worksheet.write('D32', '0', formato)
        worksheet.write('D33', '0', formato)
        worksheet.write('D34', '0', formato)
        worksheet.write('D35', '0', formato)
        worksheet.write('D36', '0', formato)
        worksheet.write('D37', '0', formato)
        worksheet.write('D38', '0', formato)
        worksheet.write('D39', '0', formato)
        worksheet.write('D40', '0', formato)
        worksheet.write('D41', '0', formato)


        workbook.close()


lista_ferragens = os.listdir('ferragens')

for ferragem in lista_ferragens:
    ler_xml(ferragem)