from fpdf import FPDF
import pandas as pd
import os
import traceback

# templates & fonts path
logo_path = 'Third_Solution\Templates\logo.png'
logo_w_phone_path = 'Third_Solution\Templates\logo_w_phone.png'
qrcode_path = 'Third_Solution\Templates\qrcode.png'
database_path = 'Proposta_Comercial.pdf'
font_jost_bold = 'Third_Solution\Fonts\Jost-Bold.ttf'
font_jost_medium = 'Third_Solution\Fonts\Jost-Medium.ttf'
font_jost_black = 'Third_Solution\Fonts\Jost-Black.ttf'

class PDF(FPDF):

    def read_excel(self, file_path):
        df = pd.read_excel(file_path, sheet_name='Proposta_Comercial')
        return df
    
    def header(self):
        # LOGO
        self.image(logo_path, 10, 6, 50)

        # BTK INFO
            # LINE 1
        self.add_font('Jost-Bold', '', font_jost_bold, uni=True) # first time adding font_just_bold
        self.set_font('Jost-Bold', '', 8)
        self.set_x(60)
        self.cell(0, 3, 'BTK SOLUTIONS', 'C', ln=False)
        self.set_x(167)
        self.cell(0, 3, 'PROPOSTA COMERCIAL', 'C', ln=True)
            # LINE 2
        self.add_font('Jost-Medium', '', font_jost_medium, uni=True)  # first time adding font_just_medium
        self.set_font('Jost-Medium', '', 7)
        self.set_x(60)
        self.cell(0, 3, 'BTK SOLUTIONS DIST. DE EQUIP. E INTEGRACAO DE SISTEMAS ELETRONICOS LTDA', 'C', ln=False)
        x_position_1 = self.w - self.get_string_width('Nº ' + self.proposta_numero) - 10 # changeable string. as it can vary its size, i added a formula that will stuck the correct alignment
        self.set_x(x_position_1)
        self.cell(0, 3, 'Nº ' + self.proposta_numero, 'C', ln=True)
            # LINE 3
        self.set_x(60)
        self.cell(0, 3, 'CNPJ 32.421.603/0001-53 - I.E. 90890745-06', 'C', ln=False)
        x_position_2 = self.w - self.get_string_width(self.proposta_data) - 10
        self.set_x(x_position_2)
        self.cell(0, 3, self.proposta_data, 'C', ln=True)
            # LINE 4
        self.set_x(60)
        self.cell(0, 3, 'contato@btk.solutions - (41) 3385-7134', 'C', ln=True)
            # LINE 5
        self.set_x(60)
        self.cell(0, 3, 'CEP 83.050-160 - Rua Professora Marieta de Souza', 'C', ln=True)
            # LINE 6
        self.set_x(60)
        self.cell(0, 3, 'Parque da Fonte - São José Dos Pinhais - PR', 'C', ln=True)
            # SKIP 1 LINE
        self.ln(5)

    # BUYER INFO
    def add_general_info(self, general_info):
            # LINE 7
        self.set_font('Jost-Bold', '', 8)
        self.set_x(10)
        self.cell(0, 3, 'Cliente', ln=False)
        self.set_x(100)
        self.cell(0, 3, 'Dados Gerais', 'C', ln=True)
            # LINE 8
            # CLIENTE
        self.set_font('Jost-Medium', '', 8)
        self.set_x(10)
        self.cell(0, 3, general_info['Nome do Cliente'], 'C', ln=False)
        self.set_x(100)
        self.cell(0, 3, 'Condição de Pagamento: ' + general_info['Condição de Pagamento'], 'C', ln=False)
        self.set_x(160)
        self.cell(0, 3, 'Validade do Orçamento: ' + general_info['Validade do Orçamento'], 'C', ln=True)
            # LINE 9
        self.set_x(10)
        self.cell(0, 3, general_info['CNPJ'], 'C', ln=False)
        self.set_x(10)
        self.set_x(100)
        self.cell(0, 3, 'Prazo de Entrega: ' + general_info['Prazo de Entrega'], 'C', ln=False)
        self.set_x(160)
        self.cell(0, 3, 'Transportadora: ' + general_info['Transportadora'], 'C', ln=True)
            # LINE 10
        self.set_x(10)
        self.cell(0, 3, general_info['Endereço'] + ' - ' + general_info['Cidade'] + ' - ' + general_info['UF'] + ' - CEP: ' + general_info['CEP'], 'C', ln=False)     
        self.set_x(100)
        self.cell(0, 3, 'Frete: ' + general_info['Frete'], 'C', ln=False)
        self.set_x(160)
        self.cell(0, 3, 'Vendedor: ' + general_info['Vendedor'], 'C', ln=True)
        self.ln(2)

        # LINE DIVIDER
    def line_divider(self):
        self.set_fill_color(0, 0, 0)
        self.rect(10, self.get_y(), 190, 1, 'F')
        self.ln(3)


    def add_table_header(self):
        self.set_font('Jost-Bold', '', 7)
            # LINE 11
        self.cell(0, 3, 'Produtos', 0, ln=True)
        
        # Add table header
            # LINE 12
        self.set_font('Jost-Bold', '', 8)
        total_width = self.w - 20
        self.cell(total_width*0.05, 10, '#', 'B')
        self.cell(total_width*0.14, 10, 'SKU', 'B')
        self.cell(total_width*0.35, 10, 'Descrição', 'B')
        self.cell(total_width*0.14, 10, 'Quantidade', 'B')
        self.cell(total_width*0.13, 10, 'Valor', 'B')
        self.cell(total_width*0.12, 10, 'Subtotal', 'B')
        self.cell(total_width*0.07, 10, 'NCM', 'B')
        self.ln()

    def add_table_row(self, products):
            # LINE 13
        total_width = self.w - 20
        df = self.read_excel('Third_Solution\Proposta_Comercial.xlsx')
        self.add_font('Jost-Bold', '', font_jost_bold, uni=True)
        self.set_font('Jost-Bold', '', 7)

        for _, row in products.iterrows():
            if self.get_y() > 200:  # Ajuste conforme a margem desejada
                self.add_page()
            self.cell(total_width*0.05, 10, str(row['#']), 'B')
            self.cell(total_width*0.14, 10, row['SKU'], 'B')
            self.cell(total_width*0.35, 10, row['Descrição'], 'B')
            self.cell(total_width*0.14, 10, 'PC      '+ str(row['Quantidade']), 'B')
            self.cell(total_width*0.13, 10, f'R$ {row["Valor Unitário"]:.2f}', 'B')
            self.cell(total_width*0.12, 10, f'R$ {row["Subtotal"]:.2f}', 'B')
            self.cell(total_width*0.07, 10, str(row['NCM']), 'B')
            self.ln(10)
        self.ln(3)

    

    def add_totals(self, totals):
        if self.get_y() > 200:  # Ajuste conforme a margem desejada
            self.add_page()
        self.set_font('Jost-Bold', '', 7)
        # LINE 14
        self.cell(0, 3, 'Observações', 'C', ln=False)
        self.set_x(150)
        self.cell(0, 3, 'Totais', 'C', ln=True)
        # SKIP LINE
        self.ln(2)
        # LINE 15
            # OBSERVAÇOES
        self.set_font('Jost-Medium', '', 7)
        self.set_x(10)
        y_before = self.get_y()
        self.multi_cell(100, 3, totals['Observações'], align='J')
        self.set_xy(90, y_before)

        self.set_font('Jost-Bold', '', 7)
        self.set_x(150)
        self.cell(0, 3, 'Total dos produtos', 'C', ln=False)
        self.set_x(175)
        self.cell(0, 3, 'R$', 'C', ln=False)
        x_position_4 = self.w - self.get_string_width(str(totals['Total dos Produtos'])) - 10
        self.set_x(x_position_4)
        self.cell(0, 3, str(totals['Total dos Produtos']), align='R', ln=True)
        # LINE 16
        self.set_x(150)
        self.cell(0, 3, 'Impostos', 'C', ln=False)
        self.set_x(175)
        self.cell(0, 3, 'R$', 'C', ln=False)
        x_position_4 = self.w - self.get_string_width(str(totals['Impostos'])) - 10
        self.set_x(x_position_4)
        self.cell(0, 3, str(totals['Impostos']), align='R', ln=True)
        # LINE 17
        self.set_x(150)
        self.cell(0, 3, 'Frete', 'C', ln=False)
        self.set_x(175)
        self.cell(0, 3, 'R$', 'C', ln=False)
        self.set_x(150)
        x_position_4 = self.w - self.get_string_width(str(totals['Frete'])) - 10
        self.set_x(x_position_4)
        self.cell(0, 3, str(totals['Frete']), align='R', ln=True)
        # SKIP LINE
        self.ln(5)
        # LINE 18
        self.set_x(150)
        self.cell(0, 3, 'Total Geral', 'C', ln=False)
        self.set_x(175)
        self.cell(0, 3, 'R$', 'C', ln=False)
        self.set_x(150)
        self.cell(0, 3, str(totals['Total Geral']), align='R', ln=True)
        # SKIP LINE
        self.ln(3)

    # # LINE DIVIDER
    # def line_divider_2(self):
    #     self.set_fill_color(0, 0, 0)
    #     self.rect(10, self.get_y(), 190, 0.4, 'F')
    #     self.ln(3)

    #     self.ln(10)

    # FOOTER
    def footer(self):
        self.set_font('Jost-Medium', '', 7)
        self.set_y(-80)
        self.multi_cell(0, 3, '''Os termos e condições abaixo, assim como quaisquer modificações posteriormente realizadas, integrarão o contrato de compra e venda a ser celebrado entre BTK SOLUTIONS e o cliente acima indicado(”Cliente”), no que diz respeito aos produtos descritos acima e as obrigações de cada parte. Somente após a oferta emitida pela BTK SOLUTIONS ter sido aceita pelo Cliente passará a existir contrato de compra e venda, que será interpretado como um acordo final entre as partes e substitui qualquer outro acordo ou contato prévio oral ou escrito.
Condições Gerais:
1)Validade da Proposta: Conforme descrito no cabeçalho da cotação.
2)Forma de Pagamento: Conforme descrito no cabeçalho da cotação, em dias corridos da data do faturamento, a critério exclusivo da BTK SOLUTIONS (sujeito à análise e aprovação de crédito do Cliente).
3)Frete: Conforme descrito no cabeçalho da cotação.
4)Prazo de Entrega: A previsão informada nesta proposta deverá ser confirmada quando a proposta for efetivamente aceita pelo Cliente.
5)Transferência da propriedade e do risco: A propriedade e o risco sobre os produtos e/ou equipamentos, inclusive no que diz respeito a danos, é transferida ao Cliente quando da entrega ao primeiro transportador.
6)Garantiados Produtos/Equipamentos: O prazo de garantia é de 12 meses.
7)Os valores acima são válidos para solução completa e/ou pacotes como apresentados; não são válidos para vendas separadas de um ou mais itens.
8)Os valores de venda incluem todos os impostos incidentes sobre a operação de responsabilidade da BTK SOLUTIONS, sendo que todos os destaques serão feitos no termo da legislação em vigor.''', 1, 'J')
        self.ln(3)

        self.image('Third_Solution\Templates\logo_w_phone.png', 15, 275, 30)

        self.image('Third_Solution\Templates\qrcode.png', 175, 270, 20)

        self.add_font('Jost-Black', '', font_jost_black, uni=True)  # first time adding font_just_black
        self.set_font('Jost-Black', '', 15)
        self.set_xy(15, 275)
        self.multi_cell(0, 7, '''DESENVOLVENDO SOLUÇÕES
SOLUCIONANDO O AMANHÃ''', 0, align='C')
        
        


        


# Function to create the PDF
def criar_pdf(caminho_excel, caminho_pdf):
    xls = pd.ExcelFile(caminho_excel)
    
    proposta_comercial = pd.read_excel(xls, 'Proposta_Comercial').iloc[0]
    dados_gerais = pd.read_excel(xls, 'Dados_Gerais').iloc[0]
    produtos = pd.read_excel(xls, 'Produtos')
    totais = pd.read_excel(xls, 'Totais').iloc[0]
    
    pdf = PDF()
    pdf.proposta_numero = proposta_comercial['Número da Proposta']
    pdf.proposta_data = proposta_comercial['Data']
    pdf.add_page()
    

    pdf.add_general_info(dados_gerais)
    pdf.line_divider()
    pdf.add_table_header()
    pdf.add_table_row(produtos)
    pdf.add_totals(totais)
    # pdf.line_divider_2()
    
    pdf.output(caminho_pdf)

    # pdf.output('Proposta_Comercial.pdf')

# # Path to the Excel file
# caminho_excel = 'Third_Solution\Proposta_Comercial.xlsx'
# criar_pdf(caminho_pdf)