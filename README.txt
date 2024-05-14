# BTK Solutions PDF Generator

## Descrição
O PDF Generator da BTK Solutions é uma ferramenta projetada para automatizar a criação de propostas comerciais em formato PDF a partir de dados de um arquivo Excel. 

## Como usar
- Verifique todas as dependências necessárias
- Certifique-se de que a planilha do Excel a ser utilizada está na mesma pasta do executável
- Estou vendo como faz para criar executáveis com imagens, então por enquanto, o codigo funciona apenas no terminal
- proposta_btk.py : é o backend, a criação toda do orçamento automatico em pdf
- btk_executavel.py : frontend, clique ali para rodar, selecione o arquivo excel, onde quer salvar, clique em gerar e então digite como quer salvar o nome do arquivo

## Requisitos
- Python 3.8+
- Bibliotecas: pandas, fpdf, customtkinter, Pillow
- Sistema operacional: Windows, Linux ou macOS

## Instalação
Clone o repositório e instale as dependências:
```bash
git clone https://github.com/seuusuario/btk-pdf-generator.git
cd btk-pdf-generator
pip install -r requirements.txt

