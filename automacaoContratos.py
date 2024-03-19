'''
Passos:
- Passar as informações da planilha para o arquivo word.
- Salvar essa arquivo em uma pasta espeífica (contratos).
- Repetir para todas a planilha.

Bibliotecas necessárias:
Openpyxl: Ler planilhas.
Python-docx: Cria arquivos Word.
'''
from openpyxl import load_workbook #Permite abrir planilhas.
from docx import Document #Permite abrir e criar docx.
from datetime import datetime #Permite inserir a data atual no contrato.

planilha_fornecedores = load_workbook('./fornecedores.xlsx') #Essa váriavel serve para abrir a planilha fornecedores. 
pagina_fornecedores = planilha_fornecedores['Sheet1'] #Especificando que é a página sheet1 do arquivo fornecedores.

#iter_rows função do Openpyxl que permite ler cada linha da planilha seguindo os parametros dos parentesis.
#min_row diz que é para começar a ler no minimo a partir da linha 2.
#values_only=True diz que é para ler somente valres verdadeiros, ou seja, o que for texto.
for linha in pagina_fornecedores.iter_rows(min_row=2,values_only=True)
    nome_empresa, endereco, cidade, estado, cep, telefone, email, setor = linha












