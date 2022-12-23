from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import re

nome_arquivo = input('Digite o Nome do Arquivo a ser Segmentado: ')
mes_ano_ref = input('Digite o Mês e Ano de Referência do Demonstrativo (Ex: 10 2022): ')

senhas = pd.read_excel('pdf/senhas.xlsx')
reader = PdfReader(f'pdf/{nome_arquivo}.pdf')
writer = PdfWriter('')

count = 0 
#Percorre cada página do PDF
for page in reader.pages:
    #Extrai o texto da página e salva na variável
    texto = page.extract_text()
    #Percorre o arquivo nome e senha dos colaboradores
    for i in range(len(senhas)):
        #Salva o nome e senha do colaborador
        nome = senhas.iloc[i,0]
        senha = senhas.iloc[i,1]
        #Percorre o o texto da página atual e busca o nome do colaborador
        if re.search(nome, texto, re.IGNORECASE):
            #Cria um PDF da página
            writer.add_page(page)
            with open(f'pdf/Holerite {nome} {mes_ano_ref}.pdf', 'wb') as e:
                writer.encrypt(str(senha))
                writer.write(e)
                count += 1
            writer = PdfWriter('')

print('======================================')
print(f'         {count} de {str(reader.numPages)} PDFs Salvos')
print('======================================')
final = input('Pressione Enter Para Sair')