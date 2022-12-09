from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import re


nome_arquivo = input('Digite o Nome do Arquivo a ser Segmentado: ')
mes_ano_ref = input('Digite o Mês e Ano de Referência do Demonstrativo (Ex: 10 2022): ')

senhas = pd.read_excel('pdf/senhas.xlsx')
reader = PdfReader(f'pdf/{nome_arquivo}.pdf')
writer = PdfWriter('')

count = 0 
for page in reader.pages:
    texto = page.extract_text()
    for i in range(len(senhas)):
        nome = senhas.iloc[i,0]
        senha = senhas.iloc[i,1]
        if re.search(nome, texto, re.IGNORECASE):
            writer.add_page(page)
            with open(f'pdf/Demonstrativo {nome} {mes_ano_ref}.pdf', 'wb') as e:
                writer.encrypt(str(senha))
                writer.write(e)
                count += 1
                
print('======================================')
print(f'         {count} de {str(reader.numPages)} PDFs Salvos')
print('======================================')
final = input('Pressione Enter Para Sair')