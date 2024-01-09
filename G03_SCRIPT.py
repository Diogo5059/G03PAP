import requests
import os
import pandas as pd
import tabula.io

# Função para fazer o download do arquivo do GitHub
def download_file_from_github(url, file_name):
    response = requests.get(url)
    if response.status_code == 200:
        with open(file_name, 'wb') as file:
            file.write(response.content)
        return True
    else:
        return False

# Diretório onde você deseja salvar os PDFs e os arquivos Excel convertidos
download_directory = 'C:\\pap\\PDFs'
excel_directory = 'C:\\pap\\ExcelTables'

# URL do repositório do GitHub onde os PDFs estão localizados
github_url = 'https://github.com/Diogo5059/G03PAP/raw/main/'

# Lista de arquivos no repositório do GitHub
github_files = [
    '100_faltas.pdf',
    '200_faltas.pdf',
    # Adicione todos os nomes dos arquivos PDF que você deseja baixar
]

# Baixar arquivos do GitHub e converter para Excel
for file in github_files:
    pdf_url = github_url + file
    local_pdf_name = os.path.join(download_directory, file)
    excel_name = os.path.splitext(file)[0] + '_converted.xlsx'  # Modificação no nome do arquivo Excel

    excel_output = os.path.join(excel_directory, excel_name)

    # Baixar o arquivo PDF do GitHub
    download_success = download_file_from_github(pdf_url, local_pdf_name)

    # Se o download for bem-sucedido, extrair tabelas de cada página e salvar como Excel
    if download_success:
        # Obter o número de páginas no PDF
        with open(local_pdf_name, 'rb') as file:
            pages = len(tabula.io.read_pdf(file, pages='all', multiple_tables=True))

        # Extrair tabelas de cada página e salvar como Excel
        for page in range(1, pages + 1):
            tables = tabula.io.read_pdf(local_pdf_name, pages=page, multiple_tables=True)
            if tables:
                for i, table in enumerate(tables):
                    excel_file_name = f'{os.path.splitext(excel_name)[0]}_page{page}_table{i+1}.xlsx'
                    table.to_excel(os.path.join(excel_directory, excel_file_name), index=False)
                    print(f'Table {i+1} on page {page} extracted from {file} and saved as {excel_file_name}')
            else:
                print(f'No tables found on page {page} in {file}')
    else:
        print(f'Failed to download the file {file} from GitHub.')


caminho_arquivo_excel1 = r'C:\pap\ExcelTables\200_faltas_converted_page1_table1.xlsx'
Faltas_200 = pd.read_excel(caminho_arquivo_excel1, sheet_name='Sheet1')

caminho_arquivo_excel2 = r'C:\pap\ExcelTables\Alunos.xlsx'
Alunos = pd.read_excel(caminho_arquivo_excel2, sheet_name='Page 1')

caminho_arquivo_excel3 = r'C:\pap\ExcelTables\professores.xlsx'
Professores = pd.read_excel(caminho_arquivo_excel3, sheet_name='Page 1')

caminho_arquivo_excel4 = r'C:\pap\ExcelTables\Turmas.xlsx'
Turmas = pd.read_excel(caminho_arquivo_excel4, sheet_name='Page 1')

caminho_arquivo_excel5 = r'C:\pap\ExcelTables\100_faltas_converted_page1_table1.xlsx'
Faltas_100 = pd.read_excel(caminho_arquivo_excel5, sheet_name='Sheet1')

# Retornar o DataFrame para o Power BI

Faltas_100
Faltas_200
Alunos
Professores
Turmas