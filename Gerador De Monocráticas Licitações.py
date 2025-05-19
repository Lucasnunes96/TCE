from pathlib import Path  # Para manipulação de caminhos de arquivos
from docxtpl import DocxTemplate  # Para preencher templates do Word
import pandas as pd  # Para trabalhar com dados em formato de tabela e excel
from num2words import num2words  # Para converter números em extenso
import locale  # Para formatação de acordo com a localidade
#também instalar openpyxl

# Configuração do locale para o Brasil (formatação de moeda, etc.)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Definição dos caminhos dos arquivos
base_dir = Path(__file__).parent
word_template_path = base_dir / "Modelo Contratos - sem relatório - Python.docx"
excel_path = base_dir / "Lista de Dados - Contratos - Sem relatório - Python.xlsx"
output_dir = base_dir / "Monocraticas"

# Cria o diretório de saída se ele não existir
output_dir.mkdir(exist_ok=True)

# Lê os dados da planilha Excel, especificando os tipos de dados (datatypes) de cada coluna
df = pd.read_excel(excel_path, sheet_name='Elaborar Processos', dtype={'PROCESSO': str, 'CONTRATANTE': str, 'INTERESSADO': str,
                                                              'NMODALIDADE': str, 'NCONTRATO': str,
                                                              'EXERCICIO': str, 'ASSINATURACONTRATO': str,
                                                              'CONTRATADA': str, 'OBJETO': str, 'VALOR': float, 'POREXTENSO': float,
                                                              'PJPF': str, 'FUNDAMENTAÇÃO': str, 'DIRETORIA': str,
                                                              'SIGLA': str, 'DESPACHON': str, 'ESFERA': str, 'ENTRADA': str,
                                                              'RECEBIDO': str, 'PINTECORRENTE': str,'CONTAGEMDETEMPO1': str,
                                                              'POREXTENSO1': str, 'DATADEASSINATURA': str, 'PROCESSON': str,
                                                              'ANOPROCESSO': str, 'INICIAIS': str})

# Formata a coluna VALORCONTRATO para o formato de moeda brasileira (R$)
df['VALOR'] = df['VALOR'].apply(lambda x: locale.format_string('R$ %.2f', x, grouping=True))

#realiza o loop em cada linha da planilha
for record in df.to_dict(orient="records"):
    if pd.isna(record['PROCESSO']):
        break
    # Escreve o valor do contrato em extenso - não foi possível escrever essa parte fora do loop
    record['POREXTENSO'] = num2words(record['POREXTENSO'], lang='pt_BR', to='currency')

    # Abre o template (modelo) do Word
    doc = DocxTemplate(word_template_path)

    # Preenche o template com os dados do registro
    doc.render(record)

    # Define o nome do arquivo de saída com base nos dados do processo
    output_path = output_dir / f"TC {record['PROCESSON']}.{record['ANOPROCESSO']} - DM - LC - {record['INICIAIS']} - Prescrição.docx"

    # Salva o documento preenchido
    doc.save(output_path)

    # Informa que o documento foi finalizado
    print(f"TC {record['PROCESSON']}/{record['ANOPROCESSO']} - LC - {record['INICIAIS']} - finalizado.")




