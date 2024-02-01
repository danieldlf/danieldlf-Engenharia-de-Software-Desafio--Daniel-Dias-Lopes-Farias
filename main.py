import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# Configurar as credenciais
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Abrir a planilha pelo ID
spreadsheet = client.open_by_key('13SHD-Gw7zW9q-6wo8_WzVDu0drWdPPwViDAjxmLbwVY')
worksheet = spreadsheet.get_worksheet(0)

# Obter valores da planilha
data = worksheet.get_values()

# Separar valores, colunas e informações importantes da planilha
number_of_classes = int(data[1][0][-2:]) # Número de aulas
columns = data[2] # Colunas
info = data[3:] # Informações das colunas

# Dicionário utilizado para armazenar os dados da planilha
spreadsheet_dict = dict()

# Índice utilizado na lógica para pegar as informações corretamente da planilha
index = 0

# Pegando as informações de cada coluna e colocando no dicionário
for colum in columns:
    row = list()
    for student in info:
        row.append(student[index])
    spreadsheet_dict[colum] = row
    index += 1

# Criação do Pandas DataFrame a partir do dicionário
df = pd.DataFrame.from_dict(spreadsheet_dict)

# Settando as colunas de notas e faltas como numéricas para permitir operações algébricas
df['P1'] = pd.to_numeric(df['P1'])
df['P2'] = pd.to_numeric(df['P2'])
df['P3'] = pd.to_numeric(df['P3'])
df['Faltas'] = pd.to_numeric(df['Faltas'])

# Cálculo da média de cada aluno e arredondamento -> Cria-se uma nova coluna para armazenar esses dados no dataframe
df['Média'] = round((df['P1'] + df['P2'] + df['P3']) / 3)

# Função utilizada para definir a situação do aluno a partir de sua média
def mean_situation(mean):

    if mean < 50:
        return 'Reprovado por Nota'
    if mean >= 50 and mean < 70:
        return 'Exame Final'
    if mean >= 70:
        return 'Aprovado'

# Função utilizada para definir a situação do aluno a partir de suas faltas
def absences_situation(absences):

    limit = number_of_classes/4 # 25%

    if absences > limit:
        return 'Reprovado por Falta'
    else:
        return None

# Normalizar os valores da coluna
def normalize_colum(value):

    if 'Reprovado por Falta' in value:
        return 'Reprovado por Falta'
    else:
        return value.strip()

# Função utilizada para definir a nota final necessária para aprovação do aluno
def final(mean):

    final_mean = 100-mean
    return final_mean

# Aplicar as funções criadas nas colunas Média e Faltas -> Resultado é passado para novas colunas <Situação Notas>, <Situação Faltas>
df['Situação Notas'] = df['Média'].apply(mean_situation)
df['Situação Faltas'] = df['Faltas'].apply(absences_situation)

# Juntar os resultados calculados anteriormente na coluna de Situação -> Remover os NaN da tabela
df['Situação'] = df[['Situação Notas', 'Situação Faltas']].fillna('').agg(' '.join, axis=1)

# Necessário normalizar a coluna baseado nos seus valores -> Remover o espaço final do agg e também definir
# qual o valor correto para as linhas que tiveram dois valores concatenados
df['Situação'] = df['Situação'].apply(normalize_colum)

# Aplicar a função de cálculo da nota final apenas nos valores que estão em linhas na qual a situação é de Exame Final
df['Nota para Aprovação Final'] = df.loc[df['Situação'] == 'Exame Final', 'Média'].apply(final)

# Trocar os NaN por 0
df['Nota para Aprovação Final'] = df['Nota para Aprovação Final'].fillna(0)

# Remover as tabelas utilizadas apenas para cálculos -> Situação Notas, Situação Faltas, Média
df = df.drop(['Situação Notas', 'Situação Faltas', 'Média'], axis=1)

# Criar um dataframe apenas com as colunas que serão atualizadas no Google Sheets (Situação e Nota para Aprovação Final)
df_updated = df[['Situação', 'Nota para Aprovação Final']]

# Transformar os valores do dataframe (sem as colunas) em uma lista
data_updated =  df_updated.values.tolist()

# Definir o intervalo de células que será atualizado no Google Sheets
start_cell = 'G4'
end_cell = 'H27'
cell_range = f'{start_cell}:{end_cell}'

# Atualizar o Google Sheets
worksheet.update(data_updated, cell_range)