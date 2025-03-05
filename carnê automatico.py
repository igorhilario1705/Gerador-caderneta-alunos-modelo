import pandas as pd
import os
import shutil

"""###############################################################################################################
                                                  Extrair
###############################################################################################################"""

lista_dos_alunos = pd.read_excel("Alunos dos alunos da escola.xlsx") # lista das informações dos alunos
df = pd.read_excel("Modelo_carne.xlsx") # modelo de carnê de pagamento

"""###############################################################################################################
                                                  Transformar
###############################################################################################################"""

# Criar pasta principal dos resultados
os.makedirs("Carnês criados")

# criando variaveis da lista dos alunos
Alunos = lista_dos_alunos['Alunos']
Pagador = lista_dos_alunos['Responsáveis']
Serie = lista_dos_alunos['Turmas']
Mensalidade = lista_dos_alunos['Mensalidades']
Mensalidade = list(map(lambda x: "R$ {:.2f}".format(x), Mensalidade))

# Criar lista com todas as combinações de alunos e pagadores
combinaçao = zip(Alunos, Pagador, Serie, Mensalidade)

# tratando carnê de pagamento
df.fillna((""), inplace=True)
df.drop(["Unnamed: 2", "Unnamed: 5", "Unnamed: 6", "Unnamed: 7"], axis=1, inplace=True)
df.rename(columns={'Unnamed: 4': '', 'Unnamed: 8': ' '}, inplace=True)
data_vencimento = "01/01/2024"
coluna_vencimento = data_vencimento
df.rename(columns={df.columns[1]: data_vencimento}, inplace=True)
df['   '] = '' # Criando nova coluna

# Criando df com as informações
dataframes = []
for aluno, pai, serie, valor in combinaçao:
    df_new = df.copy()
    df_new.loc[df_new["Vencimento"] == "Aluno", coluna_vencimento] = aluno
    df_new.loc[df_new["Vencimento"] == "Pagador", coluna_vencimento] = pai
    df_new.loc[df_new["Vencimento"] == "Série", coluna_vencimento] = serie
    df_new.loc[df_new["Vencimento"] == "Valor Original", coluna_vencimento] = valor
    df_new.loc[df_new["Pagável no setor Financeiro do Colégio"] == "Aluno"  , ''] = (aluno)
    df_new.loc[df_new["Pagável no setor Financeiro do Colégio"] == "Pagador"  , ''] = (pai)

    diretorio = 'Carnês criados'
    arquivo = f"carne {aluno}.{serie}.xlsx"
    caminho_completo = f"{diretorio}/{arquivo}"
    df_new.to_excel(caminho_completo, index=False)


# Criar pasta principal dos resultados
#os.makedirs("Carnês criados")

# Organizar os alunos por série em cada pasta
# Cria listas para cada turma de arquivos
alunos_Infantil = []
alunos_Fundamental_1 = []
alunos_Fundamental_2 = []
alunos_Ensino_Médio_1 = []
alunos_Ensino_Médio_2 = []
alunos_Ensino_Médio_3 = []

for nome_arquivo in os.listdir(diretorio):

    if 'Infantil' in nome_arquivo:
        alunos_Infantil.append(nome_arquivo)

    elif 'Fundamental 1' in nome_arquivo:
        alunos_Fundamental_1.append(nome_arquivo)

    elif 'Fundamental 2' in nome_arquivo:
        alunos_Fundamental_2.append(nome_arquivo)

    elif 'Ensino Médio 1' in nome_arquivo:
        alunos_Ensino_Médio_1.append(nome_arquivo)

    elif 'Ensino Médio 2' in nome_arquivo:
        alunos_Ensino_Médio_2.append(nome_arquivo)

    elif 'Ensino Médio 3' in nome_arquivo:
        alunos_Ensino_Médio_3.append(nome_arquivo)


# Cria pastas para cada série de arquivos
for serie, arquivos in [('Alunos do Infantil', alunos_Infantil),
                        ('Alunos do Fundamental 1', alunos_Fundamental_1),
                        ('Alunos do Fundamental 2', alunos_Fundamental_2),
                        ('Alunos do Ensino Médio 1', alunos_Ensino_Médio_1),
                        ('Alunos do Ensino Médio 2', alunos_Ensino_Médio_2),
                        ('Alunos do Ensino Médio 3', alunos_Ensino_Médio_3)]:

    if len(arquivos) > 0:
        # Cria a pasta para a série
        pasta_serie = os.path.join( diretorio, serie)

        if not os.path.isdir(pasta_serie):
            os.mkdir(pasta_serie)

        # Move cada arquivo para a pasta da série
        for nome_arquivo in arquivos:
            caminho_arquivo = os.path.join(diretorio, nome_arquivo)
            caminho_destino = os.path.join(pasta_serie, nome_arquivo)
            shutil.move(caminho_arquivo, caminho_destino)

"""###############################################################################################################
                                                  Carregar
###############################################################################################################"""

# Caminho para a pasta
caminho = 'Carnês criados'

# Abre a pasta no Explorer
os.startfile(caminho)




