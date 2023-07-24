import numpy as np
from datetime import date
import pandas as pd
from mip import Model, xsum, INTEGER, BINARY, OptimizationStatus
import gurobipy as gp


#Ler o total de colaboradores da Planilha Excel
C_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=1, names=['Colaborador'])
print("======================================")
print("| C_df                                |")
print("======================================")
print(C_df)
print()


# Ler as áreas da planilha Excel.
A_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=0, usecols=[0])
print("======================================")
print("| A_df                                  |")
print("======================================")
print(A_df)
print()

# Ler os turnos da planilha Excel.
T_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=2, names=['Turno'])
print("======================================")
print("| T_df                                  |")
print("======================================")
print(T_df)
print()

C_d_t_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=3, names=['ColaboradoresDiaTurno'])
print("======================================")
print("|C_d_t_df                                   |")
print("======================================")
print(C_d_t_df)
print()

Distancia_d = pd.read_excel('ProblemaToy.xlsx',sheet_name=5, header=[0], index_col=0)


print("======================================")
print("Distancia_d                                  |")
print("======================================")
print(Distancia_d)
print()

'''
#Dicionário com os dias da semana.
#D = ("Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom")
D = pd.read_excel('ProblemaToy.xlsx',sheet_name=0, usecols=[4,5,6,7,8,9])
print("======================================")
print("| D                                  |")
print("======================================")
print(D)
print()
'''
# Dicionário com os dias da semana.
# D = ("Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom")
D = pd.read_excel('ProblemaToy.xlsx', sheet_name=0, usecols=[4, 5, 6, 7, 8, 9])

# Converter DataFrame para uma lista de listas com os dias completos
D = D.values.tolist()

#PARÂMETROS

# h[c, d, t] = carga horária máxima (em minutos) do colaborador c no dia d no turno t
dadosminutos = pd.read_excel('ProblemaToy.xlsx',sheet_name=4)
# Cria o dicionário
h = {}

# Itera sobre as linhas do DataFrame
for index, row in dadosminutos.iterrows():
    # Extrai as informações de colaborador, dia e turno
    colaborador = row['Colaborador']
    dia = row['Dia']
    turno = row['Turno']
    
    # Extrai a carga horária máxima
    carga_horaria = row['Carga Horaria Maxima']
    
    # Adiciona a carga horária máxima ao dicionário
    h[(colaborador, dia, turno)] = carga_horaria

# Exibe o dicionário
print("======================================")
print("| h                                  |")
print("======================================")
print(h)
print()

# t[a] = tempo (em minutos) que a área a leva para ser limpada
df_tempoLimpeza = pd.read_excel('ProblemaToy.xlsx')

# Cria um dicionário com o tempo que cada área leva para ser limpa
t_a = dict(zip(df_tempoLimpeza["Area"], df_tempoLimpeza["Tempo_limpeza_em_minutos"]))
print("======================================")
print("| t_a                                  |")
print("======================================")
print(t_a)
print()

#ld[a, d, t] = número de vezes por dia que a área a tem que ser limpa no turno t 
#ld = pd.read_excel('ProblemaToy.xlsx',sheet_name=0, usecols=[0,3,4,5,6,7,8,9])
#print("======================================")
#print("| ld[a, d, t]                                  |")
#print("======================================")
#print(ld)
#print()

limpeza_df = pd.read_excel('ProblemaToy.xlsx', sheet_name=0, usecols=[0, 3, 4, 5, 6, 7, 8, 9])

# Criar um dicionário que mapeia os dias da semana para suas abreviações
dias_semana = {
    'Seg': 'Seg',
    'Ter': 'Ter',
    'Qua': 'Qua',
    'Qui': 'Qui',
    'Sex': 'Sex',
    'Sab': 'Sab'
}

# Criar o parâmetro ld[a, d, t] = número de vezes por dia que a área a tem que ser limpa no turno t
ld = {}
for index, row in limpeza_df.iterrows():
    area = row['Area']
    turno = row['Turno']
    for dia, dia_abreviado in dias_semana.items():
        ld[area, dia_abreviado, turno] = row[dia]

# Exibir o parâmetro ld
print("======================================")
print("| Parâmetro ld                       |")
print("======================================")
for key, value in ld.items():
    print(f"{key}: {value}")
print()


df_predioArea = pd.read_excel('ProblemaToy.xlsx')
# Criando um dicionário que associa cada área ao prédio correspondente
ar = {}
listaPredios = [*set(df_predioArea.loc[:,"Predio"].tolist())]
print("======================================")
print("| lista Predio                        |")
print("======================================")
print(listaPredios)
print()

for index, row in df_predioArea.iterrows():
    area = row['Area']
    predio = row['Predio']
    ar[area] = predio
print("======================================")
print("| ar                                  |")
print("======================================")
print(ar)
print()

def predios_diferentes(p1, p2):
    return 1 if ar[p1] != ar[p2] else 0

f_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=0, usecols=[10])
f_df_clean = f_df.dropna()
print("======================================")
print("| f_df                                  |")
print("======================================")
print(f_df_clean)
print()




# Definir o número de colaboradores, prédios e combinações de prédios: u(i,j,j2)
num_colaboradores = 5
num_predios = 3

# CRIAR MODELO MIP
model = Model()

# CRIAR VARIÁVEIS DE DECISÃO
x = {(c, a): model.add_var(var_type=INTEGER,name="x({},{})".format(c, a)) for c in C_df['Colaborador'].values.tolist() for a in A_df['Area'].values.tolist()}

# Criar variável U
u = {(c, p1, p2): model.add_var(var_type=BINARY, name="u({},{},{})".format(c, p1, p2)) for c in C_df['Colaborador'].values.tolist() for p1 in listaPredios for p2 in listaPredios}


# DEFINIR A FUNÇÃO OBJETIVO
model.objective = xsum(u[c, p1, p2] * Distancia_d.loc[p1, p2] for c in C_df['Colaborador'].values.tolist() for p1 in listaPredios for p2 in listaPredios)


#Cada área deve ser alocada ao número requerido de colaboradores em cada turno e dia da semana
for a in A_df['Area'].values.tolist():
    for dia, dia_abreviado in dias_semana.items():
        for t in T_df['Turno'].values.tolist():
            if (a, dia, t) in ld:   
                model  += xsum(x[(c,a)] for c in C_df['Colaborador'].values.tolist()) == ld[(a,dia,t)]


#A soma dos tempos de limpeza e deslocamentos não pode exceder a carga horária diária por turno de cada colaborador
for dia, dia_abreviado in dias_semana.items():
    for t in T_df['Turno'].values.tolist():
        for p1 in listaPredios:
            for p2 in listaPredios:
                for d in Distancia_d:
                    for c in C_df['Colaborador'].values.tolist():
                        if (colaborador, dia, turno) in h:
                            model += xsum(x[(c,a)] for c in C_df['Colaborador'].values.tolist() * t[(a)] * Distancia_d.loc[p1, p2]) <= h[(colaborador, dia, turno)]

  #Tem que considerar o parametro Distancia na soma?
                

# Respeitar as fixações de colaborador para área quando determinado
for c in C_df['Colaborador'].values.tolist():
    for a in A_df['Area'].values.tolist():
        if (c, a) in f_df_clean:
            f_fixo = f_df_clean.loc[(c, a)].values[0]
            model += x[(c, a)] >= f_fixo

# RESOLVER O MODELO
model.optimize()
model.write('teste.lp')


# IMPRIMIR RESULTADOS
if model.status == OptimizationStatus.OPTIMAL:
    for c in C_df['Colaborador'].values.tolist():
        for p1 in listaPredios:
            for p2 in listaPredios:
                if u[c, p1, p2].x >= 0.99:
                    print(f"O colaborador {c} está designado para o prédio {p1}.")
else:
    print("A solução ótima não foi encontrada.")


#Imprimir o resultado da minimização
print("======================================")
print("|   funcao_objetiva                  |")
print("======================================")
print(model.objective)
print(model)



