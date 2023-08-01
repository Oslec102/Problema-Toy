import numpy as np
from datetime import date
import pandas as pd
from mip import Model, xsum, INTEGER, BINARY, OptimizationStatus
import gurobipy as gp


#Ler o total de colaboradores da Planilha Excel
C_d_t_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=1, names=['Colaborador'])
print("======================================")
print("| C_df                                |")
print("======================================")
print(C_d_t_df)
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

C_d_t_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=3, header=[0])
C_dt_dict = {}
for index, row in C_d_t_df.iterrows():
    dia = row['Dia']
    turno = row['Turno']
    colaboradores = []
    if str(row["Colaboradores Disponiveis"]) != "nan":
        colaboradores = row["Colaboradores Disponiveis"].split("+")
    C_dt_dict[(dia, turno)] = colaboradores

print(C_dt_dict)
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
    d = row['Dia']
    turno = row['Turno']
    
    # Extrai a carga horária máxima
    carga_horaria = row['Carga Horaria Maxima']
    
    # Adiciona a carga horária máxima ao dicionário
    h[(colaborador, d, turno)] = carga_horaria

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

limpeza_df = pd.read_excel('ProblemaToy.xlsx', sheet_name=0, usecols=[0, 3, 4, 5, 6, 7, 8, 9,10])
D = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]

# Criar o parâmetro nc[a, d, t] = número de vezes por dia que a área a tem que ser limpa no turno t
nc = {}
for index, row in limpeza_df.iterrows():
    a = row['Area']
    turno = row['Turno']
    for d in D:
        nc[a, d, turno] = row[d]

# Exibir o parâmetro nc
print("======================================")
print("| Parâmetro nc                       |")
print("======================================")
for key, value in nc.items():
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
    a = row['Area']
    predio = row['Predio']
    ar[a] = predio
print("======================================")
print("| ar                                  |")
print("======================================")
print(ar)
print()

def predios_diferentes(p1, p2):
    return 1 if ar[p1] != ar[p2] else 0

f_df = pd.read_excel('ProblemaToy.xlsx',sheet_name=0, usecols=[0,11])
#f_df_clean = f_df.dropna()
print("======================================")
print("| f_df                                  |")
print("======================================")
print(f_df)
print()




# Definir o número de colaboradores, prédios e combinações de prédios: u(i,j,j2)
num_colaboradores = 5
num_predios = 3

# CRIAR MODELO MIP
model = Model()

# CRIAR VARIÁVEIS DE DECISÃO

x = {(c, a, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{})".format(c, a, d, t))  for a in A_df['Area'].values.tolist()  for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)]}

# Criar variável U
#u = {(c, p1, p2): model.add_var(var_type=BINARY, name="u({},{},{})".format(c, p1, p2)) for c in C_d_t_df[(d,t)] for p1 in listaPredios for p2 in listaPredios}


# DEFINIR A FUNÇÃO OBJETIVO
#model.objective = xsum(u[c, p1, p2] * Distancia_d.loc[p1, p2] for c in C_d_t_df[(d,t)] for p1 in listaPredios for p2 in listaPredios)


#Cada área deve ser alocada ao número requerido de colaboradores em cada turno e dia da semana
for a in A_df['Area'].values.tolist():
    for d in D:
        for t in T_df['Turno'].values.tolist():
            if (a, d, t) in nc:   
                model  += xsum(x[(c,a,d,t)] for c in C_dt_dict[(d,t)]) == nc[(a,d,t)]
            else:
                model += xsum(x[c,a,d,t] for c in C_dt_dict[(d,t)]) == 0


#A soma dos tempos de limpeza e deslocamentos não pode exceder a carga horária diária por turno de cada colaborador
# for d in D:
#         for t in T_df['Turno'].values.tolist():
#                 for c in C_dt_dict[(d,t)]:
#                     if (c, d, t) in h:
#                         model += xsum(x[(c,a,d,t)] * t_a[a] for a in t_a) <= h[(c, d, t)]
                    



                #Tem que considerar o parametro Distancia na soma?
                

#Respeitar as fixações de colaborador para área quando determinado
# for index, row in f_df.iterrows():
#      for d, dia_abreviado in dias_semana.items():
#         for t in T_df['Turno'].values.tolist():
#             c = row['ColaboradorFixo']
#             a = row['Area']
#             if str(c) != "nan":
#                 if (a,d,t) in nc:
#                     if nc[(a,d,t)] > 0:
#                         for colaborador in c.split("+"):
#                             model += x[(c,a,d,t)] >= 1 

            # RESOLVER O MODELO
model.optimize()
model.write('teste.lp')


# IMPRIMIR RESULTADOS
if model.status == OptimizationStatus.OPTIMAL:
    for a in A_df['Area'].values.tolist():
        for d in D:
            for t in T_df['Turno'].values.tolist():
                for c in C_dt_dict[(d,t)]:
                    if x[(c, a,d,t)].x >= 0.99:
                        print(f"x({c}, {a}, {d}, {t}) = 1")
else:
    print("A solução ótima não foi encontrada.")


#Imprimir o resultado da minimização
print("======================================")
print("|   funcao_objetiva                  |")
print("======================================")
print(model.objective)
print(model)



#- nc[a, d, t] = número de colaboradores que limpam conjuntamente a area a no dia d no turno t