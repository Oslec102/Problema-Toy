import numpy as np
from datetime import date
import pandas as pd
from mip import Model, xsum, INTEGER, BINARY, OptimizationStatus
import gurobipy as gp

file_name = "ProblemaToy1.xlsx"

#Ler o total de colaboradores da Planilha Excel
C_d_t_df = pd.read_excel(file_name,sheet_name=1, names=['Colaborador'])
print("======================================")
print("| C_df                                |")
print("======================================")
print(C_d_t_df)
print()


# Ler as áreas da planilha Excel.
<<<<<<< HEAD
A_df = pd.read_excel(file_name,sheet_name=0)
print(A_df)
# A = A_df["Area"].values.tolist()
# A0 = A_df["Area"].values.tolist()
# A1 = A_df["Area"].values.tolist()
# A0.append("DUMMY_FIM")
# A1.append("DUMMY_INICIO")
# print("======================================")
# print("| A                                  |")
# print("======================================")
# print(A)
# print()

# # Ler as áreas da planilha Excel.
# print("======================================")
# print("| A0                                  |")
# print("======================================")
# print(A0)
# print()

# # Ler as áreas da planilha Excel.
# print("======================================")
# print("| A1                                  |")
# print("======================================")
# print(A1)
# print()
=======
A_df = pd.read_excel(file_name,sheet_name=0, usecols=[0])
A = A_df["Area"].values.tolist()
A0 = A_df["Area"].values.tolist()
A0.append("DUMMY")
print("======================================")
print("| A                                  |")
print("======================================")
print(A)
print()

# Ler as áreas da planilha Excel.
print("======================================")
print("| A0                                  |")
print("======================================")
print(A0)
print()
>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007

# Ler os turnos da planilha Excel.
T_df = pd.read_excel(file_name,sheet_name=2, names=['Turno'])
print("======================================")
print("| T_df                                  |")
print("======================================")
print(T_df)
print()

<<<<<<< HEAD
# Dicionário com os dias da semana.
D = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
#D = pd.read_excel(file_name, sheet_name=0, usecols=[4, 5, 6, 7, 8, 9,10])

# Converter DataFrame para uma lista de listas com os dias completos
#D = D.values.tolist()

A_d_t = {}
for index, row in A_df.iterrows():
    area = row['Area']
    turno = row['Turno']
    for d in D:
        x = row[d]
        print(row[d])
        if int(row[d]) != 0:
            if (d, turno) not in A_d_t:
                A_d_t[(d, turno)] = [area]
            else:
                A_d_t[(d, turno)].append(area)
       
for d in D:
    for t in T_df['Turno'].values.tolist():
        if (d,t) not in A_d_t:
            A_d_t[(d,t)] = []
print(A_d_t)
A_d_t_0 = {}
A_d_t_1 = {}
for (d, t) in A_d_t:
    A_d_t_0[(d, t)] = [a for a in A_d_t[(d, t)]]
    A_d_t_1[(d, t)] = [a for a in A_d_t[(d, t)]]
    if  len(A_d_t[(d, t)]) > 0:
        A_d_t_0[(d, t)].append("DUMMY_FIM")
    
        A_d_t_1[(d, t)].append("DUMMY_INICIO")


=======
>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
C_d_t_df = pd.read_excel(file_name,sheet_name=3, header=[0])
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

Distancia_d = pd.read_excel(file_name,sheet_name=5, index_col=0)
<<<<<<< HEAD
print(Distancia_d["Agua e Ar"]["Agua e Ar"])
=======
print(Distancia_d)
print(Distancia_d["Agua e Ar"]["Agua e Ar"])
# Dist = {}

# for index, row in Distancia_d.iterrows():

#     p1 = row ['predio 1']
#     p2 = row ['predio 2']

#     Dist[( p1, p2)] = Distancia_d

>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
print("======================================")
print("Distancia_d                          |")
print("======================================")
print(Distancia_d)
print()

<<<<<<< HEAD
=======
# Dicionário com os dias da semana.
# D = ("Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom")
D = pd.read_excel(file_name, sheet_name=0, usecols=[4, 5, 6, 7, 8, 9])

# Converter DataFrame para uma lista de listas com os dias completos
D = D.values.tolist()

>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
#PARÂMETROS

# h[c, d, t] = carga horária máxima (em minutos) do colaborador c no dia d no turno t
dadosminutos = pd.read_excel(file_name,sheet_name=4)
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
df_tempoLimpeza = pd.read_excel(file_name)

# Cria um dicionário com o tempo que cada área leva para ser limpa
t_a = dict(zip(df_tempoLimpeza["Area"], df_tempoLimpeza["Tempo_limpeza_em_minutos"]))
print("======================================")
print("| t_a                                  |")
print("======================================")
print(t_a)
print()

<<<<<<< HEAD
limpeza_df = pd.read_excel(file_name, sheet_name=0, usecols=[0, 3, 4, 5, 6, 7, 8, 9,10])
#D ={"Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"}
D = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
=======
#ld[a, d, t] = número de vezes por dia que a área a tem que ser limpa no turno t 
#ld = pd.read_excel(file_name,sheet_name=0, usecols=[0,3,4,5,6,7,8,9])
#print("======================================")
#print("| ld[a, d, t]                                  |")
#print("======================================")
#print(ld)
#print()

limpeza_df = pd.read_excel(file_name, sheet_name=0, usecols=[0, 3, 4, 5, 6, 7, 8, 9,10])
D ={"Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"}

>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
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


df_predioArea = pd.read_excel(file_name)
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

f_df = pd.read_excel(file_name,sheet_name=0, usecols=[0,11])
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
#model = Model(solver_name=CBC)

# CRIAR VARIÁVEIS DE DECISÃO

<<<<<<< HEAD
x = {(c, a, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{})".format(c, a, d, t)) for d in D for t in T_df['Turno'].values.tolist() for a in A_d_t_1[(d, t)] for c in C_dt_dict[(d,t)]}

y = {(c, a, a2, d, t): model.add_var(var_type=BINARY,name="y({},{},{},{},{})".format(c, a, a2, d, t)) for d in D for t in T_df['Turno'].values.tolist() for a2 in A_d_t_0[(d, t)] for a in A_d_t_1[(d, t)] for c in C_dt_dict[(d,t)]}
#y = {(c, a, a2, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{},{})".format(c, a, a2, d, t))  for a in A for a2 in A0 for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)]}

# DEFINIR A FUNÇÃO OBJETIVO
#model.objective = xsum(y[(c,a,a2,d,t)] * Distancia_d[ar[a]][ar[a2]] for d in D for t in T_df['Turno'].values.tolist() for a in A_d_t for a2 in A_d_t_0 for c in C_dt_dict[(d,t)])

#Cada área deve ser alocada ao número requerido de colaboradores em cada turno e dia da semana
for d in D:
    for t in T_df['Turno'].values.tolist():
        for a in A_d_t[(d, t)]:
=======
x = {(c, a, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{})".format(c, a, d, t))  for a in A  for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)]}

y = {(c, a, a2, d, t): model.add_var(var_type=BINARY,name="y({},{},{},{},{})".format(c, a, a2, d, t))  for a in A for a2 in A0 for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)]}
#y = {(c, a, a2, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{},{})".format(c, a, a2, d, t))  for a in A for a2 in A0 for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)]}

# DEFINIR A FUNÇÃO OBJETIVO
model.objective = xsum(y[(c,a,a2,d,t)] * Distancia_d[ar[a]][ar[a2]] for a in A for a2 in A for d in D for t in T_df['Turno'].values.tolist() for c in C_dt_dict[(d,t)])

#Cada área deve ser alocada ao número requerido de colaboradores em cada turno e dia da semana
for a in A:
    for d in D:
        for t in T_df['Turno'].values.tolist():
>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
            if (a, d, t) in nc:   
                model  += xsum(x[(c,a,d,t)] for c in C_dt_dict[(d,t)]) == nc[(a,d,t)]
            else:
                model += xsum(x[(c,a,d,t)] for c in C_dt_dict[(d,t)]) == 0


#A soma dos tempos de limpeza e deslocamentos não pode exceder a carga horária diária por turno de cada colaborador
<<<<<<< HEAD
# for d in D:
#   for t in T_df['Turno'].values.tolist():
#       for c in C_dt_dict[(d,t)]:
#           for a in A_d_t[(d, t)]:
#                 if (c, d, t) in h:
#                       model += xsum((x[(c,a,d,t)] * t_a[a]) + y[(c,a,a2,d,t)] for a in t_a for a2 in A_d_t_1[(d, t)]) <= h[(c, d, t)]
#KEYERROR - NÃO ESTA BUSCANDO A AREA A_D_T_1        
                

# #Respeitar as fixações de colaborador para área quando determinado
# for index, row in f_df.iterrows():
#      for d in D:
#         for t in T_df['Turno'].values.tolist():
#             for a in A_d_t[(d, t)]:
#                 for c in C_dt_dict[(d, t)]:
#                     c = row['ColaboradorFixo']
#                     a = row['Area']
#                     if str(c) != "nan":
#                         if (a,d,t) in nc:
#                             if nc[(a,d,t)] > 0:
#                                 for colaborador in c.split("+"):
#                                    model += (x[c,a,d,t]) >= 1
##                                    model += (x[c,a,d,t]) >= f_df[(c,a)]


# # #Para cada colaborador, dia, turno e área uma área deve ser a próxima a ser limpa
for d in D:
    for t in T_df['Turno'].values.tolist():
       for a2 in A_d_t_0[(d, t)]:
            for c in C_dt_dict[(d,t)]:
                      if len(A_d_t[(d, t)]) > 0:
                        model.add_constr(xsum(y[(c,a,a2,d,t)] for a in A_d_t[(d, t)] if a!=a2 ) == 1,name=f"5({d},{t},{a},{a2},{c})")


#
for d in D:
    for t in T_df['Turno'].values.tolist():
        for a in A_d_t_1[(d, t)]:
            for c in C_dt_dict[(d,t)]:
                      if len(A_d_t_0[(d, t)]) > 0:
                        model.add_constr(xsum(y[(c,a,a2,d,t)] for a2 in A_d_t_0[(d, t)] if a!=a2 ) == 1,name=f"6({d},{t},{a},{a2},{c})")





# #Vincular as váriaveis X as váriaveis Y
for d in D:
    for t in T_df['Turno'].values.tolist():
        for a in A_d_t_1[(d, t)]:
            for c in C_dt_dict[(d,t)]:
                model += (x[c,a,d,t]) <= xsum(y[(c,a,a2,d,t)] for a2 in A_d_t_0[(d, t)] if a!=a2 )
=======
for d in D:
  for t in T_df['Turno'].values.tolist():
      for c in C_dt_dict[(d,t)]:
          for a in A:
              for a2 in A:
                  if (c, d, t) in h:
                      model += xsum((x[(c,a,d,t)] * t_a[a]) + y[(c,a,a2,d,t)] for a in t_a) <= h[(c, d, t)]
        
                

#Respeitar as fixações de colaborador para área quando determinado
for index, row in f_df.iterrows():
     for d in D:
        for t in T_df['Turno'].values.tolist():
            c = row['ColaboradorFixo']
            a = row['Area']
            if str(c) != "nan":
                if (a,d,t) in nc:
                    if nc[(a,d,t)] > 0:
                        for colaborador in c.split("+"):
                            model += (x[c,a,d,t]) >= 1


# #Para cada colaborador, dia, turno e área uma área deve ser a próxima a ser limpa
for a in A:
    for d in D:
        for t in T_df['Turno'].values.tolist():
            for c in C_dt_dict[(d,t)]:
                      model += xsum(y[(c,a,a2,d,t)] for a2 in A0 if a!=a2 ) == 1

#Vincular as váriaveis X as váriaveis Y
for a in A:
    for d in D:
        for t in T_df['Turno'].values.tolist():
            for c in C_dt_dict[(d,t)]:
                model += (x[c,a,d,t]) <= xsum(y[(c,a,a2,d,t)] for a2 in A0 if a!=a2 )
>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007


            # RESOLVER O MODELO
model.optimize()
model.write('teste.lp')
model.write('teste.sol')


# IMPRIMIR RESULTADOS
if model.status == OptimizationStatus.OPTIMAL:
<<<<<<< HEAD
    for d in D:
        for t in T_df['Turno'].values.tolist():
            for a in A_d_t[(d, t)]:
=======
    for a in A:
        for d in D:
            for t in T_df['Turno'].values.tolist():
>>>>>>> 752e5a034ce8d639e19b57fac751a9e0c0766007
                for c in C_dt_dict[(d,t)]:
                    if x[(c,a,d,t)].x >= 0.99:
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