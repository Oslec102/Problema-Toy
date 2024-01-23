import numpy as np
from datetime import date
import pandas as pd
from mip import Model, xsum, INTEGER, BINARY, OptimizationStatus

file_name = "ProblemaToy1.xlsx"

# Dias da semana
D = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]

# Ler os turnos da planilha Excel.
T_df = pd.read_excel(file_name,sheet_name=2, names=['Turno'])
Turnos = T_df['Turno'].values.tolist()
print("======================================")
print("| T_df                                  |")
print("======================================")
print(T_df)
print()

#Ler o total de colaboradores da Planilha Excel
C_d_t_df = pd.read_excel(file_name,sheet_name=1, names=['Colaborador'])
print("======================================")
print("| C_df                                |")
print("======================================")
print(C_d_t_df)
print()

# Ler as áreas da planilha Excel.
A_df = pd.read_excel(file_name,sheet_name=0)
print(A_df)

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
    for t in Turnos:
        if (d,t) not in A_d_t:
            A_d_t[(d,t)] = []
print(A_d_t)
A_d_t_DUMMY_FIM = {}
A_d_t_DUMMY_INICIO = {}
A_d_t_DUMMY_INICIO_FIM = {}
for (d, t) in A_d_t:
    A_d_t_DUMMY_FIM[(d, t)] = [a for a in A_d_t[(d, t)]]
    A_d_t_DUMMY_INICIO[(d, t)] = [a for a in A_d_t[(d, t)]]
    A_d_t_DUMMY_INICIO_FIM[(d, t)] = [a for a in A_d_t[(d, t)]]
    if  len(A_d_t[(d, t)]) > 0:
        A_d_t_DUMMY_FIM[(d, t)].append("DUMMY_FIM")
        A_d_t_DUMMY_INICIO[(d, t)].append("DUMMY_INICIO")

        A_d_t_DUMMY_INICIO_FIM[(d, t)].append("DUMMY_FIM")
        A_d_t_DUMMY_INICIO_FIM[(d, t)].append("DUMMY_INICIO")

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
print("======================================")
print("Distancia_d                          |")
print("======================================")
print(Distancia_d)
print()

#PARÂMETROS

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

limpeza_df = pd.read_excel(file_name, sheet_name=0, usecols=[0, 3, 4, 5, 6, 7, 8, 9,10])
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
ar["DUMMY_FIM"] = "DUMMY_FIM"
ar["DUMMY_INICIO"] = "DUMMY_INICIO"
print("======================================")
print("| ar                                  |")
print("======================================")
print(ar)
print()

f_df = pd.read_excel(file_name,sheet_name=0, usecols=[0,11])
print("======================================")
print("| f_df                                  |")
print("======================================")
print(f_df)
print()

# CRIAR MODELO MIP
model = Model()
#model = Model(solver_name=CBC)

# CRIAR VARIÁVEIS DE DECISÃO

x = {(c, a, d, t): model.add_var(var_type=BINARY,name="x({},{},{},{})".format(c, a, d, t)) for d in D for t in Turnos for a in A_d_t[(d, t)] for c in C_dt_dict[(d,t)]}

for d in D:
    for t in Turnos:
        for c in C_dt_dict[(d,t)]:
            x[(c,"DUMMY_INICIO",d,t)] =  model.add_var(var_type=BINARY,lb=1.0, ub=1.0, name="x({},{},{},{})".format(c, "DUMMY_INICIO", d, t)) 
            x[(c,"DUMMY_FIM",d,t)] =  model.add_var(var_type=BINARY,lb=1.0, ub=1.0, name="x({},{},{},{})".format(c, "DUMMY_FIM", d, t))
            
y = {(c, a, a2, d, t): model.add_var(var_type=BINARY,name="y({},{},{},{},{})".format(c, a, a2, d, t)) for d in D for t in Turnos for a2 in A_d_t_DUMMY_INICIO_FIM[(d, t)] for a in A_d_t_DUMMY_INICIO_FIM[(d, t)] for c in C_dt_dict[(d,t)]}


u = {(c, a, d, t): model.add_var(var_type=INTEGER,name="u({},{},{},{})".format(c, a, d, t)) for d in D for t in Turnos  for c in C_dt_dict[(d,t)] for a in A_d_t_DUMMY_INICIO_FIM[(d, t)]}


# DEFINIR A FUNÇÃO OBJETIVO
model.objective = xsum(y[(c,a,a2,d,t)] * Distancia_d[ar[a]][ar[a2]] for d in D for t in Turnos for a in A_d_t[(d,t)] for a2 in A_d_t[(d,t)] for c in C_dt_dict[(d,t)])

#Cada área deve ser alocada ao número requerido de colaboradores em cada turno e dia da semana
for d in D:
    for t in Turnos:
        for a in A_d_t[(d, t)]:
            if (a, d, t) in nc:   
                model.add_constr(xsum(x[(c,a,d,t)] for c in C_dt_dict[(d,t)]) == nc[(a,d,t)], name = f"constr2({a},{d},{t})")
            else:
                model.add_constr(xsum(x[(c,a,d,t)] for c in C_dt_dict[(d,t)]) == 0, name = f"constr2({a},{d},{t})")


#A soma dos tempos de limpeza e deslocamentos não pode exceder a carga horária diária por turno de cada colaborador
for d in D:
  for t in Turnos:
      for c in C_dt_dict[(d,t)]:
          if (c, d, t) in h:
            model.add_constr(xsum((x[(c,a,d,t)] * t_a[a]) + xsum(y[(c,a,a2,d,t)] for a2 in A_d_t[(d, t)]) for a in A_d_t[(d, t)]) <= h[(c, d, t)], name = f"constr3({c},{d},{t})")
                

#Respeitar as fixações de colaborador para área quando determinado
for index, row in f_df.iterrows():
    c = row['ColaboradorFixo']
    a = row['Area']
    if str(c) != "nan":
        for d in D:
            for t in Turnos:
                if (a,d,t) in nc and nc[(a,d,t)] > 0:
                    for colaborador in c.split("+"):
                        model.add_constr(x[c,a,d,t] == 1, name = f"constr4({c},{a},{d},{t})")


#Para cada colaborador, dia, turno e área uma área deve ser a próxima a ser limpa
#sucessor:
for d in D:
    for t in Turnos:
       for a in A_d_t_DUMMY_INICIO[(d, t)]:
            for c in C_dt_dict[(d,t)]:
                model.add_constr(xsum(y[(c,a,a2,d,t)] for a2 in A_d_t_DUMMY_FIM[(d, t)] if a!=a2 ) == x[(c,a,d,t)],name = f"constr5({c},{a},{d},{t})")

#antecessor:
for d in D:
    for t in Turnos:
       for a2 in A_d_t_DUMMY_FIM[(d, t)]:
            for c in C_dt_dict[(d,t)]:
                model.add_constr(xsum(y[(c,a,a2,d,t)] for a in A_d_t_DUMMY_INICIO[(d, t)] if a!=a2 ) == x[(c,a2,d,t)],name = f"constr6({c},{a2},{d},{t})")


# #Vincular as váriaveis X as váriaveis Y
# for d in D:
#     for t in Turnos:
#         for a in A_d_t_1[(d, t)]:
#             for c in C_dt_dict[(d,t)]:
#                 model.add_constr(x[c,a,d,t] <= xsum(y[(c,a,a2,d,t)] for a2 in A_d_t_0[(d, t)] if a!=a2 ), name = f"constr7({c},{a},{d},{t})")
                

# for d in D:
#     for t in Turnos:
#         for a in A_d_t[(d,t)]:
# #       for a in A_d_t_DUMMY_INICIO[(d, t)]:
#             for a2 in A_d_t[(d,t)]:
# #           for a2 in A_d_t_DUMMY_FIM[(d, t)]:
#                 if a!=a2:
#                     for c in C_dt_dict[(d,t)]:
#                         print(u[(c,a,d,t)])


# # #Eliminando sub-rotas
for d in D:
    for t in Turnos:
        for c in C_dt_dict[(d,t)]:
            model.add_constr(u[(c,"DUMMY_INICIO",d,t)] == 1, name=f"constr7a({c},{d},{t})")

for d in D:
    for t in Turnos:
        for a in A_d_t[(d, t)]:
            for a2 in A_d_t_DUMMY_FIM[(d, t)]:
                if a!=a2:
                    for c in C_dt_dict[(d,t)]:
                        model.add_constr(u[(c,a,d,t)] - u[(c,a2,d,t)] + (len(A_d_t[(d,t)]) * y[(c,a,a2,d,t)]) <= (len(A_d_t[(d,t)]) - 1), name=f"constr7b({c},{a},{a2},{d},{t})")


# for d in D:
#     for t in Turnos:
#         # for a in A_d_t[(d,t)]:
#         for a in A_d_t_DUMMY_INICIO[(d, t)]:
#             # for a2 in A_d_t[(d,t)]:
#             for a2 in A_d_t_DUMMY_FIM[(d, t)]:
#                 if a!=a2:
#                     for c in C_dt_dict[(d,t)]:
#                         model.add_constr(u[(c,a,d,t)] - u[(c,a2,d,t)] + (len(A_d_t[(d,t)]) - 1) * y[(c,a,a2,d,t)] + (len(A_d_t[(d,t)]) - 3) * y[(c,a2,a,d,t)] <= (len(A_d_t[(d,t)]) - 2), name=f"constr7({c},{a},{a2},{d},{t})")

# for i in range(1, n):
#              for j in range(1, n):
#                  if i != j:
#                      model.add_constr(u[i] - u[j] + ((n - 1) * x[i, j]) + ((n - 3) * x[j, i]) <= n - 2, name=f"cons9_{i}_{j}")

model.write('teste.lp')

# RESOLVER O MODELO
model.optimize()

# IMPRIMIR RESULTADOS
if model.status == OptimizationStatus.OPTIMAL:
    model.write('teste.sol')
    print(f"Custo da solucao: {model.objective_value}")
    for d in D:
        for t in Turnos:
            for a in A_d_t[(d, t)]:
                for c in C_dt_dict[(d,t)]:
                    if x[(c,a,d,t)].x >= 0.99:
                        print(f"x({c}, {a}, {d}, {t}) = 1")
    
    #testando solucao:
    print("\n\n--------------------------\n\n")
    for a in A_d_t_DUMMY_INICIO[("Seg", "M6")]:
        if ("Maria",a,"Seg", "M6") in x and x[("Maria",a,"Seg", "M6")].x >= 0.99:
            print(f"x[(Maria,{a},Seg,M6)] = 1")
    print()
    for a in A_d_t_DUMMY_INICIO[("Seg", "M6")]:
        for a2 in A_d_t_DUMMY_FIM[("Seg", "M6")]:
            if ("Maria",a,a2,"Seg", "M6") in y and y[("Maria",a,a2,"Seg", "M6")].x >= 0.99:
                print(f"y[(Maria,{a},{a2},Seg,M6)] = 1")
else:
    print("A solução ótima não foi encontrada.")


