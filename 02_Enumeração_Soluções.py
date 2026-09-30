# =========================================================
# EXPLICATIVO RESUMIDO
# =========================================================

# Este script tem como finalidade gerar, por enumeração,
# todas as combinações possíveis de mercados e todas as
# respectivas ordens de visita.

# O resultado representa o espaço completo de soluções
# possíveis para os mercados disponíveis.

# A base gerada neste script NÃO realiza a seleção da
# melhor solução. Sua finalidade é fornecer uma referência
# exaustiva para posterior comparação e validação dos
# resultados encontrados pelo modelo de seleção e
# otimização.


#%%
# =========================================================
# BIBLIOTECAS
# =========================================================

from itertools import combinations, permutations
import pandas as pd
from pathlib import Path


# =========================================================
# CAMINHO DO ARQUIVO EXCEL
# =========================================================

from config import BASE_FONTE, ENUMERACAO
caminho_excel = BASE_FONTE
arquivo = ENUMERACAO


# =========================================================
# LEITURA DA BASE DE MERCADOS
# =========================================================

# Lê a aba "Base Mercado", que contém os mercados
# considerados no estudo.
df_mercados = pd.read_excel(
    caminho_excel,
    sheet_name="Base Mercado")


# =========================================================
# DEFINIÇÃO DOS MERCADOS
# =========================================================

# Obtém a lista de mercados disponíveis na base.

# drop_duplicates()
# → garante que cada mercado seja considerado apenas uma vez.

# sort_values()
# → organiza os mercados em ordem alfabética,
#   facilitando a padronização da enumeração.

# tolist()
# → converte o resultado para uma lista Python.
mercados = (
    df_mercados["Nome"]
    .drop_duplicates()
    .sort_values()
    .tolist())


# =========================================================
# INICIALIZAÇÃO DA ENUMERAÇÃO
# =========================================================

# Lista que armazenará todas as soluções geradas.
cenarios = []

# Contador utilizado para identificar individualmente
# cada cenário/rota gerado.
id_cenario = 1

# Define a quantidade máxima de mercados que poderá
# participar de uma solução.

# Como o valor é obtido diretamente da quantidade de
# mercados disponíveis, a enumeração considera desde
# soluções com apenas 1 mercado até soluções contendo
# todos os mercados.
MAX_MERCADOS = len(mercados)


# =========================================================
# ENUMERAÇÃO DAS SOLUÇÕES
# =========================================================

# Percorre todas as quantidades possíveis de mercados:

# 1 mercado
# 2 mercados
# 3 mercados

# até todos os mercados disponíveis.
for qtd in range(1, MAX_MERCADOS + 1):

    # -----------------------------------------------------
    # COMBINAÇÕES DE MERCADOS
    # -----------------------------------------------------
    
    # Para cada quantidade de mercados, gera todas as
    # combinações possíveis.
    
    # Exemplo:
    
    # Mercados = A, B, C
    
    # Para 2 mercados:
    # (A,B) e também (B,A) pois a ordem dos mercados pode afetar o custo total final
    # (A,C)
    # (B,C)
    
    # Neste momento, a ordem dos mercados ainda não é
    # considerada.
    for grupo in combinations(mercados, qtd):


        # -------------------------------------------------
        # PERMUTAÇÕES / ORDENS DE VISITA
        # -------------------------------------------------
        
        # Para cada combinação de mercados, são geradas
        # todas as ordens possíveis de visita.
        
        # Exemplo:
        
        # Combinação = (A,B,C)
        
        # Possíveis sequências:
        # A → B → C
        # A → C → B
        # B → A → C
        # B → C → A
        # C → A → B
        # C → B → A
        
        # Portanto, o script considera tanto:
        
        # 1. quais mercados serão visitados;
        # 2. em qual ordem eles serão visitados.
        
        # Esta enumeração exaustiva representa o espaço
        # completo de soluções utilizado posteriormente
        # como referência para validação do modelo.
        for rota in permutations(grupo):


            # =============================================
            # ESTRUTURA DO CENÁRIO
            # =============================================

            # Cria um registro para armazenar uma solução.
            
            # ID_CENARIO
            # → identificador único da solução.
            
            # QTD_MERCADOS
            # → quantidade de mercados presentes na solução.
            
            # SAIDA
            # → define o ponto inicial da rota como "Casa".
            linha = {
                "ID_CENARIO": id_cenario,
                "QTD_MERCADOS": qtd,
                "SAIDA": "Casa"}


            # =============================================
            # PREENCHIMENTO DA SEQUÊNCIA DE MERCADOS
            # =============================================

            # Cria as colunas M1, M2, ..., MN.
            
            # Quando a rota possui determinada posição,
            # o mercado correspondente é armazenado na coluna.
            
            # Caso a rota tenha menos mercados que o número
            # máximo disponível, as posições restantes são
            # preenchidas com texto vazio.
            
            # Isso mantém uma estrutura padronizada para
            # todas as soluções, independentemente da
            # quantidade de mercados visitados.
            for i in range(MAX_MERCADOS):
                if i < len(rota):
                    linha[f"M{i+1}"] = rota[i]
                else:
                    linha[f"M{i+1}"] = ""


            # =============================================
            # PONTO DE CHEGADA
            # =============================================

            # Define "Casa" como o ponto final da rota.
            
            # Dessa forma, cada solução possui a estrutura:
            
            # Casa → Mercado(s) → Casa
            linha["CHEGADA"] = "Casa"


            # =============================================
            # ARMAZENAMENTO DO CENÁRIO
            # =============================================

            # Adiciona a solução completa à lista de cenários.
            cenarios.append(linha)

            # Incrementa o identificador para que a próxima
            # solução receba um ID diferente.
            id_cenario += 1


# =========================================================
# CONSOLIDAÇÃO DOS RESULTADOS
# =========================================================

# Converte a lista de cenários em um DataFrame.

# Cada linha representa uma solução possível,
# contendo determinada combinação de mercados
# e determinada ordem de visita.
Todas_Rotas = pd.DataFrame(cenarios)


# Exibe a quantidade total de soluções geradas.

# Esse número representa o tamanho do espaço de soluções
# enumerado pelo script para a quantidade de mercados
# disponível.
print(f"Total de rotas geradas: {len(Todas_Rotas)}")


# =========================================================
# EXPORTAÇÃO
# =========================================================


# Exporta todas as soluções para Excel.

# index=False
# → evita que o índice interno do DataFrame seja criado
#   como uma coluna adicional no arquivo.
Todas_Rotas.to_excel(arquivo, index=False)


# Confirma no console o local onde o arquivo foi salvo.
print(f"\nArquivo salvo em:\n{arquivo}")
