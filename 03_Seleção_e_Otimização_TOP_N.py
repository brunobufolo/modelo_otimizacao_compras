# =========================================================
# EXPLICATIVO RESUMIDO
# =========================================================
#
# Este script concentra o cálculo principal proposto no trabalho.
#
# Sua função é:
#
# 1. preparar a base de compras e preços;
# 2. identificar o mercado inicial de referência (M1);
# 3. avaliar, de forma incremental, a inclusão de novos mercados;
# 4. selecionar mercados enquanto houver saldo econômico positivo;
# 5. após a seleção, testar as possíveis sequências de visita entre
#    os mercados escolhidos;
# 6. determinar a sequência de menor custo de deslocamento;
# 7. calcular o custo total da solução multimercado;
# 8. comparar a solução multimercado com a compra integral em M1;
# 9. selecionar como resultado final a alternativa de menor custo;
# 10. consolidar os resultados econômicos e operacionais por lista.
#
# IMPORTANTE:
# O Script 02 possui finalidade distinta. Ele enumera o espaço de soluções
# possíveis e serve como referência para validação.
#
# Neste Script 03, a seleção dos mercados ocorre por uma regra incremental
# de saldo econômico. A enumeração por permutações é utilizada somente
# depois da seleção dos mercados, com o objetivo de determinar a melhor
# sequência de visita.
#
# REGRA DE DISPONIBILIDADE:
#
# Para concorrer como solução de compra única (M1), o mercado precisa
# possuir todos os produtos necessários daquela lista.
#
# Na solução multimercado, essa exigência não se aplica.
# Um mercado pode ser incorporado mesmo que não possua todos os produtos,
# concorrendo somente pelos itens para os quais possui preço disponível.
#
# Ao final, a solução multimercado obtida pelo algoritmo é confrontada
# com a alternativa de compra integral em M1. A solução final recomendada
# corresponde à alternativa de menor custo total.
#
# =========================================================


#%%
# =========================================================
# SCRIPT 03
# SELEÇÃO DE MERCADOS + OTIMIZAÇÃO DE ROTAS
# =========================================================


# =========================================================
# IMPORTS
# =========================================================

import pandas as pd

# permutations será utilizado apenas na etapa posterior à seleção dos mercados,
# quando todas as ordens possíveis de visita entre os mercados já selecionados
# serão avaliadas para encontrar a sequência de menor custo.
from itertools import permutations


#%%
# =========================================================
# CAMINHO EXCEL
# =========================================================
#
from config import BASE_FONTE, MATRIZ_ROTAS, ESCOLHA_MERCADOS_TOP_N, RESULTADO_OTIMO_TOP_N
caminho_excel = BASE_FONTE
output = ESCOLHA_MERCADOS_TOP_N
caminho_rotas = MATRIZ_ROTAS
output_resultado = RESULTADO_OTIMO_TOP_N

#%%
# =========================================================
# PARÂMETRO AJUSTÁVEL
# =========================================================

TOP_K = 2   # quantidade de candidatos avaliados em paralelo a cada passo




#%%
# =========================================================
# LEITURA DAS ABAS
# =========================================================
#
# "Compras":
# contém as listas que serão processadas, seus produtos, quantidades
# e respectivos pontos de partida.
#
# "Produtos & Mercados":
# contém a disponibilidade/preço dos produtos em cada mercado.

df_compras = pd.read_excel(
    caminho_excel,
    sheet_name="Compras"
)

df_produtos_mercados = pd.read_excel(
    caminho_excel,
    sheet_name="Produtos & Mercados"
)


#%%
# =========================================================
# PADRONIZAÇÃO DAS COLUNAS
# =========================================================
#
# Remove espaços residuais nos nomes das colunas para evitar falhas
# nos filtros, cruzamentos e referências realizadas ao longo do script.

df_compras.columns = (
    df_compras.columns
    .str.strip()
)

df_produtos_mercados.columns = (
    df_produtos_mercados.columns
    .str.strip()
)


#%%
# =========================================================
# BASE DE COMPRAS
# =========================================================
#
# Mantém apenas os campos necessários para identificar:
# - a lista;
# - o produto;
# - a quantidade;
# - o ponto de partida associado àquela lista.

df_compras = df_compras[
    [
        "ID Lista",
        "Item | Marca",
        "ID Produto",
        "Quantidade",
        "Ponto de Partida",
        "ID Ponto de Partida"
    ]
].copy()


#%%
# =========================================================
# BASE DE PRODUTOS E MERCADOS
# =========================================================
#
# Restringe a base de preços apenas aos produtos que aparecem
# nas listas de compras que efetivamente serão avaliadas.

df_produtos_mercados = df_produtos_mercados[
    df_produtos_mercados["ID Produto"].isin(
        df_compras["ID Produto"]
    )
].copy()


#%%
# =========================================================
# CRUZAMENTO:
# COMPRAS x PRODUTOS & MERCADOS
# =========================================================
#
# Expande cada item da lista para os mercados em que esse produto
# possui informação disponível, permitindo comparar seu custo
# entre diferentes estabelecimentos.

df_escolha_mercados = df_compras.merge(
    df_produtos_mercados,
    on="ID Produto",
    how="left"
)


#%%
# =========================================================
# PONTO DE CHEGADA
# =========================================================
#
# Para esta problemática, o ponto de chegada é considerado igual
# ao ponto de partida.
#
# Dessa forma, uma rota completa possui a estrutura:
#
# Casa → Mercado(s) → Casa

df_escolha_mercados["Ponto de Chegada"] = (
    df_escolha_mercados["Ponto de Partida"]
)


#%%
# =========================================================
# CUSTO TOTAL DO PRODUTO
# =========================================================
#
# Converte o preço unitário em custo efetivo do item na lista:
#
# Custo Total do Produto = Preço Final × Quantidade
#
# Esse valor será utilizado nas etapas seguintes de comparação
# entre mercados e cálculo da economia potencial.
#
# Quando o preço não estiver disponível, o resultado permanece
# como NaN. Isso representa ausência de preço/disponibilidade
# daquele produto no mercado.

colunas_preco = [
    coluna
    for coluna in df_escolha_mercados.columns
    if coluna == "Preço Final"
]


for coluna in colunas_preco:

    nome_mercado = coluna.replace(
        "Preço ",
        ""
    )

    df_escolha_mercados[
        f"Custo Total Produto {nome_mercado}"
    ] = (
        df_escolha_mercados[coluna]
        *
        df_escolha_mercados["Quantidade"]
    )


#%%
# =========================================================
# ORGANIZAÇÃO DAS COLUNAS
# =========================================================
#
# Reorganiza os campos de identificação no início da base.
# A operação é apenas estrutural e não altera os valores calculados.

colunas_iniciais = [
    "ID Lista",
    "Item | Marca",
    "ID Produto",
    "Quantidade",
    "ID Ponto de Partida",
    "Ponto de Partida",
    "Ponto de Chegada"
]


colunas_restantes = [
    coluna
    for coluna in df_escolha_mercados.columns
    if coluna not in colunas_iniciais
]


df_escolha_mercados = df_escolha_mercados[
    colunas_iniciais
    +
    colunas_restantes
]


#%%
# =========================================================
# VISUALIZAÇÃO DA BASE
# =========================================================

print("\n")
print("=" * 100)
print("BASE INICIAL PARA ESCOLHA DE MERCADOS")
print("=" * 100)

print(
    df_escolha_mercados.head(10)
)


#%%
# =========================================================
# EXPORTAÇÃO DA BASE INTERMEDIÁRIA
# =========================================================
#

df_escolha_mercados.to_excel(
    output,
    index=False
)


#%%
# =========================================================
# FINAL - PARTE 1
# =========================================================

print("\n")
print("=" * 100)
print("PROCESSAMENTO DA BASE INICIAL FINALIZADO")
print("=" * 100)

print(
    f"\nQuantidade de linhas geradas: "
    f"{len(df_escolha_mercados)}"
)

print(
    f"\nArquivo salvo em:\n{output}"
)



# Validado 06/09

#%%
# =========================================================
# PARTE 2
# =========================================================


# =========================================================
# IDENTIFICAÇÃO DOS PRODUTOS NECESSÁRIOS POR LISTA
# =========================================================
#
# Define quantos itens distintos são necessários para completar
# cada lista de compras.
#
# Essa informação será utilizada para determinar se um mercado
# possui cobertura completa da lista.
#
# A combinação ID Produto + Item | Marca é utilizada para preservar
# a identificação da linha efetivamente presente na lista.

df_itens_necessarios = (
    df_compras[
        [
            "ID Lista",
            "ID Produto",
            "Item | Marca"
        ]
    ]
    .drop_duplicates()
    .assign(
        Chave_Item=lambda df:
        df["ID Produto"].astype(str)
        + "|"
        + df["Item | Marca"].astype(str)
    )
)


df_quantidade_produtos_lista = (
    df_itens_necessarios
    .groupby(
        "ID Lista"
    )["Chave_Item"]
    .nunique()
    .reset_index(
        name="Produtos Necessarios"
    )
)


#%%
# =========================================================
# CUSTO TOTAL DA LISTA POR MERCADO
# =========================================================
#
# Determina quanto custaria realizar a lista completa em cada
# mercado, sem distribuir os produtos entre estabelecimentos.
#
# ATENÇÃO:
#
# Um mercado que não possui determinado produto não pode ser
# considerado uma solução completa de compra única.
#
# Portanto, além do custo da lista, será calculada a quantidade
# de produtos efetivamente disponíveis em cada mercado.

df_escolha_mercados["_Chave_Item"] = (
    df_escolha_mercados["ID Produto"].astype(str)
    + "|"
    + df_escolha_mercados["Item | Marca"].astype(str)
)


# ---------------------------------------------------------
# CUSTO DA LISTA
# ---------------------------------------------------------

df_custo_lista = (
    df_escolha_mercados
    .groupby(
        [
            "ID Lista",
            "ID Ponto de Partida",
            "Ponto de Partida",
            "Ponto de Chegada",
            "ID Mercado",
            "Mercado"
        ],
        as_index=False
    )["Custo Total Produto Final"]
    .sum()
    .rename(
        columns={
            "Custo Total Produto Final":
            "Custo Lista"
        }
    )
)


# ---------------------------------------------------------
# QUANTIDADE DE PRODUTOS DISPONÍVEIS
# ---------------------------------------------------------
#
# Um produto é considerado disponível para o mercado quando
# possui preço válido.
#
# Portanto, valores NaN não são considerados disponibilidade.

df_disponibilidade = (
    df_escolha_mercados[
        df_escolha_mercados["ID Mercado"].notna()
        &
        df_escolha_mercados[
            "Custo Total Produto Final"
        ].notna()
    ]
    .drop_duplicates(
        [
            "ID Lista",
            "ID Mercado",
            "_Chave_Item"
        ]
    )
    .groupby(
        [
            "ID Lista",
            "ID Mercado"
        ],
        as_index=False
    )["_Chave_Item"]
    .nunique()
    .rename(
        columns={
            "_Chave_Item":
            "Produtos Disponiveis"
        }
    )
)


# ---------------------------------------------------------
# INCORPORA A INFORMAÇÃO DE COBERTURA
# ---------------------------------------------------------

df_custo_lista = df_custo_lista.merge(
    df_disponibilidade,
    on=[
        "ID Lista",
        "ID Mercado"
    ],
    how="left"
)


df_custo_lista = df_custo_lista.merge(
    df_quantidade_produtos_lista,
    on="ID Lista",
    how="left"
)


df_custo_lista["Produtos Disponiveis"] = (
    df_custo_lista["Produtos Disponiveis"]
    .fillna(0)
)


# ---------------------------------------------------------
# COBERTURA COMPLETA
# ---------------------------------------------------------
#
# True:
# o mercado possui todos os produtos necessários da lista.
#
# False:
# falta pelo menos um produto para completar a lista.

df_custo_lista["Cobertura Completa"] = (
    df_custo_lista["Produtos Disponiveis"]
    ==
    df_custo_lista["Produtos Necessarios"]
)


#%%
# =========================================================
# LEITURA DA MATRIZ DE ROTAS
# =========================================================
#
# Utiliza como entrada o arquivo gerado pelo Script 01.
#
# A matriz contém as distâncias entre:
#
# - pontos de partida e mercados;
# - mercados e mercados.


df_rotas = pd.read_excel(
    caminho_rotas
)


#%%
# =========================================================
# ROTAS:
# PONTO DE PARTIDA → MERCADO
# =========================================================
#
# Isola da matriz os deslocamentos cuja origem é um ponto
# de partida e o destino é um mercado.
#
# Esses trajetos são utilizados:
#
# - na definição do M1;
# - no primeiro trecho da rota;
# - no retorno do último mercado para o ponto de partida.

df_rotas_mercado = df_rotas[
    (df_rotas["TIPO_ORIGEM"] == "PARTIDA")
    &
    (df_rotas["TIPO_DESTINO"] == "MERCADO")
].copy()


#%%
# =========================================================
# CUSTO DO TRAJETO
#
# R$ 1,00 POR KM
# =========================================================
#
# Premissa do modelo:
# cada quilômetro percorrido é convertido diretamente em R$ 1,00
# de custo de deslocamento.
#
# Assim:
#
# Custo do Trajeto = Distância em km × R$ 1,00
#
# A justificativa dessa premissa e sua análise de sensibilidade
# são tratadas na metodologia do trabalho.

df_rotas_mercado["Custo Trajeto"] = (
    df_rotas_mercado["Custo Logistico"]
)

df_rotas_mercado["Custo Trajeto Volta"] = (
    df_rotas_mercado["Custo Logistico"]
)


#%%
# =========================================================
# VALIDAÇÃO DAS ROTAS
# =========================================================

print("\n")
print("=" * 100)
print("VALIDAÇÃO - MATRIZ DE ROTAS POR PONTO DE PARTIDA")
print("=" * 100)

print(
    df_rotas_mercado[
        [
            "ID_ORIGEM",
            "NOME_ORIGEM",
            "ID_DESTINO",
            "NOME_DESTINO",
            "Custo Logistico",
            "Custo Trajeto",
            "Custo Trajeto Volta"
        ]
    ]
    .sort_values(
        [
            "ID_ORIGEM",
            "ID_DESTINO"
        ]
    )
    .to_string(index=False)
)


#%%
# =========================================================
# CRUZAMENTO:
# CUSTO DA LISTA + CUSTO DO TRAJETO
# =========================================================
#
# O cruzamento é realizado por IDs:
#
# ID Ponto de Partida + ID Mercado
#             ↓
# ID_ORIGEM + ID_DESTINO

df_custo_lista = df_custo_lista.merge(
    df_rotas_mercado[
        [
            "ID_ORIGEM",
            "ID_DESTINO",
            "Custo Logistico",
            "Custo Trajeto",
            "Custo Trajeto Volta"
        ]
    ],
    left_on=[
        "ID Ponto de Partida",
        "ID Mercado"
    ],
    right_on=[
        "ID_ORIGEM",
        "ID_DESTINO"
    ],
    how="left"
)


#%%
# =========================================================
# VALIDAÇÃO:
# CUSTO POR LISTA + PONTO DE PARTIDA + MERCADO
# =========================================================

print("\n")
print("=" * 100)
print("VALIDAÇÃO - CUSTO POR LISTA E PONTO DE PARTIDA")
print("=" * 100)

print(
    df_custo_lista[
        [
            "ID Lista",
            "ID Ponto de Partida",
            "Ponto de Partida",
            "ID Mercado",
            "Mercado",
            "Custo Lista",
            "Produtos Disponiveis",
            "Produtos Necessarios",
            "Cobertura Completa",
            "Custo Logistico",
            "Custo Trajeto",
            "Custo Trajeto Volta"
        ]
    ]
    .sort_values(
        [
            "ID Lista",
            "ID Ponto de Partida",
            "ID Mercado"
        ]
    )
    .to_string(index=False)
)


#%%
# =========================================================
# CUSTO TOTAL DA SIMULAÇÃO
# =========================================================
#
# Nesta etapa, cada mercado é avaliado como uma solução individual.
#
# Porém, somente mercados com cobertura completa podem concorrer
# como solução de compra única.
#
# Custo Total =
# Custo da lista completa no mercado
# +
# Custo do deslocamento até esse mercado
#
# O indicador será utilizado para definir o mercado inicial M1.

df_custo_lista["Custo Total"] = (
    df_custo_lista["Custo Lista"]
    +
    df_custo_lista["Custo Trajeto"]
    +
    df_custo_lista["Custo Trajeto Volta"]
)


#%%
# =========================================================
# RANKING DOS MERCADOS
# =========================================================
#
# O ranking é realizado separadamente para cada ID Lista.
#
# IMPORTANTE:
#
# O ranking de M1 considera somente mercados com:
#
# Cobertura Completa = True
#
# Portanto, um mercado que não possui todos os produtos da lista
# é desconsiderado como solução de compra única, mesmo que o
# custo dos produtos disponíveis seja inferior.
#
# Isso evita que uma compra incompleta seja artificialmente
# considerada mais barata.

df_custo_lista["Ranking"] = pd.NA


mascara_m1_valido = (
    df_custo_lista["Cobertura Completa"]
)


df_custo_lista.loc[
    mascara_m1_valido,
    "Ranking"
] = (
    df_custo_lista.loc[
        mascara_m1_valido
    ]
    .groupby("ID Lista")["Custo Total"]
    .rank(
        method="min",
        ascending=True
    )
)


#%%
# =========================================================
# ORDENAÇÃO
# =========================================================

df_custo_lista = (
    df_custo_lista
    .sort_values(
        [
            "ID Lista",
            "Ranking"
        ],
        na_position="last"
    )
    .reset_index(drop=True)
)


#%%
# =========================================================
# MERCADO 1 DE CADA LISTA
# =========================================================
#
# M1 representa o mercado de menor custo total para realizar
# integralmente a lista, considerando:
#
# - disponibilidade de todos os produtos;
# - custo dos produtos;
# - custo do deslocamento desde o ponto de partida.
#
# M1 é a solução inicial de referência do algoritmo.
#
# A partir dele, outros mercados podem ser incorporados caso
# a economia adicional obtida nos produtos compense o custo
# adicional de deslocamento.
#
# IMPORTANTE:
#
# Um mercado que não possua todos os produtos da lista NÃO pode
# ser definido como M1.

df_M1 = (
    df_custo_lista[
        (df_custo_lista["Ranking"] == 1)
        &
        (df_custo_lista["Cobertura Completa"])
    ]
    .copy()
)


#%%
# =========================================================
# VALIDAÇÃO DO M1
# =========================================================

print("\n")
print("=" * 100)
print("MERCADO 1 - MELHOR MERCADO POR LISTA")
print("=" * 100)

print(
    df_M1[
        [
            "ID Lista",
            "ID Ponto de Partida",
            "Ponto de Partida",
            "ID Mercado",
            "Mercado",
            "Produtos Disponiveis",
            "Produtos Necessarios",
            "Cobertura Completa",
            "Custo Lista",
            "Custo Trajeto",
            "Custo Total"
        ]
    ]
    .sort_values("ID Lista")
    .to_string(index=False)
)


# Validado 06/09



#%%
# =========================================================
# LISTAS DE COMPRAS
# =========================================================
#
# Cada ID Lista representa um cenário de compra independente.
# O algoritmo executará todo o processo de seleção e otimização
# separadamente para cada lista.

listas = (
    df_escolha_mercados[
        "ID Lista"
    ]
    .drop_duplicates()
    .sort_values()
    .tolist()
)



#%%
# =========================================================
# VALIDAÇÃO:
# LISTAS SEM MERCADO COM COBERTURA COMPLETA
# =========================================================
#
# Essas listas não possuem um M1 válido segundo a regra definida
# para a solução de compra única.
#
# O algoritmo não cria um M1 artificial para essas listas.
# Elas serão interrompidas posteriormente, pois a metodologia
# utiliza M1 como ponto inicial da seleção incremental.

listas_com_m1 = set(
    df_M1["ID Lista"]
)


listas_sem_m1 = [
    id_lista
    for id_lista in listas
    if id_lista not in listas_com_m1
] if "listas" in globals() else []


# A lista de IDs é obtida diretamente da base neste ponto caso
# ainda não exista a variável "listas".
if not listas_sem_m1:

    listas_validacao = (
        df_escolha_mercados[
            "ID Lista"
        ]
        .drop_duplicates()
        .sort_values()
        .tolist()
    )

    listas_com_m1 = set(
        df_M1["ID Lista"]
    )

    listas_sem_m1 = [
        id_lista
        for id_lista in listas_validacao
        if id_lista not in listas_com_m1
    ]


if listas_sem_m1:

    print("\n")
    print("=" * 100)
    print("ATENÇÃO - LISTAS SEM MERCADO COM COBERTURA COMPLETA")
    print("=" * 100)

    print(
        "\nAs seguintes listas não possuem nenhum mercado "
        "com todos os produtos disponíveis:"
    )

    for id_lista in listas_sem_m1:

        print(
            f"  Lista {id_lista}"
        )


#%%
# =========================================================
# PARTE 3
# =========================================================


# =========================================================
# ROTAS:
# MERCADO → MERCADO
# =========================================================
#
# Isola da matriz de rotas os deslocamentos entre mercados.
#
# Esses valores são utilizados:
#
# - durante a seleção incremental;
# - posteriormente na otimização da sequência de visita.

df_rotas_mercado_mercado = df_rotas[
    (df_rotas["TIPO_ORIGEM"] == "MERCADO")
    &
    (df_rotas["TIPO_DESTINO"] == "MERCADO")
].copy()


df_rotas_mercado_mercado["Custo Trajeto"] = (
    df_rotas_mercado_mercado["Custo Logistico"]
)


#%%
# =========================================================
# FUNÇÃO:
# CUSTO MERCADO → MERCADO
# =========================================================
#
# Recebe dois IDs de mercado e consulta na matriz o custo
# do deslocamento entre eles.
#
# Retorna None quando o trecho não é encontrado.

def custo_trajeto_mercado(
    id_origem,
    id_destino
):

    linha = df_rotas_mercado_mercado[
        (df_rotas_mercado_mercado["ID_ORIGEM"] == id_origem)
        &
        (df_rotas_mercado_mercado["ID_DESTINO"] == id_destino)
    ]

    if linha.empty:

        return None

    return (
        linha["Custo Trajeto"].iloc[0]
    )


#%%
# =========================================================
# MAPA:
# ID MERCADO → NOME MERCADO
# =========================================================
#
# Cria um dicionário auxiliar para converter os IDs utilizados
# internamente pelo algoritmo em nomes legíveis nos relatórios.

mapa_nome_mercado = (
    df_escolha_mercados[
        [
            "ID Mercado",
            "Mercado"
        ]
    ]
    .drop_duplicates()
    .set_index("ID Mercado")["Mercado"]
    .to_dict()
)


#%%
# =========================================================
# FUNÇÃO:
# CUSTO PONTO DE PARTIDA → MERCADO
# =========================================================
#
# Consulta o custo do deslocamento entre um ponto de partida
# e um mercado específico.
#
# A mesma função é utilizada no retorno do último mercado
# ao ponto de partida.

def custo_trajeto_partida(
    id_partida,
    id_mercado
):

    linha = df_rotas_mercado[
        (df_rotas_mercado["ID_ORIGEM"] == id_partida)
        &
        (df_rotas_mercado["ID_DESTINO"] == id_mercado)
    ]

    if linha.empty:

        return None

    return (
        linha["Custo Trajeto"].iloc[0]
    )




#%%
# =========================================================
# ESTRUTURAS DE RESULTADO
# =========================================================

resultados_selecao = []

resultados_itens = []

resultados_resumo = []

dict_mercados_selecionados = {}

dict_ponto_partida = {}

dict_sequencia_original = {}

dict_custo_rota_original = {}



# Validado 06/09
 
 
#%%
# =========================================================
# FUNÇÃO: CUSTO COMPLETO DA ROTA
# =========================================================
#
# (Igual à função original "custo_rota", só que precisa estar
# definida ANTES do loop principal agora, pois vamos usá-la para
# avaliar cada ramo durante a própria seleção, e não só depois.)
 
def custo_rota(id_partida, sequencia_mercados):
 
    if not sequencia_mercados:
        return None
 
    custo_total = 0
 
    custo_primeiro = custo_trajeto_partida(id_partida, sequencia_mercados[0])
    if custo_primeiro is None:
        return None
    custo_total += custo_primeiro
 
    for i in range(len(sequencia_mercados) - 1):
        custo_trecho = custo_trajeto_mercado(sequencia_mercados[i], sequencia_mercados[i + 1])
        if custo_trecho is None:
            return None
        custo_total += custo_trecho
 
    custo_volta = custo_trajeto_partida(id_partida, sequencia_mercados[-1])
    if custo_volta is None:
        return None
    custo_total += custo_volta
 
    return custo_total
 
 
def melhor_rota_para_conjunto(id_partida, mercados):
    """
    Testa todas as permutações de um conjunto de mercados e retorna
    (melhor_sequencia, menor_custo). Usada tanto na seleção (para
    comparar ramos) quanto, se quiser, para substituir a etapa
    posterior de "OTIMIZAÇÃO DA ROTA" (que passaria a ficar redundante
    para a lista vencedora, já calculada aqui).
    """
 
    melhor_custo = None
    melhor_sequencia = None
 
    for sequencia in permutations(mercados):
 
        custo = custo_rota(id_partida, sequencia)
 
        if custo is None:
            continue
 
        if melhor_custo is None or custo < melhor_custo:
            melhor_custo = custo
            melhor_sequencia = sequencia
 
    return melhor_sequencia, melhor_custo
 
 
#%%
# =========================================================
# FUNÇÕES DE APOIO AO BRANCHING
# =========================================================
 
def avaliar_candidatos(mercado_atual, mercados_selecionados, melhor_preco, pivot_precos):
    """
    Calcula saldo de todos os candidatos ainda não selecionados,
    ordenados do maior saldo para o menor. Mesma lógica econômica
    do script original (economia de produto - custo de trajeto).
    """
 
    candidatos_avaliados = []
 
    candidatos = [m for m in pivot_precos.columns if m not in mercados_selecionados]
 
    for candidato in candidatos:
 
        preco_candidato = pivot_precos[candidato]
 
        diferenca = melhor_preco - preco_candidato
        mascara_disponivel = preco_candidato.notna()
        mascara_economia = mascara_disponivel & (diferenca > 0)
        desconto_total = diferenca[mascara_economia].sum()
 
        custo_trajeto = custo_trajeto_mercado(mercado_atual, candidato)
        if custo_trajeto is None:
            continue
 
        saldo = desconto_total - custo_trajeto
 
        candidatos_avaliados.append({
            "candidato": candidato,
            "desconto": desconto_total,
            "trajeto": custo_trajeto,
            "saldo": saldo
        })
 
    candidatos_avaliados.sort(key=lambda x: x["saldo"], reverse=True)
 
    return candidatos_avaliados
 
 
def copiar_estado(estado):
    return {
        "mercados_selecionados": estado["mercados_selecionados"].copy(),
        "melhor_preco": estado["melhor_preco"].copy(),
        "melhor_mercado": estado["melhor_mercado"].copy(),
        "mercado_atual": estado["mercado_atual"],
        "trechos": estado["trechos"].copy()
    }
 
 
def explorar_ramos(estado, pivot_precos, top_k, historico_selecao, id_lista):
    """
    Expande recursivamente até TOP_K candidatos por nível.
    Um ramo termina quando não há mais candidato com saldo positivo
    (mesmo critério de parada do script original).
    Retorna a lista de estados finais (folhas) gerados a partir
    do estado recebido.
    """
 
    candidatos = avaliar_candidatos(
        estado["mercado_atual"],
        estado["mercados_selecionados"],
        estado["melhor_preco"],
        pivot_precos
    )
 
    candidatos_positivos = [c for c in candidatos if c["saldo"] > 0][:top_k]
 
    # ramo se encerra: nenhum candidato viável com saldo positivo
    if not candidatos_positivos:
        return [estado]
 
    folhas = []
 
    for c in candidatos_positivos:
 
        novo_estado = copiar_estado(estado)
        candidato = c["candidato"]
 
        preco_novo_mercado = pivot_precos[candidato]
 
        mascara_melhora = (
            preco_novo_mercado.notna()
            & novo_estado["melhor_preco"].notna()
            & (preco_novo_mercado < novo_estado["melhor_preco"])
        )
 
        novo_estado["melhor_mercado"][mascara_melhora] = candidato
        novo_estado["melhor_preco"][mascara_melhora] = preco_novo_mercado[mascara_melhora]
 
        novo_estado["mercados_selecionados"].append(candidato)
        novo_estado["trechos"].append(c["trajeto"])
        novo_estado["mercado_atual"] = candidato
 
        historico_selecao.append({
            "ID Lista": id_lista,
            "Mercado Origem": estado["mercado_atual"],
            "Mercado Escolhido": candidato,
            "Desconto Total": c["desconto"],
            "Custo Trajeto": c["trajeto"],
            "Saldo": c["saldo"]
        })
 
        folhas.extend(
            explorar_ramos(novo_estado, pivot_precos, top_k, historico_selecao, id_lista)
        )
 
    return folhas
 
 
#%%
# =========================================================
# ESTRUTURAS DE RESULTADO
# (iguais às originais)
# =========================================================
 
resultados_selecao = []
resultados_itens = []
resultados_resumo = []
 
dict_mercados_selecionados = {}
dict_ponto_partida = {}
 
# aproveitamos para já guardar aqui a rota ótima calculada durante
# a própria seleção — a etapa posterior "OTIMIZAÇÃO DA ROTA" do
# script original passa a ser redundante para estas listas (pode
# manter como conferência ou remover).
dict_sequencia_otima = {}
dict_custo_otimo = {}
 
# novo: guarda quantos ramos foram gerados e quantos conjuntos únicos
# resultaram deles, por lista — evidência da amplitude da busca
dict_qtd_ramos = {}
 
 
#%%
# =========================================================
# LOOP PRINCIPAL (COM BRANCHING TOP-K)
# =========================================================
 
for id_lista in listas:
 
    df_itens_lista = df_escolha_mercados[
        df_escolha_mercados["ID Lista"] == id_lista
    ].copy()
 
    if df_itens_lista.empty:
        continue
 
    id_ponto_partida = df_itens_lista["ID Ponto de Partida"].iloc[0]
    endereco_partida = df_itens_lista["Ponto de Partida"].iloc[0]
 
    pivot_precos = (
        df_itens_lista
        .pivot_table(
            index=["ID Produto", "Item | Marca"],
            columns="ID Mercado",
            values="Custo Total Produto Final",
            aggfunc="first"
        )
    )
 
    linha_m1_lista = df_M1[df_M1["ID Lista"] == id_lista]
 
    if linha_m1_lista.empty:
        print(f"\nATENÇÃO: Lista {id_lista} não possui Mercado 1 válido.")
        continue
 
    linha_m1 = linha_m1_lista.sort_values("Custo Total").iloc[0]
    mercado_m1 = linha_m1["ID Mercado"]
 
    custo_casa_m1 = custo_trajeto_partida(id_ponto_partida, mercado_m1)
 
    if custo_casa_m1 is None:
        print(f"\nATENÇÃO: Não foi encontrada rota entre o ponto de partida "
              f"{id_ponto_partida} e o mercado {mercado_m1} para a Lista {id_lista}.")
        continue
 
    # ---------------------------------------------------
    # ESTADO INICIAL (equivalente ao ponto de partida do
    # loop original: solução contendo só M1)
    # ---------------------------------------------------
 
    estado_inicial = {
        "mercados_selecionados": [mercado_m1],
        "melhor_preco": pivot_precos[mercado_m1].copy(),
        "melhor_mercado": pd.Series(mercado_m1, index=pivot_precos[mercado_m1].index),
        "mercado_atual": mercado_m1,
        "trechos": [custo_casa_m1]
    }
 
    # ---------------------------------------------------
    # BRANCHING: gera todas as folhas (conjuntos de mercados
    # candidatos) a partir do estado inicial
    # ---------------------------------------------------
 
    folhas = explorar_ramos(estado_inicial, pivot_precos, TOP_K, resultados_selecao, id_lista)
 
    # ---------------------------------------------------
    # DEDUPLICAÇÃO: ramos diferentes podem convergir para o
    # MESMO conjunto de mercados (em ordens diferentes). Como
    # o custo final de produto e de rota só dependem do
    # CONJUNTO (não da ordem de chegada), deduplicamos por
    # frozenset para não recalcular permutations() à toa.
    # ---------------------------------------------------
 
    folhas_unicas = {}
    for folha in folhas:
        chave = frozenset(folha["mercados_selecionados"])
        if chave not in folhas_unicas:
            folhas_unicas[chave] = folha
 
    # ---------------------------------------------------
    # AVALIA CADA FOLHA: produtos + melhor rota possível
    # para aquele conjunto de mercados
    # ---------------------------------------------------
 
    melhor_folha = None
    melhor_custo_total_folha = None
    melhor_sequencia_folha = None
    melhor_custo_rota_folha = None
 
    for chave, folha in folhas_unicas.items():
 
        custo_produtos_folha = folha["melhor_preco"].sum()
 
        sequencia_otima, custo_rota_otima = melhor_rota_para_conjunto(
            id_ponto_partida,
            folha["mercados_selecionados"]
        )
 
        if sequencia_otima is None:
            continue
 
        custo_total_folha = custo_produtos_folha + custo_rota_otima
 
        if melhor_custo_total_folha is None or custo_total_folha < melhor_custo_total_folha:
            melhor_custo_total_folha = custo_total_folha
            melhor_folha = folha
            melhor_sequencia_folha = sequencia_otima
            melhor_custo_rota_folha = custo_rota_otima
 
    if melhor_folha is None:
        print(f"\nATENÇÃO: Nenhum ramo válido encontrado para a Lista {id_lista}.")
        continue
 
    # ---------------------------------------------------
    # A PARTIR DAQUI, SEGUE IGUAL AO SCRIPT ORIGINAL,
    # só que usando os dados do MELHOR RAMO em vez do
    # único caminho guloso
    # ---------------------------------------------------
 
    mercados_selecionados = melhor_folha["mercados_selecionados"]
    melhor_preco = melhor_folha["melhor_preco"]
    melhor_mercado = melhor_folha["melhor_mercado"]
    mercado_atual = melhor_folha["mercado_atual"]
 
    # já temos a rota ótima calculada — guardamos direto
    dict_sequencia_otima[id_lista] = melhor_sequencia_folha
    dict_custo_otimo[id_lista] = melhor_custo_rota_folha
 
    # ---------------------------------------------------
    # RESULTADO: ITEM -> MELHOR MERCADO
    # (idêntico ao original)
    # ---------------------------------------------------
 
    for (id_produto, item_marca) in melhor_preco.index:
 
        id_mercado_selecionado = melhor_mercado[(id_produto, item_marca)]
 
        resultados_itens.append({
            "ID Lista": id_lista,
            "ID Ponto de Partida": id_ponto_partida,
            "Ponto de Partida": endereco_partida,
            "ID Produto": id_produto,
            "Item | Marca": item_marca,
            "Mercado Selecionado": id_mercado_selecionado,
            "Nome Mercado Selecionado": mapa_nome_mercado[id_mercado_selecionado],
            "Custo Item": melhor_preco[(id_produto, item_marca)]
        })
 
    dict_mercados_selecionados[id_lista] = mercados_selecionados.copy()
    dict_ponto_partida[id_lista] = id_ponto_partida
 
    # registra a amplitude da busca para esta lista
    dict_qtd_ramos[id_lista] = {
        "Ramos Gerados": len(folhas),
        "Conjuntos Unicos Avaliados": len(folhas_unicas)
    }
 
    print(f"\nLista {id_lista} — {len(folhas_unicas)} conjunto(s) único(s) avaliado(s) "
          f"de {len(folhas)} ramo(s) gerado(s). Vencedor: "
          f"{[mapa_nome_mercado[m] for m in mercados_selecionados]} "
          f"(custo total R$ {melhor_custo_total_folha:,.2f})")
 
 

#%%
# =========================================================
# DATAFRAMES FINAIS DA SELEÇÃO
# =========================================================
#
# Converte as estruturas acumuladas durante o loop principal
# em tabelas de análise:
#
# df_selecao_mercados
#     → decisões incrementais realizadas.
#
# df_itens_por_mercado
#     → destino final de cada item.
#
# df_resumo_rotas
#     → resumo econômico/logístico por lista.

df_selecao_mercados = (
    pd.DataFrame(
        resultados_selecao
    )
)


df_itens_por_mercado = (
    pd.DataFrame(
        resultados_itens
    )
)


df_resumo_rotas = (
    pd.DataFrame(
        resultados_resumo
    )
    .fillna(0)
)


#%%
# =========================================================
# VALIDAÇÃO:
# MERCADOS ESCOLHIDOS POR LISTA
# =========================================================

print("\n")
print("=" * 100)
print("VALIDAÇÃO - MERCADOS SELECIONADOS POR LISTA")
print("=" * 100)


for id_lista in dict_mercados_selecionados:

    mercados_lista = (
        dict_mercados_selecionados[
            id_lista
        ]
    )

    id_partida = (
        dict_ponto_partida[
            id_lista
        ]
    )

    print(
        f"\nLista {id_lista}"
    )

    print(
        f"ID Ponto de Partida: "
        f"{id_partida}"
    )

    print(
        "Mercados: "
        +
        " -> ".join(
            [
                mapa_nome_mercado[m]
                for m in mercados_lista
            ]
        )
    )


#%%
# =========================================================
# OTIMIZAÇÃO DA ROTA
# =========================================================
#
# Testa todas as ordens possíveis dos mercados já selecionados
# para cada lista.
#
# IMPORTANTE:
#
# A composição dos mercados já foi definida pela etapa incremental.
# Portanto, a enumeração realizada aqui NÃO busca novas combinações
# de mercados.
#
# Ela testa somente as permutações possíveis do conjunto previamente
# selecionado.
#
# Para cada permutação é calculado o circuito completo:
#
# Casa → primeiro mercado → ... → último mercado → Casa
#
# A sequência de menor custo é armazenada como rota ótima.


# =========================================================
# FUNÇÃO:
# CUSTO COMPLETO DA ROTA
# =========================================================

def custo_rota(
    id_partida,
    sequencia_mercados
):

    if not sequencia_mercados:

        return None


    custo_total = 0


    # =====================================================
    # CASA → PRIMEIRO MERCADO
    # =====================================================

    custo_primeiro = (
        custo_trajeto_partida(
            id_partida,
            sequencia_mercados[0]
        )
    )


    if custo_primeiro is None:

        return None


    custo_total += (
        custo_primeiro
    )


    # =====================================================
    # MERCADO → MERCADO
    # =====================================================

    for i in range(
        len(sequencia_mercados) - 1
    ):

        custo_trecho = (
            custo_trajeto_mercado(
                sequencia_mercados[i],
                sequencia_mercados[i + 1]
            )
        )


        if custo_trecho is None:

            return None


        custo_total += (
            custo_trecho
        )


    # =====================================================
    # ÚLTIMO MERCADO → CASA
    # =====================================================

    custo_volta = (
        custo_trajeto_partida(
            id_partida,
            sequencia_mercados[-1]
        )
    )


    if custo_volta is None:

        return None


    custo_total += (
        custo_volta
    )


    return custo_total


#%%
# =========================================================
# ESTRUTURAS DA ROTA ÓTIMA
# =========================================================

resultados_rota_otima = []

dict_sequencia_otima = {}

dict_custo_otimo = {}


#%%
# =========================================================
# LOOP DE OTIMIZAÇÃO
# =========================================================

for id_lista in listas:

    # =====================================================
    # VERIFICA SE A LISTA FOI PROCESSADA
    # =====================================================

    if id_lista not in dict_mercados_selecionados:

        continue


    # =====================================================
    # RECUPERA OS DADOS DA LISTA ATUAL
    # =====================================================

    mercados = (
        dict_mercados_selecionados[
            id_lista
        ]
    )


    id_partida = (
        dict_ponto_partida[
            id_lista
        ]
    )


    # =====================================================
    # GARANTE QUE EXISTE PELO MENOS UM MERCADO
    # =====================================================

    if not mercados:

        continue


    # =====================================================
    # MELHOR ROTA
    # =====================================================

    melhor_custo_rota = None

    melhor_sequencia = None


    # =====================================================
    # TESTA TODAS AS PERMUTAÇÕES
    # =====================================================

    for sequencia in permutations(
        mercados
    ):

        custo = (
            custo_rota(
                id_partida,
                sequencia
            )
        )


        if custo is None:

            continue


        if (
            melhor_custo_rota is None
            or
            custo < melhor_custo_rota
        ):

            melhor_custo_rota = (
                custo
            )

            melhor_sequencia = (
                sequencia
            )


    # =====================================================
    # SE NÃO FOI POSSÍVEL CALCULAR
    # =====================================================

    if melhor_sequencia is None:

        print(
            f"\nATENÇÃO: Não foi possível "
            f"calcular uma rota válida "
            f"para a Lista {id_lista}."
        )

        continue


    # =====================================================
    # CUSTO DA ROTA ORIGINAL
    # =====================================================

    custo_rota_original = (
        custo_rota(
            id_partida,
            mercados
        )
    )


    if custo_rota_original is None:

        continue


    # =====================================================
    # GANHO DA OTIMIZAÇÃO
    # =====================================================

    ganho_otimizacao = (
        custo_rota_original
        -
        melhor_custo_rota
    )


    # =====================================================
    # RESULTADO DA ROTA
    # =====================================================

    linha_rota_otima = {

        "ID Lista":
        id_lista,

        "ID Ponto de Partida":
        id_partida,

        "Ponto de Partida":
        df_escolha_mercados[
            df_escolha_mercados[
                "ID Lista"
            ] == id_lista
        ]["Ponto de Partida"].iloc[0],

        "QTD_MERCADOS":
        len(mercados),


        "Sequencia Original":
        " -> ".join(
            [
                mapa_nome_mercado[m]
                for m in mercados
            ]
        ),

        "Custo Original":
        custo_rota_original,

        "Sequencia Otima":
        " -> ".join(
            [
                mapa_nome_mercado[m]
                for m in melhor_sequencia
            ]
        ),


        "Custo Otimo":
        melhor_custo_rota,

        "Ganho Otimizacao":
        ganho_otimizacao

    }


    # =====================================================
    # DICIONÁRIOS
    # =====================================================

    dict_sequencia_otima[
        id_lista
    ] = (
        melhor_sequencia
    )


    dict_custo_otimo[
        id_lista
    ] = (
        melhor_custo_rota
    )


    resultados_rota_otima.append(
        linha_rota_otima
    )


#%%
# =========================================================
# DATAFRAME FINAL:
# ROTA OTIMIZADA
# =========================================================

df_rota_otima = (
    pd.DataFrame(
        resultados_rota_otima
    )
)


#%%
# =========================================================
# RELATÓRIO FINAL
# RESUMO POR LISTA
# =========================================================
#
# Nesta etapa são calculados, para cada lista:
#
# - custo dos produtos da solução multimercado;
# - custo da rota otimizada;
# - custo total da solução multimercado;
# - custo da compra integral em M1;
# - diferença econômica entre as duas alternativas.
#
# A partir desses valores, a solução final recomendada será definida
# pela alternativa de menor custo total.

resultados_finais = []

for id_lista in listas:

    # =====================================================
    # VERIFICA SE A LISTA TEM RESULTADO
    # =====================================================

    if id_lista not in dict_sequencia_otima:

        continue


    # =====================================================
    # DADOS BASE DA LISTA
    # =====================================================

    linha_m1 = (
        df_M1[
            df_M1["ID Lista"] == id_lista
        ]
        .sort_values("Custo Total")
        .iloc[0]
    )


    id_m1 = (
        linha_m1[
            "ID Mercado"
        ]
    )


    nome_m1 = (
        mapa_nome_mercado[
            id_m1
        ]
    )


    custo_casa_m1 = (
        linha_m1[
            "Custo Trajeto"
        ]
    )


    # =====================================================
    # PONTO DE PARTIDA DA LISTA
    # =====================================================

    ponto_partida = (
        linha_m1[
            "Ponto de Partida"
        ]
    )


    id_ponto_partida = (
        linha_m1[
            "ID Ponto de Partida"
        ]
    )


    # =====================================================
    # CENÁRIO DE REFERÊNCIA:
    # COMPRA ÚNICA EM M1
    # =====================================================
    #
    # Como M1 possui cobertura completa, o custo abaixo representa
    # efetivamente a compra integral de todos os produtos da lista.

    custo_produtos_m1 = (
        df_escolha_mercados[
            (df_escolha_mercados[
                "ID Lista"
            ] == id_lista)
            &
            (df_escolha_mercados[
                "ID Mercado"
            ] == id_m1)
        ]["Custo Total Produto Final"]
        .sum()
    )


    # =====================================================
    # IDA + VOLTA PARA M1
    # =====================================================

    custo_total_apenas_m1 = (
        custo_produtos_m1
        +
        custo_casa_m1 * 2
    )


    # =====================================================
    # TABELA DOS ITENS DA SOLUÇÃO MULTIMERCADO
    # =====================================================

    df_tabela = (
        df_itens_por_mercado[
            df_itens_por_mercado[
                "ID Lista"
            ] == id_lista
        ]
        .copy()
    )


    # =====================================================
    # MERGE COM QUANTIDADE E PREÇO UNITÁRIO
    # =====================================================

    df_tabela = df_tabela.merge(
        df_escolha_mercados[
            [
                "ID Lista",
                "ID Produto",
                "ID Mercado",
                "Quantidade",
                "Preço Final"
            ]
        ],
        left_on=[
            "ID Lista",
            "ID Produto",
            "Mercado Selecionado"
        ],
        right_on=[
            "ID Lista",
            "ID Produto",
            "ID Mercado"
        ],
        how="left"
    )


 
    # =====================================================
  # CUSTOS DA SOLUÇÃO MULTIMERCADO
  # =====================================================

    custo_lista_total = (
              df_tabela[
                  "Custo Item"
                  ]
              .sum()
              )


    custo_deslocamento_otimo = (
      dict_custo_otimo[
          id_lista
      ]
  )


    custo_total_otimizado = (
      custo_lista_total
      +
      custo_deslocamento_otimo
  )


  # =====================================================
  # VALIDAÇÃO FINAL DA SOLUÇÃO
  # =====================================================

    if round(custo_total_otimizado, 2) < round(custo_total_apenas_m1, 2):

        custo_total_final = custo_total_otimizado
        estrategia_final = "MULTIMERCADO"

    else:

        custo_total_final = custo_total_apenas_m1
        estrategia_final = "MERCADO ÚNICO"
  # =====================================================
  # SINCRONIZA COM A ESTRATÉGIA FINAL (AJUSTE 1 + AJUSTE 2)
  # =====================================================

    if estrategia_final == "MERCADO ÚNICO":

      custo_lista_total = custo_produtos_m1

      custo_deslocamento_otimo = custo_casa_m1 * 2

      # custo_total_otimizado = custo_total_apenas_m1

      df_tabela = (
          df_escolha_mercados[
              (df_escolha_mercados["ID Lista"] == id_lista)
              &
              (df_escolha_mercados["ID Mercado"] == id_m1)
          ]
          [
              [
                  "ID Lista",
                  "ID Produto",
                  "Item | Marca",
                  "Quantidade",
                  "Preço Final",
                  "Custo Total Produto Final"
              ]
          ]
          .rename(
              columns={
                  "Custo Total Produto Final": "Custo Item"
              }
          )
          .copy()
      )

      df_tabela["Mercado Selecionado"] = id_m1

      df_tabela["Nome Mercado Selecionado"] = nome_m1  
    
    # =====================================================
    # QUANTIDADE DE PRODUTOS DA LISTA
    # =====================================================
    
    qtd_produtos = (
    df_tabela["ID Produto"]
    .nunique()
    )
    
    
    # =====================================================
    # ECONOMIA FINAL
    # =====================================================

    economia_total = (
    custo_total_apenas_m1
    -
    custo_total_otimizado)


    # =====================================================
    # ECONOMIA %
    # =====================================================

    if custo_total_apenas_m1 != 0:

        economia_percentual = (
            economia_total
            /
            custo_total_apenas_m1
        )

    else:

        economia_percentual = 0


    # =====================================================
    # ROTA ÓTIMA
    # =====================================================

    sequencia_otima = (
        dict_sequencia_otima[
            id_lista
        ]
    )


    if estrategia_final == "MULTIMERCADO":
        qtd_mercados = (
            len(
                dict_mercados_selecionados[
                    id_lista
                ]
            )
        )
    else:
        qtd_mercados = 1
        
        
        
 # Append do resultado final


    resultados_finais.append({
    
        "ID Lista":
        id_lista,
    
        "ID Ponto de Partida":
        id_ponto_partida,
    
        "Ponto de Partida":
        ponto_partida,
    
        "Mercado Inicial":
        nome_m1,
    
        "Qtd Mercados":
        qtd_mercados,
        
        "Qtd Produtos":
        qtd_produtos,
    
        "Custo Produtos":
        custo_lista_total,
    
        "Custo Deslocamento":
        custo_deslocamento_otimo,
    
        "Custo Total Multimercado":
        custo_total_otimizado,
    
        "Custo Compra Única":
        custo_total_apenas_m1,
    
        "Custo Final Recomendado":
        custo_total_final,
    
        "Estratégia Recomendada":
        estrategia_final,
    
        "Economia Total":
        economia_total,
    
        "Economia %":
        economia_percentual
    
    })      
        

    # =====================================================
    # IMPRESSÃO
    # =====================================================

    print("\n")
    print("=" * 100)

    print(
        f"LISTA {id_lista}"
    )

    print("=" * 100)

    print(
        f"Ponto de Partida:       "
        f"{ponto_partida}"
    )

    print(
        f"ID Ponto de Partida:    "
        f"{id_ponto_partida}"
    )

    print(
        f"Quantidade de Mercados: "
        f"{qtd_mercados}"
    )

    print(
        f"Custo Total Produtos:   "
        f"R$ {custo_lista_total:,.2f}"
    )

    print(
        f"Custo Deslocamento:     "
        f"R$ {custo_deslocamento_otimo:,.2f}"
    )

    print(
        f"Custo Multimercado:     "
        f"R$ {custo_total_otimizado:,.2f}"
    )

    print(
        f"Custo Compra Única:     "
        f"R$ {custo_total_apenas_m1:,.2f}"
    )

    print(
        f"Custo Final:            "
        f"R$ {custo_total_final:,.2f}"
    )

    print(
        f"Estratégia Recomendada: "
        f"{estrategia_final}"
    )


    # =====================================================
    # DESCRIÇÃO DA LISTA
    # =====================================================

    print("\nDescrição da Lista:")

    print("-" * 100)

    print(
        f"{'Item | Marca':<40}"
        f"{'Qtd':>6}"
        f"{'Preço Unit.':>16}"
        f"{'Preço Total':>16}"
        f"{'Mercado':>20}"
    )

    print("-" * 100)


    # =====================================================
    # ITENS
    # =====================================================

    for _, item in df_tabela.iterrows():

        print(
            f"{str(item['Item | Marca'])[:39]:<40}"
            f"{item['Quantidade']:>6}"
            f"{'R$ ' + format(item['Preço Final'], ',.2f'):>16}"
            f"{'R$ ' + format(item['Custo Item'], ',.2f'):>16}"
            f"{str(item['Nome Mercado Selecionado'])[:19]:>20}"
        )


    print("-" * 100)


    # =====================================================
    # ROTA ÓTIMA
    # =====================================================

    print("\nRota Ótima:")

    print(
        f"  {ponto_partida}"
    )


    for id_mercado in sequencia_otima:

        print(
            f"    -> "
            f"{mapa_nome_mercado[id_mercado]}"
        )


    print(
        f"    -> "
        f"{ponto_partida}"
    )


    # =====================================================
    # ECONOMIA
    # =====================================================

    print("\nEconomia:")

    print(
        f"Custo de Compra Único em {nome_m1}: "
        f"R$ {custo_total_apenas_m1:,.2f}"
    )

    print(
        f"Custo Multimercado:                  "
        f"R$ {custo_total_otimizado:,.2f}"
    )

    print(
        f"Custo Final Recomendado:             "
        f"R$ {custo_total_final:,.2f}"
    )

    print(
        f"Estratégia Recomendada:               "
        f"{estrategia_final}"
    )

    print(
        f"Economia Total:                       "
        f"R$ {economia_total:,.2f}"
    )

    print(
        f"Economia %:                           "
        f"{economia_percentual:.2%}"
    )


    print("\n")


#%%
# =========================================================
# FINAL
# =========================================================

print("\n")
print("=" * 100)
print("PROCESSAMENTO DO SCRIPT 03 FINALIZADO")
print("=" * 100)

print(
    f"\nQuantidade de listas processadas: "
    f"{len(dict_sequencia_otima)}"
)

print(
    f"\nListas processadas: "
    f"{list(dict_sequencia_otima.keys())}"
)

print("\n")



# Validado 06/09 - Não mexer acima daqui



# =========================================================
# DATAFRAME FINAL DOS RESULTADOS
# =========================================================

df_resultado_final = pd.DataFrame(
    resultados_finais
)



#%%
# =========================================================
# EXPORTAÇÃO DOS RESULTADOS OTIMIZADOS PARA EXCEL
# =========================================================
#
# Esta etapa possui exclusivamente finalidade de apresentação.
#
# Os cálculos econômicos e logísticos já foram realizados
# anteriormente pelo algoritmo e consolidados em:
#
# df_resultado_final
# df_itens_por_mercado
# df_rota_otima
#
# Portanto, esta etapa não recalcula:
#
# - custo dos produtos;
# - custo de deslocamento;
# - custo total multimercado;
# - custo da compra única;
# - estratégia recomendada;
# - economia;
# - rota ótima.
#
# A função deste bloco é apenas organizar e formatar
# os resultados já calculados para o arquivo Excel.

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side




# =========================================================
# CRIAÇÃO DO WORKBOOK
# =========================================================

wb = Workbook()

ws_inicial = wb.active

wb.remove(
    ws_inicial
)


# =========================================================
# ESTILOS
# =========================================================

fonte_titulo = Font(
    bold=True,
    size=14
)

fonte_subtitulo = Font(
    bold=True,
    size=11
)

fonte_cabecalho = Font(
    bold=True
)

borda = Border(
    bottom=Side(
        style="thin"
    )
)

alinhamento_centro = Alignment(
    horizontal="center",
    vertical="center"
)

alinhamento_esquerda = Alignment(
    horizontal="left",
    vertical="center"
)


# =========================================================
# UMA ABA PARA CADA ID LISTA
# =========================================================

for _, resultado in df_resultado_final.iterrows():

    id_lista = resultado["ID Lista"]


    # =====================================================
    # DADOS JÁ CALCULADOS
    # =====================================================

    ponto_partida = (
        resultado["Ponto de Partida"]
    )

    nome_m1 = (
        resultado["Mercado Inicial"]
    )

    qtd_mercados = (
        resultado["Qtd Mercados"]
    )
    
    qtd_produtos = (
        resultado["Qtd Produtos"]
    )

    custo_lista_total = (
        resultado["Custo Produtos"]
    )

    custo_deslocamento_otimo = (
        resultado["Custo Deslocamento"]
    )

    custo_total_otimizado = (
        resultado["Custo Total Multimercado"]
    )

    custo_total_apenas_m1 = (
        resultado["Custo Compra Única"]
    )

    custo_total_final = (
        resultado["Custo Final Recomendado"]
    )

    estrategia_final = (
        resultado["Estratégia Recomendada"]
    )

    economia_total = (
        resultado["Economia Total"]
    )

    economia_percentual = (
        resultado["Economia %"]
    )


    # =====================================================
    # ROTA JÁ CALCULADA
    # =====================================================

    sequencia_otima = (
        dict_sequencia_otima[
            id_lista
        ]
    )


    # =====================================================
    # TABELA DE ITENS
    # =====================================================
    #
    # Utiliza diretamente o resultado da seleção dos produtos.
    #
    # Apenas recupera quantidade e preço unitário para apresentação.


    df_tabela = (
        df_itens_por_mercado[
            df_itens_por_mercado["ID Lista"] == id_lista
        ]
        .copy()
    )

    df_tabela = df_tabela.merge(
        df_escolha_mercados[
            [
                "ID Lista",
                "ID Produto",
                "ID Mercado",
                "Quantidade",
                "Preço Final"
            ]
        ],
        left_on=[
            "ID Lista",
            "ID Produto",
            "Mercado Selecionado"
        ],
        right_on=[
            "ID Lista",
            "ID Produto",
            "ID Mercado"
        ],
        how="left"
    )


    # =====================================================
    # TABELA FINAL COERENTE COM A ESTRATÉGIA (AJUSTE 3)
    # =====================================================

    if resultado["Estratégia Recomendada"] == "MERCADO ÚNICO":

        id_m1_lista = (
            df_M1[
                df_M1["ID Lista"] == id_lista
            ]
            .sort_values("Custo Total")
            ["ID Mercado"]
            .iloc[0]
        )

        df_tabela = (
            df_escolha_mercados[
                (df_escolha_mercados["ID Lista"] == id_lista)
                &
                (df_escolha_mercados["ID Mercado"] == id_m1_lista)
            ]
            [
                [
                    "ID Lista",
                    "ID Produto",
                    "Item | Marca",
                    "Quantidade",
                    "Preço Final",
                    "Custo Total Produto Final"
                ]
            ]
            .rename(
                columns={
                    "Custo Total Produto Final": "Custo Item"
                }
            )
            .copy()
        )

        df_tabela["Mercado Selecionado"] = id_m1_lista

        df_tabela["Nome Mercado Selecionado"] = nome_m1


    # =====================================================
    # CRIA ABA
    # =====================================================

    nome_aba = (
        f"Lista {id_lista}"
    )

    nome_aba = (
        str(nome_aba)[:31]
    )

    ws = wb.create_sheet(
        title=nome_aba
    )


    # =====================================================
    # CABEÇALHO
    # =====================================================

    ws["A1"] = (
        f"RESULTADO ÓTIMO - LISTA {id_lista}"
    )

    ws["A1"].font = (
        fonte_titulo
    )

    ws.merge_cells(
        "A1:F1"
    )


    # =====================================================
    # RESUMO
    # =====================================================

    ws["A3"] = (
        "RESUMO DA LISTA"
    )

    ws["A3"].font = (
        fonte_subtitulo
    )


    resumo = [

        (
            "ID Lista",
            id_lista
        ),

        (
            "Ponto de Partida",
            ponto_partida
        ),

        (
            "Quantidade de Mercados",
            qtd_mercados
        ),
        
        (
            "Quantidade de Produtos",
            resultado["Qtd Produtos"]
        ),

        (
            "Primeiro Mercado Avaliado",
            nome_m1
        ),

        (
            "Custo Produtos Multimercado",
            custo_lista_total
        ),

        (
            "Custo Deslocamento Multimercado",
            custo_deslocamento_otimo
        ),

        (
            "Custo Total Multimercado",
            custo_total_otimizado
        ),

        (
            "Custo Compra Única em M1",
            custo_total_apenas_m1
        ),

        (
            "Custo Final Recomendado",
            custo_total_final
        ),

        (
            "Estratégia Recomendada",
            estrategia_final
        ),

        (
            "Economia Total",
            economia_total
        ),

        (
            "Economia %",
            economia_percentual
        )
    ]


    linha = 4


    for descricao, valor in resumo:

        ws.cell(
            row=linha,
            column=1,
            value=descricao
        )

        ws.cell(
            row=linha,
            column=1
        ).font = (
            fonte_cabecalho
        )

        ws.cell(
            row=linha,
            column=2,
            value=valor
        )

        linha += 1


    # =====================================================
    # FORMATAÇÃO DOS VALORES MONETÁRIOS
    # =====================================================

    for linha_excel in [
        9,
        10,
        11,
        12,
        13,
        15
    ]:

        ws.cell(
            row=linha_excel,
            column=2
        ).number_format = (
            'R$ #,##0.00'
        )


    # =====================================================
    # FORMATAÇÃO DA ECONOMIA %
    # =====================================================

    ws.cell(
        row=linha - 1,
        column=2
    ).number_format = (
        '0.00%'
    )


    # =====================================================
    # ROTA ÓTIMA
    # =====================================================

    linha_rota = 19

    ws.cell(
        row=linha_rota,
        column=1,
        value="ROTA ÓTIMA"
    )

    ws.cell(
        row=linha_rota,
        column=1
    ).font = (
        fonte_subtitulo
    )

    linha_rota += 1

    ws.cell(
        row=linha_rota,
        column=1,
        value="Casa"
    )

    linha_rota += 1


    for id_mercado in sequencia_otima:

        ws.cell(
            row=linha_rota,
            column=1,
            value=(
                f"→ "
                f"{mapa_nome_mercado[id_mercado]}"
            )
        )

        linha_rota += 1


    ws.cell(
        row=linha_rota,
        column=1,
        value="→ Casa"
    )


    # =====================================================
    # TABELA DE ITENS
    # =====================================================

    linha_tabela = (
        linha_rota
        +
        3
    )

    ws.cell(
        row=linha_tabela,
        column=1,
        value="DETALHAMENTO DA COMPRA"
    )

    ws.cell(
        row=linha_tabela,
        column=1
    ).font = (
        fonte_subtitulo
    )

    linha_tabela += 1


    # =====================================================
    # CABEÇALHOS
    # =====================================================

    cabecalhos = [

        "ID Produto",
        "Item | Marca",
        "Quantidade",
        "Preço Unitário",
        "Custo Total",
        "Mercado"

    ]


    for coluna, cabecalho in enumerate(
        cabecalhos,
        start=1
    ):

        celula = ws.cell(
            row=linha_tabela,
            column=coluna,
            value=cabecalho
        )

        celula.font = (
            fonte_cabecalho
        )

        celula.alignment = (
            alinhamento_centro
        )

        celula.border = (
            borda
        )


    linha_tabela += 1


    # =====================================================
    # ITENS
    # =====================================================

    for _, item in df_tabela.iterrows():

        ws.cell(
            row=linha_tabela,
            column=1,
            value=item["ID Produto"]
        )

        ws.cell(
            row=linha_tabela,
            column=2,
            value=item["Item | Marca"]
        )

        ws.cell(
            row=linha_tabela,
            column=3,
            value=item["Quantidade"]
        )

        ws.cell(
            row=linha_tabela,
            column=4,
            value=item["Preço Final"]
        )

        ws.cell(
            row=linha_tabela,
            column=5,
            value=item["Custo Item"]
        )

        ws.cell(
            row=linha_tabela,
            column=6,
            value=item["Nome Mercado Selecionado"]
        )


        ws.cell(
            row=linha_tabela,
            column=4
        ).number_format = (
            'R$ #,##0.00'
        )

        ws.cell(
            row=linha_tabela,
            column=5
        ).number_format = (
            'R$ #,##0.00'
        )


        linha_tabela += 1


    # =====================================================
    # RESUMO FINANCEIRO FINAL
    # =====================================================

    linha_final = (
        linha_tabela
        +
        2
    )

    ws.cell(
        row=linha_final,
        column=1,
        value="RESUMO FINANCEIRO"
    )

    ws.cell(
        row=linha_final,
        column=1
    ).font = (
        fonte_subtitulo
    )

    linha_final += 1


    financeiro = [

        (
            "Compra Única em M1",
            custo_total_apenas_m1
        ),

        (
            "Custo Multimercado",
            custo_total_otimizado
        ),

        (
            "Custo Final Recomendado",
            custo_total_final
        ),

        (
            "Estratégia Recomendada",
            estrategia_final
        ),

        (
            "Economia Total",
            economia_total
        ),

        (
            "Economia %",
            economia_percentual
        )
    ]


    for descricao, valor in financeiro:

        ws.cell(
            row=linha_final,
            column=1,
            value=descricao
        )

        ws.cell(
            row=linha_final,
            column=2,
            value=valor
        )


        if descricao in [
            "Compra Única em M1",
            "Custo Multimercado",
            "Custo Final Recomendado",
            "Economia Total"
        ]:

            ws.cell(
                row=linha_final,
                column=2
            ).number_format = (
                'R$ #,##0.00'
            )


        if descricao == "Economia %":

            ws.cell(
                row=linha_final,
                column=2
            ).number_format = (
                '0.00%'
            )


        linha_final += 1


    # =====================================================
    # AJUSTE DE LARGURA
    # =====================================================

    larguras = {

        "A": 30,
        "B": 40,
        "C": 14,
        "D": 18,
        "E": 18,
        "F": 30

    }


    for coluna, largura in larguras.items():

        ws.column_dimensions[
            coluna
        ].width = largura


    # =====================================================
    # ALINHAMENTO
    # =====================================================

    for row in ws.iter_rows():

        for cell in row:

            if cell.value is not None:

                if cell.column in [
                    1,
                    2,
                    6
                ]:

                    cell.alignment = (
                        alinhamento_esquerda
                    )

                else:

                    cell.alignment = (
                        alinhamento_centro
                    )


#%%
# =========================================================
# ABA CENTRAL - RESUMO DE TODAS AS LISTAS
# =========================================================

ws_resumo = wb.create_sheet(
    title="Resumo Geral",
    index=0
)


# =========================================================
# CABEÇALHO
# =========================================================

cabecalhos_resumo = [

    "ID Lista",
    "ID Ponto de Partida",
    "Ponto de Partida",
    "Qtd Mercados",
    "Qtd Produtos",
    "Mercado Inicial",
    "Custo Produtos",
    "Custo Deslocamento",
    "Custo Total Multimercado",
    "Custo Compra Única",
    "Custo Final Recomendado",
    "Estratégia Recomendada",
    "Economia Total",
    "Economia %",
    "Sequência Original",
    "Sequência Ótima"

]


for coluna, cabecalho in enumerate(
    cabecalhos_resumo,
    start=1
):

    celula = ws_resumo.cell(
        row=1,
        column=coluna,
        value=cabecalho
    )

    celula.font = (
        fonte_cabecalho
    )

    celula.alignment = (
        alinhamento_centro
    )

    celula.border = (
        borda
    )


# =========================================================
# APPEND DOS RESULTADOS JÁ CALCULADOS
# =========================================================

linha_resumo = 2


for _, resultado in df_resultado_final.iterrows():

    id_lista = (
        resultado["ID Lista"]
    )


    # =====================================================
    # RECUPERA ROTA JÁ CALCULADA
    # =====================================================

    mercados_selecionados = (
        dict_mercados_selecionados[
            id_lista
        ]
    )

    sequencia_otima = (
        dict_sequencia_otima[
            id_lista
        ]
    )


    if resultado["Estratégia Recomendada"] == "MULTIMERCADO":
    
        sequencia_original = (
            " -> ".join(
                [
                    mapa_nome_mercado[m]
                    for m in mercados_selecionados
                ]
            )
        )
    
        sequencia_otima_texto = (
            " -> ".join(
                [
                    mapa_nome_mercado[m]
                    for m in sequencia_otima
                ]
            )
        )
    
    else:
    
        sequencia_original = (
            mapa_nome_mercado[
                dict_mercados_selecionados[
                    id_lista
                ][0]
            ]
        )
    
        sequencia_otima_texto = (
            mapa_nome_mercado[
                dict_mercados_selecionados[
                    id_lista
                ][0]
            ]
        )

    # =====================================================
    # VALORES JÁ CALCULADOS
    # =====================================================

    valores = [

        resultado["ID Lista"],

        resultado["ID Ponto de Partida"],

        resultado["Ponto de Partida"],

        resultado["Qtd Mercados"],
        
        resultado["Qtd Produtos"],

        resultado["Mercado Inicial"],

        resultado["Custo Produtos"],

        resultado["Custo Deslocamento"],

        resultado["Custo Total Multimercado"],

        resultado["Custo Compra Única"],

        resultado["Custo Final Recomendado"],

        resultado["Estratégia Recomendada"],

        resultado["Economia Total"],

        resultado["Economia %"],

        sequencia_original,

        sequencia_otima_texto

    ]


    for coluna, valor in enumerate(
        valores,
        start=1
    ):

        ws_resumo.cell(
            row=linha_resumo,
            column=coluna,
            value=valor
        )


    # =====================================================
    # FORMATAÇÃO MONETÁRIA
    # =====================================================

    for coluna in [
        6,
        7,
        8,
        9,
        10,
        11,
        12,
        13
    ]:

        ws_resumo.cell(
            row=linha_resumo,
            column=coluna
        ).number_format = (
            'R$ #,##0.00'
        )


    # =====================================================
    # FORMATAÇÃO %
    # =====================================================

    ws_resumo.cell(
        row=linha_resumo,
        column=14
    ).number_format = (
        '0.00%'
    )


    linha_resumo += 1


#%%
# =========================================================
# FILTRO
# =========================================================

ws_resumo.auto_filter.ref = (
    f"A1:P{linha_resumo - 1}"
)


# =========================================================
# LARGURA DAS COLUNAS
# =========================================================

larguras_resumo = {

    "A": 12,
    "B": 20,
    "C": 45,
    "D": 15,
    "E": 25,
    "F": 18,
    "G": 20,
    "H": 24,
    "I": 20,
    "J": 24,
    "K": 24,
    "L": 18,
    "M": 14,
    "N": 45,
    "O": 45,
    "P": 45

}


for coluna, largura in larguras_resumo.items():

    ws_resumo.column_dimensions[
        coluna
    ].width = largura


# =========================================================
# ALINHAMENTO
# =========================================================

for row in ws_resumo.iter_rows():

    for cell in row:

        if cell.value is not None:

            if cell.column in [
                3,
                5,
                11,
                14,
                15
            ]:

                cell.alignment = (
                    alinhamento_esquerda
                )

            else:

                cell.alignment = (
                    alinhamento_centro
                )


#%%
# =========================================================
# SALVAMENTO
# =========================================================

wb.save(
    output_resultado
)


print(
    f"\nArquivo de resultados salvo em:\n"
    f"{output_resultado}"
)