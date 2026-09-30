# =========================================================
# EXPLICATIVO RESUMIDO
# =========================================================
#
# Este script tem como finalidade gerar, a partir da base de
# produtos e preços por mercado, um conjunto diversificado de
# listas de compras (cenários), a serem posteriormente
# utilizadas como entrada para o modelo de otimização
# (Scripts 01 a 04).
#
# A base de origem (Lista_Produtos_e_Mercados.xlsx) possui a
# seguinte estrutura:
#
# - cada produto pertence a uma "Categoria Produto"
#   (ex.: Arroz, Açúcar, Carne...);
#
# - dentro de cada categoria, existem 3 variações de produto
#   (marca, tipo e/ou peso), identificadas pela coluna
#   "Chave de Escolha";
#
# - cada variação possui 7 preços, um por mercado avaliado.
#
# Como o mercado, isoladamente, não é relevante para a
# definição das listas (ele serve apenas como fonte de
# variação de preço), este script ignora a informação de
# mercado no resultado final e retorna somente:
#
# - ID Lista de Compras
# - Chave de Escolha
#
# São geradas 11 listas, agrupadas em três blocos de regras:
#
# BLOCO 1 - "1 item de cada categoria", segundo 3 critérios:
#   Lista 1: sempre o item mais barato da categoria;
#   Lista 2: sempre o item mais caro da categoria;
#   Lista 3: o item mais próximo da mediana de preços
#            da categoria (não é o preço médio calculado,
#            e sim o item real mais próximo dele).
#
# BLOCO 2 - Cestas de tamanhos diferentes, com categorias e
# itens sorteados aleatoriamente:
#   Lista 4: cesta pequena (1 a 4 itens);
#   Lista 5: cesta média (5 a 10 itens);
#   Lista 6: cesta grande (mais de 10 itens).
#
# BLOCO 3 - 5 listas "variadas", combinando os critérios
# acima de formas não padronizadas, para ampliar a
# diversidade de cenários testados pelo modelo:
#   Lista 7:  mix de itens baratos/caros/médios;
#   Lista 8:  foco temático (proteínas e carboidratos);
#   Lista 9:  cesta compacta, somente itens mais caros;
#   Lista 10: cesta ampla, somente itens mais baratos;
#   Lista 11: cesta totalmente aleatória.
#
# IMPORTANTE - REPRODUTIBILIDADE:
#
# O sorteio de categorias e itens utiliza uma semente fixa
# (random.seed) para que a execução do script produza sempre
# o mesmo resultado. Isso é importante para a rastreabilidade
# acadêmica: qualquer pessoa que rodar este script novamente,
# nas mesmas condições, deve obter exatamente as mesmas 11
# listas aqui documentadas.
#
# =========================================================


#%%
# =========================================================
# IMPORTS
# =========================================================

import random

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


#%%
# =========================================================
# SEMENTE DE ALEATORIEDADE
# =========================================================
#
# Fixar a semente garante que os sorteios de categorias e
# itens (usados nos Blocos 2 e 3) sejam sempre os mesmos a
# cada execução do script.

SEMENTE_ALEATORIA = 42

random.seed(SEMENTE_ALEATORIA)


#%%
# =========================================================
# CAMINHOS
# =========================================================

from config import BASE_FONTE, CENARIOS_LISTAS
caminho_entrada = BASE_FONTE
caminho_saida = CENARIOS_LISTAS

#%%
# =========================================================
# LEITURA DA BASE
# =========================================================
#
# A base contém, para cada combinação de produto x mercado,
# o preço final praticado. É a partir dela que os cenários
# de compra serão montados.

df = pd.read_excel(
    caminho_entrada,
    sheet_name="Produtos & Mercados"
)


# ---------------------------------------------------------
# NORMALIZAÇÃO DE TEXTO (CORREÇÃO DEFENSIVA)
# ---------------------------------------------------------
#
# Remove espaços extras no início/fim do nome da categoria.
# Diferenças desse tipo (ex.: "Carne " com espaço sobrando)
# fazem com que um filtro por nome de categoria retorne uma
# base vazia, sem nenhum aviso claro do porquê.

df["Categoria Produto"] = (
    df["Categoria Produto"]
    .astype(str)
    .str.strip()
)


#%%
# =========================================================
# LISTA DE CATEGORIAS
# =========================================================
#
# Cada categoria representa um "tipo" de produto da cesta
# básica (ex.: Arroz, Açúcar, Carne...). É neste nível que a
# variedade das listas de compras será definida - dentro de
# cada categoria, o script escolhe qual variação/preço será
# utilizado.

cats = (
    sorted(
        df["Categoria Produto"]
        .unique()
        .tolist()
    )
)


#%%
# =========================================================
# FUNÇÕES AUXILIARES DE SELEÇÃO DE ITEM
# =========================================================
#
# As três funções abaixo recebem um subconjunto da base
# (já filtrado para uma única categoria) e retornam a
# "Chave de Escolha" do item selecionado, de acordo com o
# critério correspondente.
#
# IMPORTANTE:
#
# Em todos os casos, o item retornado é um item que
# efetivamente existe na base (com seu preço real em algum
# mercado) - nenhum valor é inventado ou interpolado.

def validar_sub_categoria(sub_categoria, categoria):
    """
    Verifica se o filtro por categoria retornou pelo menos
    uma linha. Se estiver vazio, interrompe a execução com
    uma mensagem explícita indicando qual nome de categoria
    não foi encontrado na base - em vez de deixar o erro
    aparecer, sem contexto, dentro de uma função do pandas
    (ex.: .sample(), .idxmin()).

    Causa mais comum: o nome da categoria usado no script
    (hardcoded, como na Lista 8) não é idêntico, caractere
    por caractere, ao valor presente na coluna
    "Categoria Produto" da planilha de origem.
    """

    if sub_categoria.empty:

        categorias_disponiveis = (
            sorted(
                df["Categoria Produto"]
                .unique()
                .tolist()
            )
        )

        raise ValueError(
            f"\nCategoria não encontrada na base: "
            f"{categoria!r}\n"
            f"Categorias disponíveis na planilha: "
            f"{categorias_disponiveis}\n"
            f"Verifique acentuação, espaços e maiúsculas/"
            f"minúsculas."
        )


def item_mais_barato(sub_categoria, categoria=None):
    """
    Retorna a Chave de Escolha do item de menor preço
    dentro da categoria informada, considerando todos os
    mercados disponíveis.
    """

    validar_sub_categoria(sub_categoria, categoria)

    indice_menor_preco = (
        sub_categoria["Preço Final"].idxmin()
    )

    return (
        sub_categoria.loc[
            indice_menor_preco,
            "Chave de Escolha"
        ]
    )


def item_mais_caro(sub_categoria, categoria=None):
    """
    Retorna a Chave de Escolha do item de maior preço
    dentro da categoria informada, considerando todos os
    mercados disponíveis.
    """

    validar_sub_categoria(sub_categoria, categoria)

    indice_maior_preco = (
        sub_categoria["Preço Final"].idxmax()
    )

    return (
        sub_categoria.loc[
            indice_maior_preco,
            "Chave de Escolha"
        ]
    )


def item_mediano(sub_categoria, categoria=None):
    """
    Retorna a Chave de Escolha do item cujo preço está mais
    próximo da mediana de preços da categoria.

    Não se calcula um "preço médio" artificial: busca-se,
    entre os itens realmente existentes, aquele cujo preço
    mais se aproxima do valor central da distribuição.
    """

    validar_sub_categoria(sub_categoria, categoria)

    mediana_preco = (
        sub_categoria["Preço Final"].median()
    )

    sub_categoria = sub_categoria.copy()

    sub_categoria["Distancia_Mediana"] = (
        (sub_categoria["Preço Final"] - mediana_preco)
        .abs()
    )

    indice_mais_proximo = (
        sub_categoria["Distancia_Mediana"].idxmin()
    )

    return (
        sub_categoria.loc[
            indice_mais_proximo,
            "Chave de Escolha"
        ]
    )


def item_aleatorio(categoria, criterio=None):
    """
    Retorna a Chave de Escolha de um item da categoria
    informada.

    Parâmetro 'criterio':
        None      -> sorteia um item qualquer da categoria;
        "barato"  -> força o item mais barato;
        "caro"    -> força o item mais caro;
        "mediano" -> força o item mais próximo da mediana.

    Esta função permite reaproveitar a mesma interface tanto
    para seleção puramente aleatória (Blocos 2 e 3) quanto
    para forçar um critério específico dentro de uma lista
    "mista" (ex.: Lista 7, que alterna critérios).
    """

    sub_categoria = (
        df[df["Categoria Produto"] == categoria]
    )

    validar_sub_categoria(sub_categoria, categoria)

    if criterio == "barato":

        return item_mais_barato(sub_categoria, categoria)

    elif criterio == "caro":

        return item_mais_caro(sub_categoria, categoria)

    elif criterio == "mediano":

        return item_mediano(sub_categoria, categoria)

    else:

        return (
            sub_categoria["Chave de Escolha"]
            .sample(n=1)
            .iloc[0]
        )


#%%
# =========================================================
# ESTRUTURAS DE RESULTADO
# =========================================================
#
# 'linhas_resultado' acumula, para cada lista gerada, os
# pares (ID Lista de Compras, Chave de Escolha) - este é o
# formato final solicitado.
#
# 'linhas_legenda' acumula, para cada ID de lista, uma
# descrição textual do critério utilizado para sua
# montagem, servindo como documentação de apoio (aba
# "Legenda" no arquivo de saída) - não é obrigatório para o
# uso do modelo, mas facilita a rastreabilidade do que foi
# feito.

linhas_resultado = []

linhas_legenda = []


def registrar_lista(
    id_lista,
    nome_lista,
    criterio_texto,
    chaves_selecionadas
):
    """
    Adiciona uma lista de compras já montada às estruturas
    de resultado e de legenda.
    """

    for chave in chaves_selecionadas:

        linhas_resultado.append(
            (
                id_lista,
                chave
            )
        )

    linhas_legenda.append(
        (
            id_lista,
            nome_lista,
            criterio_texto,
            len(chaves_selecionadas)
        )
    )


#%%
# =========================================================
# BLOCO 1
# 1 ITEM DE CADA CATEGORIA (LISTAS 1, 2 E 3)
# =========================================================
#
# Para cada uma das 13 categorias, seleciona-se exatamente
# 1 item, segundo o critério da lista. O resultado são 3
# listas de 13 itens cada (uma por categoria).

# ---------------------------------------------------------
# LISTA 1 - MAIS BARATO
# ---------------------------------------------------------

registrar_lista(
    id_lista=1,
    nome_lista="1 Item de cada - Mais barato",
    criterio_texto=(
        "Para cada uma das 13 categorias, item de menor "
        "preço encontrado, considerando todos os mercados."
    ),
    chaves_selecionadas=[
        item_mais_barato(
            df[df["Categoria Produto"] == categoria],
            categoria
        )
        for categoria in cats
    ]
)


# ---------------------------------------------------------
# LISTA 2 - MAIS CARO
# ---------------------------------------------------------

registrar_lista(
    id_lista=2,
    nome_lista="1 Item de cada - Mais caro",
    criterio_texto=(
        "Para cada uma das 13 categorias, item de maior "
        "preço encontrado, considerando todos os mercados."
    ),
    chaves_selecionadas=[
        item_mais_caro(
            df[df["Categoria Produto"] == categoria],
            categoria
        )
        for categoria in cats
    ]
)


# ---------------------------------------------------------
# LISTA 3 - MÉDIO (MAIS PRÓXIMO DA MEDIANA)
# ---------------------------------------------------------

registrar_lista(
    id_lista=3,
    nome_lista="1 Item de cada - Médio",
    criterio_texto=(
        "Para cada uma das 13 categorias, item cujo preço "
        "fica mais próximo da mediana de preços da "
        "categoria."
    ),
    chaves_selecionadas=[
        item_mediano(
            df[df["Categoria Produto"] == categoria],
            categoria
        )
        for categoria in cats
    ]
)


#%%
# =========================================================
# BLOCO 2
# CESTAS DE TAMANHOS DIFERENTES (LISTAS 4, 5 E 6)
# =========================================================
#
# Nestas listas, tanto a quantidade de categorias
# participantes quanto o item escolhido dentro de cada
# categoria são definidos por sorteio (random.sample /
# item_aleatorio), respeitando a semente fixa definida no
# início do script.

# ---------------------------------------------------------
# LISTA 4 - CESTA COM POUCOS ITENS (1 A 4)
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 3)

registrar_lista(
    id_lista=4,
    nome_lista="Cesta com poucos itens (1 a 4)",
    criterio_texto=(
        "3 categorias sorteadas aleatoriamente, com 1 item "
        "sorteado em cada."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria)
        for categoria in categorias_sorteadas
    ]
)


# ---------------------------------------------------------
# LISTA 5 - CESTA COM QUANTIDADE MÉDIA DE ITENS (5 A 10)
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 7)

registrar_lista(
    id_lista=5,
    nome_lista="Cesta com quantidade média de itens (5 a 10)",
    criterio_texto=(
        "7 categorias sorteadas aleatoriamente, com 1 item "
        "sorteado em cada."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria)
        for categoria in categorias_sorteadas
    ]
)


# ---------------------------------------------------------
# LISTA 6 - CESTA COM MUITOS ITENS (ACIMA DE 10)
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 12)

registrar_lista(
    id_lista=6,
    nome_lista="Cesta com muitos itens (Acima de 10)",
    criterio_texto=(
        "12 categorias sorteadas aleatoriamente (do total "
        "de 13 disponíveis), com 1 item sorteado em cada."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria)
        for categoria in categorias_sorteadas
    ]
)


#%%
# =========================================================
# BLOCO 3
# LISTAS VARIADAS (LISTAS 7 A 11)
# =========================================================
#
# Estas 5 listas combinam os critérios anteriores de formas
# não padronizadas, com o objetivo de ampliar a diversidade
# de cenários submetidos ao modelo de otimização - incluindo
# combinações que não se encaixam estritamente em "tamanho
# de cesta" ou "critério único de preço".

# ---------------------------------------------------------
# LISTA 7 - MIX DE CRITÉRIOS DE PREÇO
# ---------------------------------------------------------
#
# 5 categorias sorteadas, alternando manualmente entre item
# mais barato, mais caro e mediano - simula uma compra onde
# o consumidor economiza em alguns itens e não em outros.

categorias_sorteadas = random.sample(cats, 5)

criterios_alternados = [
    "barato",
    "caro",
    "mediano",
    "barato",
    "caro"
]

registrar_lista(
    id_lista=7,
    nome_lista="Lista Variada 1 - Mix barato/caro/mediano",
    criterio_texto=(
        "5 categorias sorteadas aleatoriamente, alternando "
        "manualmente entre item mais barato, mais caro e "
        "mediano."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria, criterio)
        for categoria, criterio in zip(
            categorias_sorteadas,
            criterios_alternados
        )
    ]
)


# ---------------------------------------------------------
# LISTA 8 - FOCO TEMÁTICO (PROTEÍNAS E CARBOIDRATOS)
# ---------------------------------------------------------
#
# Diferente das demais, esta lista NÃO sorteia as
# categorias: elas são fixadas manualmente, simulando uma
# compra temática (base proteica e energética da
# alimentação).

categorias_fixas = [
    "Carne",
    "Feijão",
    "Arroz",
    "Farinha de Trigo"
]

# Validação antecipada: como estes nomes foram digitados
# manualmente (não vêm de 'cats'), são o ponto mais provável
# de divergência com a planilha. Verifica todos de uma vez,
# antes de tentar montar a lista.
categorias_ausentes = [
    categoria
    for categoria in categorias_fixas
    if categoria not in cats
]

if categorias_ausentes:

    raise ValueError(
        f"\nAs seguintes categorias da Lista 8 não foram "
        f"encontradas na base: {categorias_ausentes!r}\n"
        f"Categorias disponíveis na planilha: {cats}\n"
        f"Verifique acentuação, espaços e maiúsculas/"
        f"minúsculas em 'categorias_fixas'."
    )

registrar_lista(
    id_lista=8,
    nome_lista="Lista Variada 2 - Proteínas e carboidratos",
    criterio_texto=(
        "Categorias fixas (Carne, Feijão, Arroz e Farinha), "
        "com 1 item sorteado em cada."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria)
        for categoria in categorias_fixas
    ]
)


# ---------------------------------------------------------
# LISTA 9 - CESTA COMPACTA PREMIUM (SOMENTE MAIS CAROS)
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 4)

registrar_lista(
    id_lista=9,
    nome_lista="Lista Variada 3 - Cesta compacta premium",
    criterio_texto=(
        "4 categorias sorteadas aleatoriamente, sempre "
        "utilizando o item mais caro de cada uma."
    ),
    chaves_selecionadas=[
        item_mais_caro(
            df[df["Categoria Produto"] == categoria],
            categoria
        )
        for categoria in categorias_sorteadas
    ]
)


# ---------------------------------------------------------
# LISTA 10 - CESTA AMPLA ECONÔMICA (SOMENTE MAIS BARATOS)
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 10)

registrar_lista(
    id_lista=10,
    nome_lista="Lista Variada 4 - Cesta ampla econômica",
    criterio_texto=(
        "10 categorias sorteadas aleatoriamente, sempre "
        "utilizando o item mais barato de cada uma."
    ),
    chaves_selecionadas=[
        item_mais_barato(
            df[df["Categoria Produto"] == categoria],
            categoria
        )
        for categoria in categorias_sorteadas
    ]
)


# ---------------------------------------------------------
# LISTA 11 - ALEATÓRIA MISTA
# ---------------------------------------------------------

categorias_sorteadas = random.sample(cats, 8)

registrar_lista(
    id_lista=11,
    nome_lista="Lista Variada 5 - Aleatória mista",
    criterio_texto=(
        "8 categorias sorteadas aleatoriamente, com 1 item "
        "sorteado (sem critério de preço) em cada."
    ),
    chaves_selecionadas=[
        item_aleatorio(categoria)
        for categoria in categorias_sorteadas
    ]
)


#%%
# =========================================================
# CONSOLIDAÇÃO DOS RESULTADOS
# =========================================================
#
# 'df_listas' é o entregável principal: exatamente as duas
# colunas solicitadas, prontas para servir de chave de
# consulta na base original de produtos e preços.
#
# 'df_legenda' é material de apoio/documentação, não é
# exigido pelo modelo de otimização.

df_listas = pd.DataFrame(
    linhas_resultado,
    columns=[
        "ID Lista de Compras",
        "Chave de Escolha"
    ]
)

df_legenda = pd.DataFrame(
    linhas_legenda,
    columns=[
        "ID Lista de Compras",
        "Nome da Lista",
        "Critério de Montagem",
        "Qtd. Itens"
    ]
)


#%%
# =========================================================
# VALIDAÇÃO
# =========================================================

print("\n")
print("=" * 100)
print("RESUMO DAS LISTAS GERADAS")
print("=" * 100)

print(
    df_legenda.to_string(index=False)
)

print(
    f"\nTotal de linhas (ID Lista x Chave de Escolha): "
    f"{len(df_listas)}"
)


#%%
# =========================================================
# EXPORTAÇÃO PARA EXCEL
# =========================================================
#
# O arquivo de saída contém duas abas:
#
# "Listas"  -> formato solicitado (ID Lista de Compras,
#              Chave de Escolha), pronto para uso como
#              entrada dos Scripts 01/03.
#
# "Legenda" -> documentação de apoio, explicando o critério
#              de montagem de cada ID de lista.

with pd.ExcelWriter(
    caminho_saida,
    engine="openpyxl"
) as writer:

    df_listas.to_excel(
        writer,
        sheet_name="Listas",
        index=False
    )

    df_legenda.to_excel(
        writer,
        sheet_name="Legenda",
        index=False
    )


#%%
# =========================================================
# FORMATAÇÃO VISUAL
# =========================================================
#
# Aplica formatação simples (cabeçalho destacado, fonte
# padronizada, largura de colunas e congelamento da
# primeira linha) às duas abas geradas, seguindo o mesmo
# padrão visual utilizado nos demais arquivos de saída do
# projeto (Scripts 03 e 04).

wb = load_workbook(caminho_saida)

fonte_cabecalho = Font(
    name="Arial",
    bold=True,
    color="FFFFFF"
)

fill_cabecalho = PatternFill(
    start_color="305496",
    end_color="305496",
    fill_type="solid"
)

fonte_corpo = Font(
    name="Arial"
)

alinhamento_centro = Alignment(
    horizontal="center",
    vertical="center"
)


for nome_aba in wb.sheetnames:

    ws = wb[nome_aba]

    # -----------------------------------------------------
    # CABEÇALHO
    # -----------------------------------------------------

    for celula in ws[1]:

        celula.font = fonte_cabecalho

        celula.fill = fill_cabecalho

        celula.alignment = alinhamento_centro


    # -----------------------------------------------------
    # CORPO
    # -----------------------------------------------------

    for linha in ws.iter_rows(min_row=2):

        for celula in linha:

            celula.font = fonte_corpo


    # -----------------------------------------------------
    # LARGURA DAS COLUNAS (AUTO-AJUSTE SIMPLES)
    # -----------------------------------------------------

    for coluna_celulas in ws.columns:

        maior_tamanho = max(
            len(str(celula.value))
            if celula.value is not None
            else 0
            for celula in coluna_celulas
        )

        letra_coluna = get_column_letter(
            coluna_celulas[0].column
        )

        ws.column_dimensions[
            letra_coluna
        ].width = min(
            max(maior_tamanho + 2, 10),
            60
        )


    # -----------------------------------------------------
    # CONGELAMENTO DO CABEÇALHO
    # -----------------------------------------------------

    ws.freeze_panes = "A2"


wb.save(caminho_saida)


#%%
# =========================================================
# FINAL
# =========================================================

print("\n")
print("=" * 100)
print("SCRIPT 00 (GERAÇÃO DE CENÁRIOS) FINALIZADO")
print("=" * 100)

print(
    f"\nSemente de aleatoriedade utilizada: "
    f"{SEMENTE_ALEATORIA}"
)

print(
    f"\nArquivo salvo em:\n{caminho_saida}"
)

