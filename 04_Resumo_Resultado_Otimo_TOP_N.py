# =========================================================
# SCRIPT 04
# INDICADORES DE RESUMO + EXPORTAÇÃO PARA ANÁLISE
# =========================================================
#
# Este script lê o resultado consolidado do Script 03
# (Resultado_Otimo_Listas.xlsx, aba "Resumo Geral") e produz:
#
# 1. KPIs agregados (cenários, economia, logística, produtos);
# 2. uma tabela por lista com indicadores derivados;
# 3. análise cruzada por Ponto de Partida (localização);
# 4. análise cruzada por Porte da Compra (Básica/Moderada/
#    Alto Volume, classificado por tercis de Qtd Produtos);
# 5. uma aba de notas metodológicas, explicando cada fórmula
#    e as limitações dos dados de origem.
#
#

#
# Todos os KPIs financeiros deste script usam a Economia
# Líquida Realizada, por ser o valor que reflete dinheiro
# efetivamente economizado.
#
# =========================================================


#%%
# =========================================================
# IMPORTS
# =========================================================

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter


#%%
# =========================================================
# CAMINHOS
# =========================================================

from config import RESULTADO_OTIMO_TOP_N, INDICADORES_TOP_N
caminho_entrada = RESULTADO_OTIMO_TOP_N
caminho_saida = INDICADORES_TOP_N

#%%
# =========================================================
# LEITURA DA BASE
# =========================================================

df = pd.read_excel(
    caminho_entrada,
    sheet_name="Resumo Geral"
)


#%%
# =========================================================
# INDICADORES POR LISTA (COLUNAS DERIVADAS)
# =========================================================
#
# Todas as colunas abaixo são calculadas a partir de colunas
# já existentes na planilha de origem - nenhuma reprocessa a
# lógica de seleção de mercados ou de rota.

# ---------------------------------------------------------
# Multimercado = 0/1
# ---------------------------------------------------------

df["Multimercado (0/1)"] = (
    df["Estratégia Recomendada"] == "MULTIMERCADO"
).astype(int)


# ---------------------------------------------------------
# Distância Total (km)
# ---------------------------------------------------------
#
# A premissa do modelo (Script 03) é R$ 1,00 por km, portanto
# o "Custo Deslocamento" em reais equivale numericamente à
# distância percorrida em quilômetros.

df["Distância Total (km)"] = (
    df["Custo Deslocamento"]
)


# ---------------------------------------------------------
# Distância Média aos Mercados (km)
# ---------------------------------------------------------

df["Distância Média aos Mercados (km)"] = (
    df["Distância Total (km)"]
    /
    df["Qtd Mercados"]
)


# ---------------------------------------------------------
# Participação Logística (%)
# ---------------------------------------------------------
#
# Qual fração do custo final recomendado é consumida pelo
# deslocamento.

df["Participação Logística (%)"] = (
    df["Custo Deslocamento"]
    /
    df["Custo Final Recomendado"]
)


# ---------------------------------------------------------
# Economia Líquida Realizada (R$)
# ---------------------------------------------------------
#
# Ver explicação completa no cabeçalho do script.

df["Economia Líquida Realizada (R$)"] = (
    df["Economia Total"].where(
        df["Economia Total"] > 0,
        0
    )
)


# ---------------------------------------------------------
# Custo Cesta Mercado Único / Custo Cesta Otimizada
# ---------------------------------------------------------
#
# Aliases mais descritivos das colunas já existentes, para
# facilitar a leitura em uma análise estatística posterior.

df["Custo Cesta Mercado Único (R$)"] = (
    df["Custo Compra Única"]
)

df["Custo Cesta Otimizada (R$)"] = (
    df["Custo Final Recomendado"]
)


# ---------------------------------------------------------
# Produtos por Mercado (média)
# ---------------------------------------------------------

df["Produtos por Mercado (média)"] = (
    df["Qtd Produtos"]
    /
    df["Qtd Mercados"]
)


# ---------------------------------------------------------
# PORTE DA COMPRA (Básica / Moderada / Alto Volume)
# ---------------------------------------------------------
#
# Classificação por tercis (33% / 66%) da Qtd Produtos,
# calculados sobre a própria base de listas processadas -
# não são limiares fixos definidos a priori.
#
# Em bases pequenas ou com muitos valores repetidos, pd.qcut
# pode não conseguir gerar 3 faixas distintas. Nesse caso,
# a função abaixo cai para uma classificação manual baseada
# nos mesmos percentis.

def classificar_porte_compra(serie_qtd_produtos):

    try:

        categorias = pd.qcut(
            serie_qtd_produtos,
            q=3,
            labels=["Básica", "Moderada", "Alto Volume"],
            duplicates="drop"
        )

        if categorias.nunique() < 3:

            raise ValueError(
                "Poucos valores distintos para 3 faixas."
            )

    except ValueError:

        p33, p66 = serie_qtd_produtos.quantile([0.333, 0.667])

        def rotular(qtd):

            if qtd <= p33:

                return "Básica"

            elif qtd <= p66:

                return "Moderada"

            else:

                return "Alto Volume"

        categorias = serie_qtd_produtos.apply(rotular)

    return categorias


df["Porte da Compra"] = classificar_porte_compra(
    df["Qtd Produtos"]
)


#%%
# =========================================================
# KPIs AGREGADOS
# =========================================================

total_cenarios = len(df)

total_com_economia = int(
    (df["Economia Total"] > 0).sum())

    
pct_com_economia = (
    total_com_economia / total_cenarios
    if total_cenarios else 0
)

# Taxa de Viabilidade Econômica (TVE) - mesmo cálculo acima,
# nome técnico usado na literatura de apoio do trabalho.
tve = pct_com_economia

soma_economia_realizada = (
    df["Economia Líquida Realizada (R$)"].sum()
)

economia_media_por_compra = (
    soma_economia_realizada / total_cenarios
    if total_cenarios else 0
)

soma_economia_bruta_referencia = (
    df["Economia Total"].sum()
)

valor_total_listas = (
    df["Custo Final Recomendado"].sum()
)

valor_medio_lista = (
    valor_total_listas / total_cenarios
    if total_cenarios else 0
)

qtd_media_mercados = (
    df["Qtd Mercados"].mean()
)

participacao_logistica_media = (
    df["Participação Logística (%)"].mean()
)

participacao_logistica_ponderada = (
    df["Custo Deslocamento"].sum()
    /
    df["Custo Final Recomendado"].sum()
)

distancia_total_geral = (
    df["Distância Total (km)"].sum()
)

distancia_media_geral = (
    df["Distância Total (km)"].mean()
)

soma_qtd_produtos = (
    df["Qtd Produtos"].sum()
)

media_produtos_por_compra = (
    df["Qtd Produtos"].mean()
)



# Até aqui está ok validado - Double check mas está ok



#%%
# =========================================================
# ANÁLISE POR LOCALIZAÇÃO (PONTO DE PARTIDA)
# =========================================================

analise_local = (
    df.groupby("Ponto de Partida")
    .agg(
        Qtd_Cenarios=("ID Lista", "count"),
        Qtd_Multimercado=("Multimercado (0/1)", "sum"),
        Distancia_Media=("Distância Total (km)", "mean"),
        Custo_Medio_Final=("Custo Final Recomendado", "mean"),
        Economia_Media_Realizada=(
            "Economia Líquida Realizada (R$)", "mean"
        ),
    )
    .reset_index()
)

analise_local["% Multimercado"] = (
    analise_local["Qtd_Multimercado"]
    /
    analise_local["Qtd_Cenarios"]
)

colunas_local = [
    "Ponto de Partida",
    "Qtd_Cenarios",
    "Qtd_Multimercado",
    "% Multimercado",
    "Distancia_Media",
    "Custo_Medio_Final",
    "Economia_Media_Realizada"
]

analise_local = analise_local[colunas_local]


#%%
# =========================================================
# ANÁLISE POR PORTE DA COMPRA
# =========================================================

analise_porte = (
    df.groupby("Porte da Compra", observed=True)
    .agg(
        Qtd_Cenarios=("ID Lista", "count"),
        Qtd_Multimercado=("Multimercado (0/1)", "sum"),
        Qtd_Produtos_Media=("Qtd Produtos", "mean"),
        Qtd_Mercados_Media=("Qtd Mercados", "mean"),
        Economia_Media_Realizada=(
            "Economia Líquida Realizada (R$)", "mean"
        ),
    )
    .reset_index()
)

analise_porte["% Multimercado"] = (
    analise_porte["Qtd_Multimercado"]
    /
    analise_porte["Qtd_Cenarios"]
)

ordem_porte = ["Básica", "Moderada", "Alto Volume"]

analise_porte["Porte da Compra"] = pd.Categorical(
    analise_porte["Porte da Compra"],
    categories=ordem_porte,
    ordered=True
)

analise_porte = (
    analise_porte
    .sort_values("Porte da Compra")
    .reset_index(drop=True)
)

colunas_porte = [
    "Porte da Compra",
    "Qtd_Cenarios",
    "Qtd_Multimercado",
    "% Multimercado",
    "Qtd_Produtos_Media",
    "Qtd_Mercados_Media",
    "Economia_Media_Realizada"
]

analise_porte = analise_porte[colunas_porte]


#%%
# =========================================================
# CRIAÇÃO DO WORKBOOK
# =========================================================

wb = Workbook()

ws_inicial = wb.active

wb.remove(ws_inicial)


# =========================================================
# ESTILOS
# =========================================================

fonte_padrao = "Arial"

fonte_titulo = Font(
    name=fonte_padrao,
    bold=True,
    size=16
)

fonte_subtitulo = Font(
    name=fonte_padrao,
    bold=True,
    size=12
)

fonte_cabecalho = Font(
    name=fonte_padrao,
    bold=True,
    color="FFFFFF"
)

fonte_kpi_label = Font(
    name=fonte_padrao,
    bold=True
)

fonte_kpi_valor = Font(
    name=fonte_padrao,
    size=14,
    bold=True,
    color="1F4E78"
)

fonte_normal = Font(
    name=fonte_padrao
)

fill_cabecalho = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

fill_multimercado = PatternFill(
    start_color="C6EFCE",
    end_color="C6EFCE",
    fill_type="solid"
)

fill_unico = PatternFill(
    start_color="F2F2F2",
    end_color="F2F2F2",
    fill_type="solid"
)

borda = Border(
    bottom=Side(style="thin"),
    top=Side(style="thin"),
    left=Side(style="thin"),
    right=Side(style="thin")
)

alinhamento_centro = Alignment(
    horizontal="center",
    vertical="center"
)

alinhamento_esquerda = Alignment(
    horizontal="left",
    vertical="center"
)


def aplicar_cabecalho(ws, linha, colunas, col_inicial=1):

    for i, cabecalho in enumerate(colunas, start=col_inicial):

        celula = ws.cell(row=linha, column=i, value=cabecalho)

        celula.font = fonte_cabecalho

        celula.fill = fill_cabecalho

        celula.alignment = alinhamento_centro

        celula.border = borda


#%%
# =========================================================
# ABA 1 - RESUMO EXECUTIVO
# =========================================================

ws_exec = wb.create_sheet(title="Resumo Executivo", index=0)

ws_exec["A1"] = "RESUMO EXECUTIVO - INDICADORES DE OTIMIZAÇÃO DE COMPRAS"
ws_exec["A1"].font = fonte_titulo
ws_exec.merge_cells("A1:D1")


def escrever_bloco_kpi(ws, linha_inicial, titulo, pares):
    """
    Escreve um bloco de KPIs no formato:

    TÍTULO DO BLOCO
    Rótulo 1     Valor 1
    Rótulo 2     Valor 2

    'pares' é uma lista de tuplas (rótulo, valor, formato_numero).
    Retorna a próxima linha livre após o bloco.
    """

    ws.cell(row=linha_inicial, column=1, value=titulo).font = (
        fonte_subtitulo
    )

    linha = linha_inicial + 1

    for rotulo, valor, formato in pares:

        ws.cell(row=linha, column=1, value=rotulo).font = (
            fonte_kpi_label
        )

        celula_valor = ws.cell(row=linha, column=3, value=valor)

        celula_valor.font = fonte_kpi_valor

        if formato:

            celula_valor.number_format = formato

        linha += 1

    return linha + 1


linha_atual = 3

linha_atual = escrever_bloco_kpi(
    ws_exec,
    linha_atual,
    "CENÁRIOS",
    [
        (
            "Total de Cenários (Compras)",
            total_cenarios,
            "#,##0"
        ),
        (
            "Cenários com Economia (Multimercado)",
            total_com_economia,
            "#,##0"
        ),
        (
            "% de Cenários com Economia",
            pct_com_economia,
            "0.00%"
        ),
        (
            "Taxa de Viabilidade Econômica (TVE)",
            tve,
            "0.00%"
        ),
    ]
)

linha_atual = escrever_bloco_kpi(
    ws_exec,
    linha_atual,
    "INDICADORES FINANCEIROS",
    [
        (
            "Economia Líquida Realizada (Total)",
            soma_economia_realizada,
            'R$ #,##0.00'
        ),
        (
            "Economia Líquida / Compras (média)",
            economia_media_por_compra,
            'R$ #,##0.00'
        ),
        (
            "Economia Bruta de Referência (Total)*",
            soma_economia_bruta_referencia,
            'R$ #,##0.00'
        ),
        (
            "Valor Total de Todas as Listas",
            valor_total_listas,
            'R$ #,##0.00'
        ),
        (
            "Valor das Listas / Compras (média)",
            valor_medio_lista,
            'R$ #,##0.00'
        ),
    ]
)

linha_atual = escrever_bloco_kpi(
    ws_exec,
    linha_atual,
    "INDICADORES LOGÍSTICOS",
    [
        (
            "Quantidade Média de Mercados por Compra",
            qtd_media_mercados,
            "0.00"
        ),
        (
            "Distância Total Percorrida (km)",
            distancia_total_geral,
            "#,##0.0"
        ),
        (
            "Distância Média por Compra (km)",
            distancia_media_geral,
            "#,##0.0"
        ),
        (
            "Participação Logística Média (%)",
            participacao_logistica_media,
            "0.00%"
        ),
        (
            "Participação Logística Ponderada (%)**",
            participacao_logistica_ponderada,
            "0.00%"
        ),
    ]
)

linha_atual = escrever_bloco_kpi(
    ws_exec,
    linha_atual,
    "INDICADORES DE PRODUTOS",
    [
        (
            "Total de Produtos Comprados (linhas de item)",
            soma_qtd_produtos,
            "#,##0"
        ),
        (
            "Produtos por Compra (média)",
            media_produtos_por_compra,
            "0.00"
        ),
    ]
)

ws_exec.cell(
    row=linha_atual,
    column=1,
    value=(
        "* Inclui o custo evitado nos cenários MERCADO ÚNICO "
        "(ver aba 'Notas Metodológicas')."
    )
).font = Font(name=fonte_padrao, italic=True, size=9)

linha_atual += 1

ws_exec.cell(
    row=linha_atual,
    column=1,
    value=(
        "** Soma do Custo Deslocamento / Soma do Custo Final "
        "Recomendado (dá mais peso às compras de maior valor)."
    )
).font = Font(name=fonte_padrao, italic=True, size=9)

linha_atual += 3


# ---------------------------------------------------------
# DADOS PARA GRÁFICO: CENÁRIOS POR ESTRATÉGIA
# ---------------------------------------------------------

linha_grafico_estrategia = linha_atual

ws_exec.cell(
    row=linha_grafico_estrategia,
    column=1,
    value="Mercado Único"
)

ws_exec.cell(
    row=linha_grafico_estrategia,
    column=2,
    value=total_cenarios - total_com_economia
)

ws_exec.cell(
    row=linha_grafico_estrategia + 1,
    column=1,
    value="Multimercado"
)

ws_exec.cell(
    row=linha_grafico_estrategia + 1,
    column=2,
    value=total_com_economia
)

grafico_estrategia = BarChart()

grafico_estrategia.title = "Cenários por Estratégia Recomendada"

grafico_estrategia.style = 10

grafico_estrategia.y_axis.title = "Qtd de Cenários"

dados_grafico_1 = Reference(
    ws_exec,
    min_col=2,
    min_row=linha_grafico_estrategia,
    max_row=linha_grafico_estrategia + 1
)

categorias_grafico_1 = Reference(
    ws_exec,
    min_col=1,
    min_row=linha_grafico_estrategia,
    max_row=linha_grafico_estrategia + 1
)

grafico_estrategia.add_data(dados_grafico_1, titles_from_data=False)

grafico_estrategia.set_categories(categorias_grafico_1)

grafico_estrategia.legend = None

ws_exec.add_chart(grafico_estrategia, "E3")


# ---------------------------------------------------------
# LARGURA DE COLUNAS
# ---------------------------------------------------------

ws_exec.column_dimensions["A"].width = 42
ws_exec.column_dimensions["B"].width = 4
ws_exec.column_dimensions["C"].width = 18


#%%
# =========================================================
# ABA 2 - INDICADORES POR LISTA
# =========================================================

ws_lista = wb.create_sheet(title="Indicadores por Lista")

colunas_export_lista = [
    "ID Lista",
    "Ponto de Partida",
    "Mercado Inicial",
    "Estratégia Recomendada",
    "Multimercado (0/1)",
    "Porte da Compra",
    "Qtd Mercados",
    "Qtd Produtos",
    "Produtos por Mercado (média)",
    "Custo Cesta Mercado Único (R$)",
    "Custo Cesta Otimizada (R$)",
    "Economia Total",
    "Economia Líquida Realizada (R$)",
    "Economia %",
    "Distância Total (km)",
    "Distância Média aos Mercados (km)",
    "Participação Logística (%)",
    "Sequência Ótima"
]

aplicar_cabecalho(ws_lista, 1, colunas_export_lista)

colunas_monetarias = {
    "Custo Cesta Mercado Único (R$)",
    "Custo Cesta Otimizada (R$)",
    "Economia Total",
    "Economia Líquida Realizada (R$)"
}

colunas_percentuais = {
    "Economia %",
    "Participação Logística (%)"
}

colunas_decimais = {
    "Produtos por Mercado (média)",
    "Distância Total (km)",
    "Distância Média aos Mercados (km)"
}

for i, (_, linha) in enumerate(
    df[colunas_export_lista].iterrows(),
    start=2
):

    for j, coluna in enumerate(colunas_export_lista, start=1):

        valor = linha[coluna]

        celula = ws_lista.cell(row=i, column=j, value=valor)

        celula.border = borda

        if coluna in colunas_monetarias:

            celula.number_format = 'R$ #,##0.00'
            celula.alignment = alinhamento_centro

        elif coluna in colunas_percentuais:

            celula.number_format = '0.00%'
            celula.alignment = alinhamento_centro

        elif coluna in colunas_decimais:

            celula.number_format = '#,##0.00'
            celula.alignment = alinhamento_centro

        elif coluna in ("Ponto de Partida", "Mercado Inicial", "Sequência Ótima"):

            celula.alignment = alinhamento_esquerda

        else:

            celula.alignment = alinhamento_centro

    # Destaque visual da estratégia vencedora.
    col_estrategia = colunas_export_lista.index(
        "Estratégia Recomendada"
    ) + 1

    fill = (
        fill_multimercado
        if linha["Estratégia Recomendada"] == "MULTIMERCADO"
        else fill_unico
    )

    for j in range(1, len(colunas_export_lista) + 1):

        ws_lista.cell(row=i, column=j).fill = fill


ws_lista.freeze_panes = "A2"

ws_lista.auto_filter.ref = (
    f"A1:{get_column_letter(len(colunas_export_lista))}{len(df) + 1}"
)

larguras_lista = {
    "A": 10, "B": 42, "C": 20, "D": 20, "E": 14, "F": 14,
    "G": 12, "H": 12, "I": 18, "J": 20, "K": 18, "L": 16,
    "M": 22, "N": 12, "O": 16, "P": 22, "Q": 20, "R": 45
}

for coluna, largura in larguras_lista.items():

    ws_lista.column_dimensions[coluna].width = largura


#%%
# =========================================================
# ABA 3 - ANÁLISE POR LOCALIZAÇÃO
# =========================================================

ws_local = wb.create_sheet(title="Análise por Localização")

ws_local["A1"] = (
    "ANÁLISE POR PONTO DE PARTIDA (LOCALIZAÇÃO DO CONSUMIDOR)"
)
ws_local["A1"].font = fonte_subtitulo
ws_local.merge_cells("A1:G1")

cabecalhos_local = [
    "Ponto de Partida",
    "Qtd Cenários",
    "Qtd Multimercado",
    "% Multimercado",
    "Distância Média (km)",
    "Custo Médio Final (R$)",
    "Economia Líquida Média (R$)"
]

aplicar_cabecalho(ws_local, 3, cabecalhos_local)

for i, (_, linha) in enumerate(analise_local.iterrows(), start=4):

    for j, valor in enumerate(linha, start=1):

        celula = ws_local.cell(row=i, column=j, value=valor)

        celula.border = borda

        if j == 4:

            celula.number_format = '0.00%'

        elif j in (5,):

            celula.number_format = '#,##0.00'

        elif j in (6, 7):

            celula.number_format = 'R$ #,##0.00'

        if j == 1:

            celula.alignment = alinhamento_esquerda

        else:

            celula.alignment = alinhamento_centro

larguras_local = {
    "A": 45, "B": 14, "C": 16, "D": 14,
    "E": 18, "F": 20, "G": 22
}

for coluna, largura in larguras_local.items():

    ws_local.column_dimensions[coluna].width = largura


#%%
# =========================================================
# ABA 4 - ANÁLISE POR PORTE DA COMPRA
# =========================================================

ws_porte = wb.create_sheet(title="Análise por Porte de Compra")

ws_porte["A1"] = (
    "ANÁLISE POR PORTE DA COMPRA "
    "(TERCIS DE QUANTIDADE DE PRODUTOS)"
)
ws_porte["A1"].font = fonte_subtitulo
ws_porte.merge_cells("A1:G1")

cabecalhos_porte = [
    "Porte da Compra",
    "Qtd Cenários",
    "Qtd Multimercado",
    "% Multimercado",
    "Qtd Produtos (média)",
    "Qtd Mercados (média)",
    "Economia Líquida Média (R$)"
]

aplicar_cabecalho(ws_porte, 3, cabecalhos_porte)

for i, (_, linha) in enumerate(analise_porte.iterrows(), start=4):

    for j, valor in enumerate(linha, start=1):

        celula = ws_porte.cell(row=i, column=j, value=valor)

        celula.border = borda

        if j == 4:

            celula.number_format = '0.00%'

        elif j in (5, 6):

            celula.number_format = '#,##0.00'

        elif j == 7:

            celula.number_format = 'R$ #,##0.00'

        if j == 1:

            celula.alignment = alinhamento_esquerda

        else:

            celula.alignment = alinhamento_centro

larguras_porte = {
    "A": 18, "B": 14, "C": 16, "D": 14,
    "E": 18, "F": 18, "G": 22
}

for coluna, largura in larguras_porte.items():

    ws_porte.column_dimensions[coluna].width = largura


# ---------------------------------------------------------
# GRÁFICO: ECONOMIA MÉDIA POR PORTE
# ---------------------------------------------------------

linha_final_porte = 3 + len(analise_porte)

grafico_porte = BarChart()

grafico_porte.title = "Economia Líquida Média por Porte de Compra"

grafico_porte.style = 12

grafico_porte.y_axis.title = "R$"

dados_grafico_2 = Reference(
    ws_porte,
    min_col=7,
    min_row=3,
    max_row=linha_final_porte
)

categorias_grafico_2 = Reference(
    ws_porte,
    min_col=1,
    min_row=4,
    max_row=linha_final_porte
)

grafico_porte.add_data(dados_grafico_2, titles_from_data=True)

grafico_porte.set_categories(categorias_grafico_2)

ws_porte.add_chart(grafico_porte, "I3")


#%%
# =========================================================
# ABA 5 - NOTAS METODOLÓGICAS
# =========================================================

ws_notas = wb.create_sheet(title="Notas Metodológicas")

ws_notas["A1"] = "NOTAS METODOLÓGICAS"
ws_notas["A1"].font = fonte_titulo
ws_notas.merge_cells("A1:B1")

ws_notas.column_dimensions["A"].width = 42
ws_notas.column_dimensions["B"].width = 90

notas = [

    (
        "Distância Total (km)",
        "Igual ao Custo Deslocamento, pois o Script 03 adota "
        "a premissa de R$ 1,00 por km percorrido."
    ),

    (
        "Multimercado (0/1)",
        "1 quando a Estratégia Recomendada é MULTIMERCADO, "
        "0 quando é MERCADO ÚNICO."
    ),

    (
        "Participação Logística (%)",
        "Custo Deslocamento dividido pelo Custo Final "
        "Recomendado. Indica qual fração do valor pago pela "
        "compra é consumida pelo deslocamento."
    ),

    (
        "Economia Total (coluna de origem)",
        "Sempre não-negativa. Em cenários MULTIMERCADO, "
        "representa economia realmente capturada (Único - "
        "Multimercado). Em cenários MERCADO ÚNICO, representa "
        "o custo que seria incorrido a mais se a compra tivesse "
        "sido multimercado (Multimercado - Único) - ou seja, um "
        "custo evitado, não uma economia realizada."
    ),

    (
        "Economia Líquida Realizada (R$)",
        "Economia Total para ambos cenários, considerando o seguinte formato"
        "Mercado Único -> Custo Multimercado - Custo Mercado Único"
        "MultiMercado -> Custo Mercado Único - Custo Multimercado"
    ),

    (
        "Taxa de Viabilidade Econômica (TVE)",
        "Qtd de cenários MULTIMERCADO dividido pelo total de "
        "cenários. Equivalente a '% de Cenários com Economia'."
    ),

    (
        "Porte da Compra",
        "Classificação em Básica / Moderada / Alto Volume por "
        "tercis (33% / 66%) da Qtd Produtos, calculados sobre a "
        "própria base processada - não são limiares fixos "
        "predefinidos."
    ),

    (
        "Participação Logística Ponderada",
        "Soma do Custo Deslocamento dividida pela soma do Custo "
        "Final Recomendado (todas as listas). Difere da média "
        "simples por dar mais peso às compras de maior valor."
    ),

    (
        "LIMITAÇÃO - Economia de Produtos (isolada da logística)",
        "Não calculável com a planilha de origem atual. Quando a "
        "estratégia final é MERCADO ÚNICO, o Script 03 sobrescreve "
        "o custo de produtos multimercado (tentativa) pelo custo "
        "de produtos do M1, perdendo o valor bruto da tentativa. "
        "Para viabilizar esse indicador, adicione ao dicionário "
        "'resultados_finais.append({...})' do Script 03 as chaves "
        "'Custo Produtos M1' (valor da variável custo_produtos_m1) "
        "e 'Custo Produtos Multimercado (Bruto)' (valor de "
        "custo_lista_total ANTES do bloco de sincronização com a "
        "estratégia final)."
    ),

    (
        "LIMITAÇÃO - Consumo da Economia pela Logística",
        "Depende do indicador anterior (Economia de Produtos). "
        "Fórmula pretendida: Custo Deslocamento / Economia Bruta "
        "de Produtos."
    ),

    (
        "LIMITAÇÃO - Dispersão de Preços",
        "Exigiria os preços de cada produto em cada mercado "
        "(nível de item), não disponíveis na aba 'Resumo Geral'. "
        "Pode ser calculada a partir de df_escolha_mercados no "
        "Script 03, caso deseje uma versão futura deste relatório "
        "com esse detalhamento."
    ),

    (
        "LIMITAÇÃO - Distância Máxima (entre mercados na rota)",
        "Exigiria os trechos individuais da rota ótima, não "
        "disponíveis de forma agregada na aba 'Resumo Geral'."
    ),

]

linha_nota = 3

for titulo, texto in notas:

    ws_notas.cell(row=linha_nota, column=1, value=titulo).font = (
        fonte_kpi_label
    )

    ws_notas.cell(row=linha_nota, column=1).alignment = Alignment(
        vertical="top", wrap_text=True
    )

    celula_texto = ws_notas.cell(row=linha_nota, column=2, value=texto)

    celula_texto.font = fonte_normal

    celula_texto.alignment = Alignment(
        vertical="top",
        wrap_text=True
    )

    ws_notas.row_dimensions[linha_nota].height = 60

    linha_nota += 1


#%%
# =========================================================
# SALVAMENTO
# =========================================================

wb.save(caminho_saida)

print("\n")
print("=" * 100)
print("SCRIPT 04 FINALIZADO")
print("=" * 100)

print(
    f"\nTotal de cenários processados: {total_cenarios}"
)

print(
    f"Cenários com economia (multimercado): {total_com_economia} "
    f"({pct_com_economia:.2%})"
)

print(
    f"\nArquivo de indicadores salvo em:\n{caminho_saida}"
)