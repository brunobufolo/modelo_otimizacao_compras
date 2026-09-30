"""Validação por enumeração com buscas pré-calculadas.

Mantém as mesmas regras do script original, mas elimina filtros repetidos em
DataFrames dentro do laço de cenários.
"""

from datetime import datetime

import pandas as pd


inicio = datetime.now()
print("\n" + "=" * 100)
print(f"INÍCIO DO PROCESSAMENTO: {inicio.strftime('%H:%M:%S')}")
print("=" * 100)


from config import BASE_FONTE, ENUMERACAO, MATRIZ_ROTAS, BRUTE_FORCE
caminho_excel = BASE_FONTE
arquivo_enumeracao = ENUMERACAO
arquivo_matriz_rotas = MATRIZ_ROTAS
arquivo_saida = BRUTE_FORCE


# Leitura das bases
df_compras = pd.read_excel(caminho_excel, sheet_name="Compras")
df_produtos = pd.read_excel(caminho_excel, sheet_name="Produtos & Mercados")
df_partidas = pd.read_excel(caminho_excel, sheet_name="Pontos de Partida")
df_rotas = pd.read_excel(arquivo_enumeracao)
df_matriz = pd.read_excel(arquivo_matriz_rotas)

for df in (df_compras, df_produtos, df_partidas, df_rotas, df_matriz):
    df.columns = df.columns.str.strip()

df_produtos["Preço Final"] = pd.to_numeric(
    df_produtos["Preço Final"], errors="coerce"
)


# Estruturas de consulta criadas UMA vez. No script original, a matriz de
# rotas era filtrada até duas vezes para cada trecho de cada cenário.
pontos_partida = (
    df_partidas.drop_duplicates("ID Endereço")
    .set_index("ID Endereço")["Endereço Completo"]
    .to_dict()
)
mercado_para_id = (
    df_produtos[["Mercado", "ID Mercado"]]
    .drop_duplicates()
    .set_index("Mercado")["ID Mercado"]
    .to_dict()
)

precos = {
    (id_produto, id_mercado): preco
    for id_produto, id_mercado, preco in df_produtos[
        ["ID Produto", "ID Mercado", "Preço Final"]
    ].itertuples(index=False, name=None)
    if pd.notna(preco)
}

# Primeiro preserva todos os trechos explícitos. Só depois acrescenta a rota
# inversa quando ela não existe, exatamente como a função original fazia.
custos_rota = {
    (tipo_origem, id_origem, tipo_destino, id_destino): custo
    for tipo_origem, id_origem, tipo_destino, id_destino, custo in df_matriz[
        ["TIPO_ORIGEM", "ID_ORIGEM", "TIPO_DESTINO", "ID_DESTINO", "Custo Logistico"]
    ].itertuples(index=False, name=None)
}
for (tipo_origem, id_origem, tipo_destino, id_destino), custo in list(custos_rota.items()):
    custos_rota.setdefault(
        (tipo_destino, id_destino, tipo_origem, id_origem), custo
    )


colunas_mercados = sorted(
    [c for c in df_rotas.columns if c.startswith("M") and c[1:].isdigit()],
    key=lambda c: int(c[1:]),
)


# Pré-processa tudo que depende somente do cenário. Assim, para cada lista só
# restam o ponto de partida e a escolha do menor preço dos seus produtos.
cenarios = []
colunas_rota = ["ID_CENARIO", "QTD_MERCADOS", *colunas_mercados]
for valores in df_rotas[colunas_rota].itertuples(index=False, name=None):
    id_cenario, qtd_mercados, *valores_mercados = valores
    mercados = tuple(
        mercado
        for mercado in valores_mercados
        if pd.notna(mercado) and str(mercado).strip() != ""
    )
    ids_mercados = tuple(mercado_para_id[mercado] for mercado in mercados)
    custos_entre_mercados = tuple(
        custos_rota.get(("MERCADO", origem, "MERCADO", destino), 0)
        for origem, destino in zip(ids_mercados, ids_mercados[1:])
    )
    cenarios.append(
        (id_cenario, qtd_mercados, mercados, ids_mercados, custos_entre_mercados)
    )


# Agrupar compras evita filtrar df_compras novamente a cada lista.
compras_por_lista = {
    id_lista: list(zip(grupo["ID Produto"], grupo["Quantidade"].fillna(1)))
    for id_lista, grupo in df_compras.groupby("ID Lista", sort=False)
}
ponto_por_lista = (
    df_compras.drop_duplicates("ID Lista")
    .set_index("ID Lista")["ID Ponto de Partida"]
    .to_dict()
)


resultados = []
for id_lista, compras_lista in compras_por_lista.items():
    id_ponto = ponto_por_lista[id_lista]
    endereco_partida = pontos_partida.get(id_ponto, "")

    for id_cenario, qtd_mercados, mercados, ids_mercados, custos_internos in cenarios:
        # Custos que variam de acordo com o ponto de partida da lista.
        if ids_mercados:
            custo_saida = custos_rota.get(
                ("PARTIDA", id_ponto, "MERCADO", ids_mercados[0]), 0
            )
            custo_chegada = custos_rota.get(
                ("MERCADO", ids_mercados[-1], "PARTIDA", id_ponto), 0
            )
        else:
            custo_saida = custo_chegada = 0

        custos_mercados = (custo_saida, *custos_internos)
        custo_total_trajeto = sum(custos_mercados) + custo_chegada

        # A regra de empate continua a mesma: o primeiro mercado da rota é
        # escolhido quando dois preços são iguais.
        quantidade_itens = {mercado: 0 for mercado in mercados}
        menor_custo_lista = 0
        for id_produto, quantidade in compras_lista:
            menor_preco = None
            mercado_menor_preco = None
            for mercado, id_mercado in zip(mercados, ids_mercados):
                preco = precos.get((id_produto, id_mercado))
                if preco is not None and (menor_preco is None or preco < menor_preco):
                    menor_preco = preco
                    mercado_menor_preco = mercado

            if mercado_menor_preco is not None:
                menor_custo_lista += menor_preco * quantidade
                quantidade_itens[mercado_menor_preco] += 1

        linha_resultado = {
            "ID_LISTA": id_lista,
            "ID_CENARIO": id_cenario,
            "QTD_MERCADOS": qtd_mercados,
            "SAIDA": endereco_partida,
            "CHEGADA": endereco_partida,
            "ID Ponto de Partida": id_ponto,
            "Custo CHEGADA": round(custo_chegada, 2),
            "Custo Total Trajeto": round(custo_total_trajeto, 2),
            "Menor Custo Lista": round(menor_custo_lista, 2),
            "Custo Total": round(custo_total_trajeto + menor_custo_lista, 2),
        }

        for indice, coluna in enumerate(colunas_mercados):
            if indice < len(mercados):
                linha_resultado[coluna] = mercados[indice]
                linha_resultado[f"Custo {coluna}"] = round(custos_mercados[indice], 2)
                linha_resultado[f"Itens {coluna}"] = quantidade_itens[mercados[indice]]
            else:
                linha_resultado[coluna] = ""
                linha_resultado[f"Itens {coluna}"] = 0

        resultados.append(linha_resultado)


df_resultado = pd.DataFrame(resultados)
colunas_finais = [
    "ID_LISTA", "ID_CENARIO", "QTD_MERCADOS", "SAIDA",
    *colunas_mercados,
    "CHEGADA", "ID Ponto de Partida",
    *[f"Custo {coluna}" for coluna in colunas_mercados],
    "Custo CHEGADA", "Custo Total Trajeto",
    "Menor Custo Lista", "Custo Total",
    *[f"Itens {coluna}" for coluna in colunas_mercados],
]
df_resultado = df_resultado[colunas_finais]

df_top_1 = (
    df_resultado
    .sort_values(by=["ID_LISTA", "Custo Total", "ID_CENARIO"])
    .groupby("ID_LISTA", group_keys=False)
    .head(1)
    .copy()
)
df_top_1.to_excel(arquivo_saida, sheet_name="Top 1 por Lista", index=False)


fim = datetime.now()
print("\n" + "=" * 100)
print("PROCESSAMENTO FINALIZADO")
print("=" * 100)
print(f"Total de soluções calculadas: {len(df_resultado)}")
print(f"Total de soluções exportadas (top 1 por lista): {len(df_top_1)}")
print(f"\nArquivo salvo em:\n{arquivo_saida}")
print(f"\nFIM DO PROCESSAMENTO: {fim.strftime('%H:%M:%S')}")
print(f"TEMPO TOTAL DE PROCESSAMENTO: {fim - inicio}")
