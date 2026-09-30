# =========================================================
# EXPLICATIVO RESUMIDO
# =========================================================

# Este arquivo tem como finalidade construir a matriz de
# deslocamento entre os pontos considerados no estudo, que servirá como base auxiliar de forma a facilitar cálculos posteriores.
# contemplando:

# 1. Pontos de partida → Mercados
# 2. Mercados → Mercados

# A matriz resultante contém, para cada par de pontos,
# a distância estimada e o respectivo tempo de deslocamento.

# A distância é inicialmente calculada pelo método de
# Haversine, que determina a distância em linha reta entre
# duas coordenadas geográficas.

# Como a distância em linha reta não representa o percurso
# efetivamente realizado pela malha viária urbana, é aplicado
# posteriormente um fator de aproximação da malha urbana.

# O tempo de deslocamento é estimado a partir da distância
# ajustada, da velocidade média adotada e do fator de trânsito.

# Esta matriz constitui uma das bases utilizadas nos scripts
# posteriores para calcular o custo de deslocamento das
# diferentes soluções de compra.

#%%
# =========================================================
# IMPORTS
# =========================================================

import pandas as pd

from math import radians
from math import sin
from math import cos
from math import sqrt
from math import atan2


#%% 
# =========================================================
# CAMINHO EXCEL
# =========================================================

# Define o caminho da base de dados utilizada como entrada.

from config import BASE_FONTE, MATRIZ_ROTAS
caminho_excel = BASE_FONTE
output = MATRIZ_ROTAS

#%%
# =========================================================
# LEITURA DAS ABAS
# =========================================================

# Leitura da aba contendo os mercados.

# São utilizadas as coordenadas geográficas de cada mercado
# para posteriormente calcular as distâncias entre os pontos.
df_mercados = pd.read_excel(
    caminho_excel,
    sheet_name="Base Mercado")


# Leitura da aba contendo os pontos de partida.

# Cada ponto representa uma possível localização de origem
# para o consumidor realizar a compra.
df_partidas = pd.read_excel(
    caminho_excel,
    sheet_name="Pontos de Partida")


#%%
# =========================================================
# PADRONIZAÇÃO DAS COLUNAS
# =========================================================

# Remove espaços extras presentes nos nomes das colunas.

# Essa padronização reduz o risco de erros decorrentes de
# espaços acidentais nos nomes das variáveis utilizadas
# posteriormente no código.
df_mercados.columns = (
    df_mercados.columns
    .str.strip())

df_partidas.columns = (
    df_partidas.columns
    .str.strip())


#%%
# =========================================================
# HAVERSINE
# =========================================================

# Calcula a distância geodésica (linha mais curta que une dois pontos em uma superfície curva)
# aproximada entre dois pontos
# utilizando suas coordenadas de latitude e longitude.

# O método de Haversine considera a curvatura da Terra e
# fornece a distância entre os pontos em linha reta sobre
# a superfície terrestre.

# Essa distância representa o ponto de partida do modelo.
# Posteriormente, ela é ajustada pelo fator de malha urbana
# para representar de forma aproximada o deslocamento pela
# infraestrutura viária.
def haversine(lat1, lon1, lat2, lon2):

    # Raio médio da Terra em quilômetros.
    R = 6371

    # Diferenças entre as coordenadas dos dois pontos,
    # convertidas para radianos.
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)

    # Componente da fórmula de Haversine responsável por
    # relacionar as diferenças de latitude e longitude
    # considerando a curvatura terrestre.
    a = (
        sin(dlat / 2) ** 2
        + cos(radians(lat1))
        * cos(radians(lat2))
        * sin(dlon / 2) ** 2)

    # Converte o resultado anterior para o ângulo central
    # utilizado no cálculo da distância.
    c = 2 * atan2(
        sqrt(a),
        sqrt(1 - a))
    
    # Retorna a distância aproximada entre os pontos em km.
    return R * c


#%%
# =========================================================
# PARÂMETROS MODELO
# =========================================================

# Parâmetros utilizados para aproximar o deslocamento
# urbano a partir da distância em linha reta.

# ---------------------------------------------------------
# FATOR_MALHA_URBANA
# ---------------------------------------------------------

# O fator de malha urbana representa o acréscimo esperado
# na distância efetivamente percorrida em relação à
# distância em linha reta.

# Esse ajuste busca representar características da
# infraestrutura viária, como:

# - traçado das vias;
# - curvas;
# - cruzamentos;
# - limitações de circulação;
# - demais características da malha urbana.

# Portanto:

# Distância estimada pela malha =
# Distância Haversine × Fator de Malha Urbana

# ---------------------------------------------------------
# VELOCIDADE_MEDIA_KMH
# ---------------------------------------------------------

# Representa a velocidade média adotada para estimar o
# tempo de deslocamento em ambiente urbano.

# O parâmetro funciona como uma aproximação das condições
# de circulação urbana, considerando que os deslocamentos
# ocorrem predominantemente em vias sujeitas a limites
# e condições de circulação urbana.

# ---------------------------------------------------------
# FATOR_TRANSITO
# ---------------------------------------------------------

# Representa um acréscimo no tempo estimado para considerar
# o efeito das condições de circulação sobre o deslocamento,
# incluindo:

# - retenções;
# - paradas;
# - cruzamentos;
# - variações no fluxo de veículos.

# ---------------------------------------------------------
# CALIBRAÇÃO / VALIDAÇÃO
# ---------------------------------------------------------

# Os parâmetros foram ajustados e posteriormente validados
# por meio da comparação entre os tempos e distâncias
# estimados pelo modelo e trajetos de referência obtidos
# no Google Maps para diferentes deslocamentos urbanos.

# O objetivo dessa comparação foi avaliar se o modelo produz
# estimativas compatíveis com condições reais de deslocamento.

# A comparação não tem como objetivo reproduzir exatamente
# o algoritmo utilizado pelo Google Maps, mas avaliar a
# adequação dos parâmetros adotados para o contexto do estudo.

# Fator Malha Urbana estava sendo usado como 1.60

# Afeta mais a distancia
FATOR_MALHA_URBANA = 1.43

VELOCIDADE_MEDIA_KMH = 30

# Afeta mais o tempo até o local
FATOR_TRANSITO = 0.83


#%%
# =========================================================
# FATORES DE CENÁRIO (COLUNAS ADICIONAIS)
# =========================================================

# Fatores utilizados para gerar colunas adicionais no mesmo
# arquivo de saída. Cada fator é multiplicado pela coluna
# DISTANCIA_KM da matriz já calculada.

# Isso NÃO altera o cálculo original da matriz (Haversine,
# malha urbana, tempo, trânsito etc.) — é aplicado apenas
# na etapa final de exportação, sobre o resultado já pronto.

# Basta adicionar ou remover fatores aqui para controlar
# quantas colunas adicionais serão geradas.
FATOR_025 = 0.25
FATOR_05 = 0.5
FATOR_1_25 = 1.25
FATOR_1_5 = 1.5
FATOR_1_75 = 1.75
FATOR_2 = 2
FATOR_2_25 = 2.25
FATOR_2_5 = 2.5

# Dicionário usado na criação das colunas adicionais.
# Chave = sufixo que entra no nome da coluna.
# Valor = fator multiplicado na DISTANCIA_KM.
FATORES_EXPORTACAO_ADICIONAL = {
    "Fator_025": FATOR_025,
    "Fator_05": FATOR_05,
    "Fator_1_25": FATOR_1_25,
    "Fator_1_5": FATOR_1_5,
    "Fator_1_75": FATOR_1_75,
    "Fator_2": FATOR_2,
    "Fator_2_25": FATOR_2_25,
    "Fator_2_5": FATOR_2_5,}


#%%
# =========================================================
# FUNÇÃO PRINCIPAL
# =========================================================

# Esta função reúne as etapas utilizadas para estimar uma
# rota entre dois pontos.

# Fluxo do cálculo:

# 1. Calcula a distância em linha reta pelo Haversine;
# 2. Aplica o fator de malha urbana;
# 3. Calcula o tempo com base na velocidade média;
# 4. Aplica o fator de trânsito;
# 5. Converte o tempo de horas para minutos.

# O resultado final retorna apenas os indicadores necessários
# para a construção da matriz de rotas.
def calcular_rota_google_style(
    lat1,
    lon1,
    lat2,
    lon2):

    # -----------------------------------------------------
    # ETAPA 1 — DISTÂNCIA HAVERSINE
    # -----------------------------------------------------
    
    # Calcula a distância em linha reta entre os dois pontos.
    distancia_haversine = haversine(
        lat1,
        lon1,
        lat2,
        lon2)


    # -----------------------------------------------------
    # ETAPA 2 — AJUSTE DA MALHA URBANA
    # -----------------------------------------------------
    
    # Ajusta a distância Haversine para representar,
    # de forma aproximada, o percurso realizado pela
    # malha viária urbana.
    distancia_real = (
        distancia_haversine
        * FATOR_MALHA_URBANA)


    # -----------------------------------------------------
    # ETAPA 3 — ESTIMATIVA DO TEMPO
    # -----------------------------------------------------
    
    # Calcula o tempo necessário para percorrer a distância
    # ajustada considerando a velocidade média adotada.
    tempo_horas = (
        distancia_real
        / VELOCIDADE_MEDIA_KMH)


    # -----------------------------------------------------
    # ETAPA 4 — AJUSTE PELO TRÂNSITO
    # -----------------------------------------------------
    
    # Aplica o fator de trânsito para representar o impacto
    # estimado das condições de circulação urbana sobre o
    # tempo de deslocamento.
    tempo_horas = (
        tempo_horas
        * FATOR_TRANSITO)


    # -----------------------------------------------------
    # ETAPA 5 — CONVERSÃO PARA MINUTOS
    # -----------------------------------------------------
    
    # Converte o tempo estimado de horas para minutos.
    tempo_min = (
        tempo_horas
        * 60)


    # -----------------------------------------------------
    # RETORNO
    # -----------------------------------------------------
    
    # Retorna distância e tempo arredondados para duas
    # casas decimais.
    return {
        "distancia_km": round(distancia_real, 2),
        "tempo_min": round(tempo_min, 2)}


#%%
# =========================================================
# LISTA RESULTADOS
# =========================================================

# Lista que armazenará todas as relações de deslocamento
# calculadas pelo modelo.

# Cada registro posteriormente dará origem a uma linha
# da matriz final de rotas.
resultados_rotas = []


#%%
# =========================================================
# BLOCO 1
# PARTIDA -> MERCADO
# =========================================================

# Para cada ponto de partida, calcula o deslocamento até
# cada mercado disponível.

# Exemplo:

# Partida 1 → Mercado 1
# Partida 1 → Mercado 2
# Partida 1 → Mercado 3

# Partida 2 → Mercado 1
# Partida 2 → Mercado 2

# Dessa forma, cada possível origem possui uma distância
# e um tempo estimados até todos os mercados.
for _, partida in df_partidas.iterrows():

    for _, mercado in df_mercados.iterrows():

        # Calcula a rota utilizando as coordenadas
        # geográficas do ponto de partida e do mercado.
        resultado = calcular_rota_google_style(

            partida["Latitude"],
            partida["Longitude"],
            mercado["Latitude"],
            mercado["Longitude"])


        # Armazena as informações da relação
        # PARTIDA → MERCADO.
        #
        # Além da distância e do tempo, são armazenadas
        # informações de identificação da origem e destino.
        resultados_rotas.append({
            "TIPO_VIAGEM":
            "PARTIDA_MERCADO",
            "TIPO_ORIGEM":
            "PARTIDA",
            "TIPO_DESTINO":
            "MERCADO",
            "ID_ORIGEM":
            partida["ID Endereço"],
            "NOME_ORIGEM":
            partida["Endereço Completo"],
            "ID_DESTINO":
            mercado["ID Mercado"],
            "NOME_DESTINO":
            mercado["Nome"],
            "DISTANCIA_KM":
            resultado["distancia_km"],
            "TEMPO_MIN":
            resultado["tempo_min"]})


#%%
# =========================================================
# BLOCO 2
# MERCADO -> MERCADO
# =========================================================

# Calcula o deslocamento entre todos os pares possíveis
# de mercados.

# Essa etapa é necessária porque uma solução pode envolver
# mais de um estabelecimento.

# Exemplo:

# Mercado 1 → Mercado 2
# Mercado 1 → Mercado 3
# Mercado 2 → Mercado 1
# Mercado 2 → Mercado 3

# A matriz resultante permite posteriormente calcular
# o custo e o tempo de deslocamento entre mercados
# consecutivos de uma determinada rota.
for _, mercado_origem in df_mercados.iterrows():
    for _, mercado_destino in df_mercados.iterrows():

        # NÃO CALCULAR ROTA USANDO MESMO MERCADO
        # Uma rota de um mercado para ele mesmo não representa
        # um deslocamento válido e, portanto, não é incluída.
        if (mercado_origem["ID Mercado"]
            ==
            mercado_destino["ID Mercado"]):
            continue


        # Calcula a distância e o tempo entre os dois mercados.
        resultado = calcular_rota_google_style(
            mercado_origem["Latitude"],
            mercado_origem["Longitude"],
            mercado_destino["Latitude"],
            mercado_destino["Longitude"])


        # Armazena as informações da relação
        # MERCADO → MERCADO.
        resultados_rotas.append({
            "TIPO_VIAGEM":
            "MERCADO_MERCADO",
            "TIPO_ORIGEM":
            "MERCADO",
            "TIPO_DESTINO":
            "MERCADO",
            "ID_ORIGEM":
            mercado_origem["ID Mercado"],
            "NOME_ORIGEM":
            mercado_origem["Nome"],
            "ID_DESTINO":
            mercado_destino["ID Mercado"],
            "NOME_DESTINO":
            mercado_destino["Nome"],
            "DISTANCIA_KM":
            resultado["distancia_km"],
            "TEMPO_MIN":
            resultado["tempo_min"]})


#%%
# =========================================================
# DATAFRAME FINAL
# =========================================================

# Converte a lista de resultados em um DataFrame.

# Cada linha representa um deslocamento possível entre
# dois pontos considerados no modelo.
df_rotas = pd.DataFrame(
    resultados_rotas
)


#%%
# =========================================================
# ID DA ROTA
# =========================================================

# Reinicia o índice do DataFrame para garantir uma sequência
# contínua de registros.
df_rotas = df_rotas.reset_index(drop=True)


# Cria um identificador único para cada relação de rota.

# O ID começa em 1 para facilitar a identificação dos
# registros na base exportada.
df_rotas["ID_ROTA"] = (
    df_rotas.index + 1)


#%%
# =========================================================
# ORGANIZAÇÃO COLUNAS
# =========================================================

# Define a ordem final das colunas da matriz.

# A estrutura mantém:

# - identificação da rota;
# - tipo de viagem;
# - informações da origem;
# - informações do destino;
# - distância estimada;
# - tempo estimado.

# Essa padronização facilita a utilização da matriz
# pelos scripts posteriores.
df_rotas = df_rotas[
    [
        "ID_ROTA",
        "TIPO_VIAGEM",
        "TIPO_ORIGEM",
        "ID_ORIGEM",
        "NOME_ORIGEM",
        "TIPO_DESTINO",
        "ID_DESTINO",
        "NOME_DESTINO",
        "DISTANCIA_KM",
        "TEMPO_MIN"]]


#%%
# =========================================================
# CUSTOS LOGÍSTICOS POR CENÁRIO
# =========================================================

# Cria a coluna-base de custo logístico usando fator 1.
# As colunas originais permanecem inalteradas.
df_rotas["Custo Logistico"] = df_rotas["DISTANCIA_KM"]

# Cria uma coluna para cada fator configurado, sem alterar
# DISTANCIA_KM nem gerar arquivos adicionais.
for nome_fator, valor_fator in FATORES_EXPORTACAO_ADICIONAL.items():
    df_rotas[f"Custo Logistico - {nome_fator}"] = (
        df_rotas["DISTANCIA_KM"] * valor_fator)


#%%
# =========================================================
# EXPORTAÇÃO
# =========================================================

# Exporta a única matriz com as colunas originais e todos
# os cenários de custo logístico.
# index=False evita que o índice interno do DataFrame seja
# criado como uma coluna adicional.
df_rotas.to_excel(output, index=False)


#%%
# =========================================================
# FINAL
# =========================================================

# Exibe uma mensagem de conclusão do processamento
# e informa o local do arquivo gerado.
print("\n")
print("=" * 100)
print("PROCESSAMENTO FINALIZADO")
print("=" * 100)

print(f"Arquivo salvo em:\n{output}")
