# =========================================================
# EXPLICATIVO RESUMIDO
# =========================================================

# Este arquivo funciona como script principal de execução
# do projeto.

# Sua finalidade é organizar e executar, em sequência, os
# três scripts responsáveis pelas principais etapas do modelo.

# O objetivo geral do projeto é avaliar estratégias de
# compra doméstica considerando simultaneamente:

# - preços dos produtos;
# - localização dos mercados;
# - custo de deslocamento;
# - seleção dos estabelecimentos economicamente vantajosos;
# - definição da melhor sequência de visita.

# A separação do projeto em arquivos independentes facilita:

# - a compreensão da metodologia;
# - a manutenção do código;
# - a validação de cada etapa separadamente;
# - a rastreabilidade dos resultados gerados.

# =========================================================
# FLUXO GERAL DO PROJETO
# =========================================================

# SCRIPT 01 - MATRIZ DE ROTAS

# Constrói a matriz de deslocamento entre:

# - Pontos de Partida → Mercados
# - Mercados → Mercados

# As distâncias são inicialmente estimadas pelo método
# de Haversine e posteriormente ajustadas pelos parâmetros
# definidos para aproximação do deslocamento urbano.

# SCRIPT 02 - ENUMERAÇÃO DO ESPAÇO DE SOLUÇÕES

# Gera todas as combinações possíveis de mercados e todas
# as respectivas ordens de visita.

# Essa etapa não realiza a seleção econômica da solução.
# Sua principal finalidade é construir um espaço completo
# de soluções que possa ser utilizado posteriormente como
# referência para comparação e validação do modelo.

# SCRIPT 03 - SELEÇÃO E OTIMIZAÇÃO

# Executa o processo principal de decisão do trabalho.

# Inicialmente:

# 1. Calcula o custo de realizar a compra completa em
#    cada mercado individualmente;

# 2. Define M1 como o mercado de menor custo total
#    individual, considerando cesta + deslocamento;

# 3. Avalia os demais mercados de forma incremental;

# 4. Compara a economia potencial obtida nos produtos
#    com o custo adicional de deslocamento;

# 5. Adiciona novos mercados enquanto existir saldo
#    econômico positivo.

# Após a seleção dos mercados:

# 6. São avaliadas todas as ordens possíveis de visita
#    entre os mercados selecionados;

# 7. A sequência de menor custo de deslocamento é definida
#    como a rota final da solução.


# ==========================================
# BIBLIOTECAS PADRÃO
# ==========================================

from pathlib import Path
import sys, runpy

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ)) 

#%%
# =========================================================
# SCRIPT 01
# MATRIZ DE ROTAS
# =========================================================

# Principal DataFrame gerado:
# df_rotas

print("Executando Matriz de Rotas.")

runpy.run_path(
    str(RAIZ / "01_Matriz_Rotas.py"))


#%%
# =========================================================
# SCRIPT 02
# ENUMERAÇÃO DO ESPAÇO DE SOLUÇÕES
# =========================================================

# Principal DataFrame gerado:
# Todas_Rotas

print("Executando Simulação de Todas as Rotas.")

runpy.run_path(
   str(RAIZ /  "02_Enumeração_Soluções.py"))


#%%
# =========================================================
# SCRIPT 03
# SELEÇÃO DE MERCADOS + OTIMIZAÇÃO DA ROTA
# =========================================================

# Entre os resultados gerados estão:

# - mercados selecionados;
# - produtos destinados a cada mercado;
# - rota otimizada;
# - custo dos produtos;
# - custo de deslocamento;
# - custo total;
# - economia em relação à compra em um único mercado.

# Um dos principais DataFrames intermediários gerados é:
# df_escolha_mercados

print("Executando Escolha de Mercados.")

runpy.run_path(
    str(RAIZ / "03_Seleção_e_Otimização.py"))


print("Executando Escolha de Mercados Top N.")

runpy.run_path(
    str(RAIZ / "03_Seleção_e_Otimização_TOP_N.py"))



#%%
# =========================================================
# SCRIPT 04
# Indicadores
# =========================================================

print("Executando Indicadores")

runpy.run_path(
    str(RAIZ / "04_Resumo_Resultado_Otimo.py"))


print("Executando Indicadores Top N")

runpy.run_path(
    str(RAIZ / "04_Resumo_Resultado_Otimo_TOP_N.py"))


#%%
# =========================================================
# SCRIPT 05
# Validação Enumeração
# =========================================================

print("Executando Enumeração Otimizada")

runpy.run_path(
    str(RAIZ / "05_Validação_Enumeração.py"))



#%%
# =========================================================
# SCRIPT 00 - Opcional
# Geração de Listas
# =========================================================


# Este trecho está comentado pois é opcional

# runpy.run_path(
#    "00_Geração_Listas.py")
