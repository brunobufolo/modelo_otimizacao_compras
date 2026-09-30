# =========================================================
# CONFIGURAÇÃO CENTRAL DE CAMINHOS
# =========================================================
#
# Todos os scripts (00 a 05 e Main_Code) importam os caminhos
# daqui. Assim, o projeto funciona em qualquer computador sem
# editar cada arquivo: basta clonar o repositório e executar.
#
# Estrutura esperada:
#
# TCC-otimizacao-compras/
# ├── config.py
# ├── Main_Code.py
# ├── 00_..., 01_..., ..., 05_...
# ├── Bases Entrada/   -> arquivos de ENTRADA (versionados)
# └── Outputs/         -> arquivos GERADOS pelos scripts

from pathlib import Path

# Pasta onde este arquivo está (raiz do projeto).
RAIZ = Path(__file__).resolve().parent

PASTA_ENTRADA = RAIZ / "Bases Entrada"
PASTA_SAIDA = RAIZ / "Outputs"

# Garante que a pasta de saída exista (o Git não versiona pastas vazias).
PASTA_SAIDA.mkdir(exist_ok=True)


# ---------------------------------------------------------
# ENTRADA
# ---------------------------------------------------------
BASE_FONTE = PASTA_ENTRADA / "Base Fonte Itens Mercado.xlsx"


# ---------------------------------------------------------
# SAÍDAS
# ---------------------------------------------------------

# Script 00
CENARIOS_LISTAS = PASTA_SAIDA / "Cenarios_Listas_Compras.xlsx"

# Script 01
MATRIZ_ROTAS = PASTA_SAIDA / "Matriz Rotas.xlsx"

# Script 02
ENUMERACAO = PASTA_SAIDA / "Enumeração_Soluções.xlsx"

# Script 03 (Modelo Inicial)
ESCOLHA_MERCADOS = PASTA_SAIDA / "Escolha_Mercados.xlsx"
RESULTADO_OTIMO = PASTA_SAIDA / "Resultado_Otimo_Listas.xlsx"

# Script 03 (Top N)
ESCOLHA_MERCADOS_TOPN = PASTA_SAIDA / "Escolha_Mercados_TOP_N.xlsx"
RESULTADO_OTIMO_TOPN = PASTA_SAIDA / "Resultado_Otimo_Listas_TOP_N.xlsx"

# Script 04
INDICADORES = PASTA_SAIDA / "Indicadores_Resumo.xlsx"
INDICADORES_TOPN = PASTA_SAIDA / "Indicadores_Resumo_TOP_N.xlsx"

# Script 05 (Brute Force / validação)
BRUTE_FORCE = PASTA_SAIDA / "Brute_Force_Solucoes.xlsx"
