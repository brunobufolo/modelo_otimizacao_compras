# Modelo de otimização de compras domésticas baseado em preços, distância e custos de deslocamento

Código-fonte do Trabalho de Conclusão de Curso (MBA em Data Science e Analytics – USP/Esalq) de **Bruno Henrique Bufolo Cardoso**, orientado por **Murilo Henrique Tank Fortunato**.

O projeto responde à pergunta: **em quais condições a compra fracionada entre supermercados reduz o custo total, considerando os preços dos produtos e os custos estimados de deslocamento?**

## Visão geral

Para cada lista de compras e ponto de partida, o modelo decide **em quais mercados comprar** e **em qual ordem visitá-los** (Casa → Mercado(s) → Casa), minimizando:

```
Custo total = custo dos produtos + custo de deslocamento (R$ 1,00 por km)
```

Três abordagens são comparadas:

| Abordagem | Descrição |
| --- | --- |
| **Modelo Inicial (V1)** | Heurística gulosa: a cada etapa incorpora o mercado de maior saldo econômico (economia nos produtos − custo do novo trecho). |
| **Top N (Top 2)** | Evolução da V1: explora os N melhores candidatos por etapa, ampliando a busca sem enumerar tudo. |
| **Brute Force** | Enumeração exaustiva de combinações e ordens de visita; usada como referência do ótimo global. |

## Principais resultados

Base: 7 supermercados de Bragança Paulista (SP), 8 pontos de partida, 13 categorias da cesta básica (DIEESE) × 3 marcas = 39 itens, totalizando **104 cenários** (13 listas × 8 pontos de partida).

- Top N coincidiu com o Brute Force em **59 cenários (56,73%)**; o Modelo Inicial em 48 (46,15%).
- Ambas as heurísticas indicaram economia versus compra em mercado único em **81 dos 104 cenários (77,88%)**.
- Diferença de custo total versus Brute Force (fator logístico 1): Top N **0,45%**, Modelo Inicial **0,74%**.
- Quanto maior o custo logístico, menor o número médio de mercados selecionados.

Detalhes completos no texto do TCC.

## Estrutura do repositório

```
├── Main_Code.py                        # Orquestra a execução dos scripts
├── config.py                           # Caminhos centralizados (relativos ao repositório)
├── 00_Geração_Listas.py                # Gera as listas de compras (opcional)
├── 01_Matriz_Rotas.py                  # Matriz de distâncias/tempos (Haversine + fatores)
├── 02_Enumeração_Soluções.py           # Enumera combinações e ordens de visita
├── 03_Seleção_e_Otimização.py          # Modelo Inicial (V1)
├── 03_Seleção_e_Otimização_TOP_N.py    # Modelo Top N
├── 04_Resumo_Resultado_Otimo.py        # Indicadores da V1
├── 04_Resumo_Resultado_Otimo_TOP_N.py  # Indicadores do Top N
├── 05_Validação_Enumeração.py          # Brute Force (referência do ótimo global)
├── Bases Entrada/
│   └── Base Fonte Itens Mercado.xlsx   # Mercados, pontos de partida, produtos, preços e compras
├── Outputs/                            # Arquivos gerados pelos scripts
├── requirements.txt
└── LICENSE
```

## Como executar

Requisitos: Python 3.9 ou superior.

```bash
git clone https://github.com/brunobufolo/modelo_otimizacao_compras.git
cd modelo_otimizacao_compras

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/macOS
source .venv/bin/activate

pip install -r requirements.txt
python Main_Code.py
```

Os arquivos gerados são gravados na pasta `Outputs/`.


## Dados de entrada

O arquivo `Bases Entrada/Base Fonte Itens Mercado.xlsx` possui as abas:

| Aba | Conteúdo |
| --- | --- |
| Base Mercado | 7 supermercados (ID, nome, rede, coordenadas) |
| Pontos de Partida | 8 origens simuladas do consumidor, com região e coordenadas |
| Produtos & Marcas | Catálogo de 39 itens |
| Produtos & Mercados | 273 preços (produto × mercado); coleta em 01/08/2026 |
| Compras | 912 linhas de item, associadas a 104 cenários (`ID Lista`) |

## Parâmetros do modelo

| Parâmetro | Valor | Onde |
| --- | --- | --- |
| Fator de malha urbana | 1,43 (calibrado contra Google Maps) | `01_Matriz_Rotas.py` |
| Velocidade média | 30 km/h | `01_Matriz_Rotas.py` |
| Fator de trânsito | 0,83 | `01_Matriz_Rotas.py` |
| Custo de deslocamento | R$ 1,00 por km | coluna `Custo Logistico` |
| `TOP_K` | 2 | `03_Seleção_e_Otimização_TOP_N.py` |

A sensibilidade ao custo logístico é testada com fatores de 0,25 a 2,5 (colunas `Custo Logistico - Fator_*` da matriz de rotas).

## Reprodutibilidade

- A geração de listas (script 00) usa semente fixa (`random.seed(42)`).
- Todos os caminhos são relativos à raiz do repositório (`config.py`).



## Como citar

> CARDOSO, B. H. B.; FORTUNATO, M. H. T. *Modelo de otimização de compras domésticas baseado em preços, distância e custos de deslocamento*. Trabalho de Conclusão de Curso (MBA em Data Science e Analytics) – USP/Esalq, 2026.


