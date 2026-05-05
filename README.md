# 📊 Alcon · Sales Insights

Dashboard analítico em HTML standalone para análise integrada de **Sell-in** e **Sell-out** da Alcon Brasil. Gera HTML interativo com KPIs, MAT 12 meses, multi-métrica (Unidades/BRL/USD), diagnóstico cliente×produto, análise comparativa Sell-in × Sell-out, e exportação para Excel + PowerPoint.

## 🚀 Como Usar

### 1. Pré-requisitos
```bash
pip install pandas openpyxl
```

### 2. Estrutura de arquivos

Coloque na mesma pasta:
- `dashboard_template_v5.html` — template do dashboard
- `gerar_dashboard_v5.py` — script gerador
- `f_SELLIN.xlsx` — dados de sell-in (obrigatório)
- `Targets.xlsx` — targets de crescimento % (opcional)
- `Targets_Financeiros.xlsx` — targets financeiros mensais (opcional)
- `f_SELLOUT_GERENCIAL.xlsx` — dados de sell-out gerencial (opcional)
- `DePara_Produtos.xlsx` — mapeamento de nomes de produto SI↔SO (recomendado para sell-out preciso)

### 3. Configurar paths

Edite o topo de `gerar_dashboard_v5.py`:
```python
PATH_XLSX = r"C:\caminho\para\f_SELLIN.xlsx"
PATH_TARGETS = r"C:\caminho\para\Targets.xlsx"
PATH_TARGETS_FIN = r"C:\caminho\para\Targets_Financeiros.xlsx"
PATH_SELLOUT = r"C:\caminho\para\f_SELLOUT_GERENCIAL.xlsx"
PATH_DEPARA = r"C:\caminho\para\DePara_Produtos.xlsx"
```

### 4. Executar
```bash
python gerar_dashboard_v5.py
```

Gera `dashboard_alcon_sales_insights.html` — abra no navegador.

## 🔗 De-Para de Produtos (Recomendado)

**Por que existe:** Sell-in e Sell-out muitas vezes usam nomes diferentes pro mesmo produto:
- Sell-in: `SYSTANE ULTRA`
- Sell-out: `SYSTANE UL` ou `SYSTANE ULTRA (ALC)`

**Sem DePara:** O dashboard usa match aproximado (fuzzy) e pode ter pequenos gaps.
**Com DePara:** Cruzamento 100% confiável, números exatos.

### Estrutura DePara_Produtos.xlsx

| PRODUTO_SELLIN | PRODUTO_SELLOUT | INCLUIR_DASHBOARD |
|---|---|---|
| SYSTANE ULTRA | SYSTANE UL; SYSTANE ULTRA (ALC) | SIM |
| PATANOL S | PATANOL S (NVR) | SIM |
| (vazio) | SYSTANE BALANCE | NAO |

**Regras:**
- Múltiplos nomes sell-out separados por `;` apontam pro MESMO sell-in (são somados)
- `INCLUIR_DASHBOARD = NAO` ignora completamente (concorrentes, ruído)
- Produtos não listados aparecem em `produtos_nao_mapeados.csv` (gerado automaticamente) e **não entram no dashboard** até serem mapeados

### Workflow recomendado

1. Roda Python pela 1ª vez sem DePara → console mostra "**X produtos não mapeados**"
2. Abre `produtos_nao_mapeados.csv` (gerado automaticamente)
3. Copia colunas pra `DePara_Produtos.xlsx` e classifica
4. Roda Python novamente → dashboard 100% mapeado

## 📂 Estrutura dos Arquivos de Dados

### f_SELLIN.xlsx (obrigatório)

Colunas mínimas: `ANO`, `MES_NUM`, `Vendas_Unid`

Opcionais: `Vendas_BRL`, `Vendas_USD`, `GRUPO_CLIENTE_FINAL`, `TIPO_CLIENTE_FINAL`, `FRANQUIA`, `PRODUTO`, `FONTE`

### Targets.xlsx (opcional)

| PRODUTO | TARGET_PCT | FLAG |
|---|---|---|
| SYSTANE | 30 | FOCO |

Apenas produtos com `FLAG=FOCO` aparecem.

### f_SELLOUT_GERENCIAL.xlsx (opcional)

Formato **wide** (meses em colunas):

| GRUPO_PAINEL | FRANQUIA | TIPO_CLIENTE | PROD_DESC | MEDIDA | 2024_04_01 | ... | 2026_03_01 |
|---|---|---|---|---|---|---|---|
| GRUPO RAIA DROGASIL | DE&OH | 2.REDE | CLARIL (ALC) | Reais_PPP | 9903.6 | ... | 12658 |

- Coluna `MEDIDA` com valores `Reais_PPP` e `Unidades`
- Colunas de mês formato `YYYY_MM_DD`
- Sell-out tem 2 meses de defasagem (M-2)

## 🎯 Funcionalidades

### Filtros (multi-seleção estilo Power BI)
- Cliente, Produto, Tipo de Cliente, Fonte, Franquia
- Meses (Q1/Q2/Q3/Q4/YTD/**MAT 12m**)

### KPIs principais
- Total atual, vs Comparado, MoM/Variação MAT, YTD/Média MAT

### Cards de análise

#### 📋 Resumo Comparativo por Rede
6 colunas (Brasil + 5 redes principais) com Sell-in e Sell-out lado a lado.

#### 🔄 Análise Comparativa Sell-in × Sell-out
4 blocos coloridos:
- 🚨 **Estoque Empurrado** — SI cresce mais que SO (risco devolução)
- 🆘 **Ruptura/Oportunidade** — SO cresce mais que SI (alta demanda)
- ✅ **Saudável** — gap ≤ 5pp
- 📉 **Ambos em Queda** — demanda real caindo

Análise nos 3 níveis: Produto, Rede, Franquia.

#### 🎯 Status de Targets (Produtos FOCO)
Tabela com sell-in + sell-out lado a lado, gap vs target.

#### 📦 Rankings
- Top 15 Clientes / Produtos
- Maiores Crescimentos / Quedas (com diagnóstico cliente × produto)

#### Gráficos
- Evolução Mensal YoY
- Mix por Franquia (donuts)
- Franquia Comparativo (barras + linha)

### Exportações
- **Excel** com dados pivotados
- **PowerPoint** com 14+ slides estruturados

## 🟠 MAT 12 Meses

Botão laranja no card de meses. Permite analisar 12 meses corridos.

Exemplo: Mar/2026 → MAT Abr/25 a Mar/26 vs MAT Abr/24 a Mar/25.

## ⚠️ Defasagem M-2 (Sell-out)

Sell-out tem 2 meses de atraso em relação ao sell-in. O dashboard:
- Usa último período disponível em sell-out
- Adiciona aviso visual
- Recomenda **MAT 12m** para análises mais robustas

## 🎨 Identidade Visual Alcon

- Cores: #003595 (azul), #00AEEF, #00AE44 (verde), #9B1A2D (vermelho)
- Tipografia: Open Sans (Light + Bold)
- Logo Alcon oficial

## 📁 Configuração de Redes (Card Resumo)

Edite no topo de `dashboard_template_v5.html`:

```javascript
const RESUMO_CLIENTES = [
  { rotulo: 'Brasil',       match: '__TOTAL__' },
  { rotulo: 'Grupo Raia',   match: ['RAIA', 'DROGASIL'] },
  { rotulo: 'Pague Menos',  match: ['PAGUE MENOS'] },
  { rotulo: 'DPSP',         match: ['DPSP', 'DROGARIA SAO PAULO'] },
  { rotulo: 'Santa Cruz',   match: ['SANTA CRUZ'] },
  { rotulo: 'Profarma',     match: ['PROFARMA'] }
];
```

## 🔧 Mapeamento Colunas Sell-out

Em `gerar_dashboard_v5.py`:

```python
COLS_SELLOUT = {
    "GRUPO_PAINEL": "GRUPO_PAINEL",
    "FRANQUIA": "FRANQUIA",
    "TIPO_CLIENTE": "TIPO_CLIENTE",
    "PRODUTO": "PROD_DESC",
    "MEDIDA": "MEDIDA",
}
MEDIDA_REAIS = "Reais_PPP"
MEDIDA_UNID = "Unidades"
```

## 🐛 Troubleshooting

**Card Resumo mostra "—":** confira o caminho do `f_SELLOUT_GERENCIAL.xlsx` e palavras-chave em `RESUMO_CLIENTES`.

**Card Comparativo não aparece:** sell-out não foi carregado. Verifique mensagens do Python.

**Produtos não cruzam:** função `matchByName` faz match flexível. Se algum não bater, me avise.

## 📞 Próximos Passos (Roadmap)

- **Dashboard Sell-out PDV** (projeto separado): análise por PDV individual, geografia, canal
- **Migração Tableau**: camada corporativa para o time IC

---

**Versão:** 5.4 (Sales Insights — sell-in + sell-out gerencial + análise comparativa)
