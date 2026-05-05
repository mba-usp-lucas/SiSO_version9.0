# 📋 Parking Lot — Dashboard Sell-out Alcon

> **Status:** Especificação técnica para projeto futuro
> **Base:** Reaproveitamento da estrutura do Dashboard Sell-in v5
> **Aproveitamento estimado:** ~80% do código

---

## 🎯 Objetivo

Dashboard de análise de Sell-out com a mesma identidade visual e estrutura do Sell-in, adaptado para o volume e granularidade do dado.

## 📊 Características do dado de Sell-out

| Atributo | Sell-in (atual) | Sell-out (parking lot) |
|---|---|---|
| Volume | ~10-50k linhas | **~800k linhas** |
| Estrutura | Long (1 mês = 1 linha) | **Wide (12 meses em colunas)** |
| Granularidade | Cliente × Produto | **PDV × Cliente × Produto** |
| Período | Variável | **24 meses fixos** |
| Métricas | Unid + BRL + USD | A confirmar |

## 🗓️ Estrutura de agregação (resolve o volume)

Em vez de 24 meses detalhados, agregar em **4 buckets**:

| Bucket | Definição | Uso |
|---|---|---|
| **MAT Atual** | Últimos 12 meses (rolling) | Volume principal de análise |
| **MAT Anterior** | 12 meses anteriores ao MAT Atual | Comparação YoY (base) |
| **YTD Atual** | Janeiro até mês atual do ano corrente | Análise de progresso anual |
| **YTD Anterior** | Janeiro até mesmo mês do ano anterior | Comparação YTD |

**Resultado de volume estimado:**
- 800k linhas brutas → ~30-50k linhas após agregação por (PDV × Cliente × Produto × Bucket)
- JSON embutido: 3-5MB
- Performance equivalente ao Sell-in atual

## 🎯 Comparação principal

**YoY:** MAT Atual vs MAT Anterior

## 🔄 Pipeline de transformação no Python

O `gerar_dashboard_sellout.py` deve fazer:

```
1. Ler xlsx (formato wide com 12+ colunas de mês)
2. Pivot UNPIVOT (pd.melt) → formato long
3. Identificar mês mais recente disponível
4. Calcular janelas dos 4 buckets:
   - MAT_ATUAL = últimos 12 meses
   - MAT_ANT = 12 meses anteriores ao MAT_ATUAL
   - YTD_ATUAL = Jan até mês mais recente
   - YTD_ANT = Jan até mesmo mês do ano anterior
5. Marcar cada linha com seu(s) bucket(s)
   (uma mesma linha pode entrar em mais de um bucket)
6. Agregar por (PDV × Cliente × Produto × Bucket × Demais dimensões)
7. Injetar no template HTML
```

## 🧩 Aproveitamento do código Sell-in

### ✅ Aproveita 100%
- Visual completo (cores Alcon, layout, fundo azul)
- Logo Alcon no header
- Filtro multi-seleção de Fonte e Franquia
- Filtros searchable de Cliente e Produto
- Toggle de métrica global (Unid/BRL/USD)
- Insights automáticos com lógica proporcional ao mercado
- Benchmark Top 3 do mesmo tipo
- Top 15 Clientes / Produtos
- Maiores crescimentos / quedas
- Exportação Excel
- Exportação PowerPoint (7 slides)
- Estrutura modular Python + Template HTML

### 🔧 Adaptar
- **Filtro novo: PDV** (replica lógica do filtro de Cliente)
- **Filtro de período:** muda de "meses" para "buckets MAT/YTD"
- **Gráfico de evolução:** vira **gráfico de barras comparativas** (4 buckets) em vez de linhas mensais
- **MoM:** **remover** (não faz sentido sem detalhe mensal)
- **YTD acumulado mensal:** **remover** (mesmo motivo)
- **Adicionar:** Drill-down por PDV quando cliente é selecionado

### 🆕 Funcionalidades novas
- **View hierárquica:** Cliente → expande PDVs do cliente
- **Comparativo MAT vs YTD** lado a lado nos KPIs
- **Tabela PDV × Produto** (matriz) quando cliente está filtrado

## 📐 Estrutura de KPIs proposta

| KPI | Cálculo |
|---|---|
| MAT Atual | Soma do bucket MAT Atual |
| YoY MAT (%) | (MAT Atual - MAT Anterior) / MAT Anterior |
| YTD Atual | Soma do bucket YTD Atual |
| YoY YTD (%) | (YTD Atual - YTD Anterior) / YTD Anterior |

## 🏪 Filtro novo: PDV

- Searchable input (igual Cliente)
- Mostra contagem de volume ao lado do nome
- Quando PDV é selecionado, todos os outros filtros respeitam

## 🎨 Visual

Mantém 100%:
- Cores Alcon (#0079c1, #49176d, #00a0af, #49a942)
- Fundo azul escuro (#0f2744)
- Cards brancos
- Logo no canto superior esquerdo

## 📦 Entregáveis estimados

| Arquivo | Esforço |
|---|---|
| `gerar_dashboard_sellout.py` | ~4h (pivot + agregação MAT/YTD) |
| `dashboard_sellout_template.html` | ~3h (adaptar filtros e gráficos) |
| Testes com dado real (800k linhas) | ~2h |
| Ajustes finos de performance/visual | ~2h |
| **Total estimado** | **~11h** |

## 🚀 Quando atacar o projeto

Pré-requisitos antes de começar:
1. ✅ Dashboard Sell-in v5 estabilizado e em uso
2. ✅ Hospedagem definida (SharePoint / Power Pages)
3. ✅ Confirmação de quais métricas tem (Unid, BRL, USD?)
4. ✅ Confirmação dos nomes das colunas no xlsx de Sell-out
5. ✅ Definição se PDV tem código + nome ou apenas nome

## 📝 Perguntas em aberto para confirmar antes

- [ ] Coluna PDV: nome único ou tem hierarquia (cidade/estado)?
- [ ] Sell-out tem coluna de **TIPO_CLIENTE** também? (pra benchmark funcionar)
- [ ] Quais métricas? Apenas unidades ou também valor?
- [ ] Tem coluna de FRANQUIA? (importante pro mix)
- [ ] Quem é a fonte do dado? (Close-Up, IQVIA, etc)
- [ ] Atualização do xlsx é mensal? Em que dia do mês?

---

**Próximo passo quando retomar:** revisar especificação, validar perguntas em aberto, e adaptar o `gerar_dashboard_v5.py` em uma versão `gerar_dashboard_sellout.py`.
