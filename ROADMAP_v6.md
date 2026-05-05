# 🚀 Roadmap v6 — Análise Avançada

> **Contexto:** Após estabilização do v5 (Sell-in + Sell-out gerencial + DePara), próximas evoluções focam em análise estratégica para Diretoria e IC.

---

## 🅰️ Fase 1 (PRIORIDADE) — Decomposição + One-Pager

### Feature 1.1: Decomposição Preço × Volume × Mix

**Justificativa:** Hoje sabemos que cresceu, não sabemos POR QUE. Perguntas como "o crescimento veio de aumento de preço ou venda de mais unidades?" são críticas para decisões estratégicas (precificação vs volume push).

**Cálculo (decomposição clássica de bridge):**

```
ΔReceita Total = Efeito Volume + Efeito Preço + Efeito Mix

Onde:
- Preço médio = Vendas_BRL / Vendas_Unid (calculado dos próprios dados SI)
- Efeito Volume = (Qtd_atual − Qtd_comp) × Preço_comp
- Efeito Preço = (Preço_atual − Preço_comp) × Qtd_atual
- Efeito Mix = ΔReceita_total − Efeito Volume − Efeito Preço
```

**Granularidade:** Produto+Cliente (mais preciso, captura descontos por rede)
- Se filtrar 1 cliente: mostra preço daquele cliente
- Se Brasil: agrega por produto

**Onde aparecerá:**
- Card novo no HTML: "Decomposição de Crescimento"
- Visual: waterfall chart
- Slide PPT dedicado
- Insight automático em prosa

**Exemplo de saída:**
```
SYSTANE ULTRA · 2026 vs 2025
├─ Receita 2025:  R$ 10,0M
├─ Receita 2026:  R$ 12,0M
├─ Δ Total:       +R$ 2,0M (+20%)
│
├─ ▲ Volume:      +R$ 1,00M  (50% do crescimento)
├─ ▲ Preço:       +R$ 0,99M  (49% do crescimento)
└─ ◆ Mix:         +R$ 0,01M  (1% — neutro)
```

---

### Feature 1.2: One-Pager Executivo

**Justificativa:** Diretor não abre dashboard. Precisa de 1 página visual com toda a história em <10 segundos.

**Layout (1 página A4 paisagem):**

```
┌──────────────────────────────────────────────────────────────────┐
│  🏥 ALCON BRASIL · MAT 12 (Mai/25 - Abr/26) · Atualizado: 01/Mai │
├──────────────────────────────────────────────────────────────────┤
│  [GRANDE]                                                        │
│  R$ 480M  ▲ +12%       Target Anual: R$ 520M                    │
│  vs YA                  YTD: 78% atingido                        │
│                         Gap p/ atingir: -R$ 40M                  │
├──────────────────────────────────────────────────────────────────┤
│  📊 STORY                                                        │
│  Crescimento de +12% em MAT 12 puxado por LENTES (+18%) e PHARMA │
│  (+9%). Decomposição: 60% volume, 40% preço.                     │
│  Gap de R$ 40M para target concentrado em DPSP (-R$ 28M, -15%   │
│  vs YA) e Profarma (-R$ 12M, -8%).                              │
├──────────────────────────────────────────────────────────────────┤
│  TOP 3 ALAVANCAS              │  TOP 3 OFENSORES               │
│  ▲ TOTAL 30      +R$ 8M       │  ▼ DPSP          -R$ 28M      │
│  ▲ DAILIES T1    +R$ 5M       │  ▼ Profarma      -R$ 12M      │
│  ▲ SYSTANE ULTRA +R$ 4M       │  ▼ TRAVATAN Z    -R$ 6M       │
├──────────────────────────────────────────────────────────────────┤
│  STATUS TARGETS FOCO     │  SI vs SO                            │
│  ✅ 8 de 12 OK            │  ✅ Saudável: 6 redes                 │
│  ⚠️ 4 atenção             │  ⚠️ Empurrado: 2 (RAIA, PMenos)      │
│                          │  🔴 Ruptura: 1 (Panvel)              │
└──────────────────────────────────────────────────────────────────┘
```

**Geração:**
- Botão dedicado: "📄 One-Pager Executivo"
- Slide PPT único (1 slide ao invés de 13)
- Pode ser impresso direto em A4

---

## 🅱️ Fase 2 — Forecast vs Target

### Feature 2.1: Projeção Sazonalizada

**Justificativa:** Saber se vai bater target — e o quanto falta.

**Método escolhido:** Sazonalização baseada em 24 meses (2 anos completos)

**Cálculo:**
```
Para cada mês m ∈ [Jan..Dez]:
  peso[m] = média(vendas_2024[m], vendas_2025[m]) / total_anual_médio

Projeção[ano_atual] = vendas_YTD + Σ(meses_futuros × peso[m] × ritmo_atual)

ritmo_atual = vendas_YTD_2026 / vendas_YTD_2025
```

**Exemplo:**
```
Hoje: 01/Mai/2026, dados Jan-Abr
Vendas YTD 2026: R$ 12M (4 meses)
Vendas YTD 2025: R$ 11M
Ritmo atual: 12/11 = 1,09 (cresceu 9%)

Para Mai-Dez: aplica peso histórico × ritmo
Projeção Mai-Dez: R$ 24M
Projeção total 2026: R$ 36M

Target: R$ 40M
Gap: -R$ 4M (-10%)
```

### Feature 2.2: Distribuição do Gap

**Lógica:** "Onde recuperar?" — distribui o gap pelas redes com **maior potencial absoluto** (alto SI atual + baixo SO crescimento = gap recuperável).

**Output:**
```
Para fechar gap de R$ 4M, recuperar:
- DPSP:     R$ 1,8M  (atual SO -8% vs Brasil +5% = 13pp underperformance)
- Profarma: R$ 1,2M  (rede grande mas crescendo abaixo do mercado)
- Panvel:   R$ 0,5M  (sell-out negativo, possível ruptura)
Outros:     R$ 0,5M

Confiança: MÉDIA (4 meses de YTD, sazonalidade incerta para Q4)
```

### Feature 2.3: Slide "Action Plan"

Slide PPT com:
- Projeção atual vs target
- Gap em R$ e %
- Top 5 redes/produtos com potencial de recuperação
- Recomendações automáticas em prosa

---

## 📅 Cronograma sugerido

- **Sprint 1 (atual):** Decomposição + One-Pager
- **Sprint 2:** Forecast + Distribuição de gap
- **Backlog:** Drill-down PDV / sell-out por loja (após resolver dados)

---

## ⏸️ Parqueado (resolver primeiro)

Ver `PARKING_LOT_SELLOUT.md` — pendências de qualidade do dado sell-out
antes de avançar com features que dependam dele:
- Cobertura do painel Close-Up (fora de Pharma)
- Drill-down PDV
- Validação de DePara em produção

---
