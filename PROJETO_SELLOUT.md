# 📊 Projeto Sell-Out Alcon — Plano Final

## ✅ Decisões Consolidadas

### Arquitetura
- **Dashboard 1: SELL-IN** (existente, já com Card Resumo) — mantém
- **Dashboard 2: SELL-OUT** (futuro, projeto separado) — para PDV detalhado, heatmap UF

### Card Resumo no Sell-in
- 6 colunas: Brasil + Grupo Raia + Pague Menos + DPSP + Santa Cruz + Profarma
- Linha Sell-in: ✅ implementado
- Linha Sell-out: 🔜 popular com fonte unificada gerencial (M-2)

### Fonte de Dados
- **Sell-in:** xlsx mensal (já existente)
- **Sell-out:** fonte UNIFICADA com 2 meses de atraso
- **Não usar fonte mensal** nos dashboards oficiais (cobre só 50-70%)

### Comportamento M-2
- Quando filtra mês que sell-out não tem, mostra último disponível com aviso
- Ex: filtra Abril → sell-out mostra Março com badge "📅 Últ. atualização: Mar/2026"

## 🔜 Próxima Sessão — O que preciso de você

### 1. Estrutura exata do f_SELLOUT_GERENCIAL.xlsx

Cole na conversa:
```
[cabeçalho com nomes EXATOS das colunas]
[3-5 linhas de exemplo - valores podem ser fictícios]
```

Pontos a confirmar:
- [ ] Nome da coluna ANO
- [ ] Nome da coluna MÊS
- [ ] Nome da coluna CLIENTE/REDE
- [ ] Nome da coluna VENDAS R$ ("VENDAS PPP"?)
- [ ] Nome da coluna UNIDADES
- [ ] Tem coluna PRODUTO?
- [ ] Tem outras colunas (UF, Canal, Franquia)?

### 2. Lista de redes que aparecem em sell-out

Quero conferir se as palavras-chave do Card Resumo capturam todas:
```
Espero ver algo como:
- DROGA RAIA (ou DROGASIL ou GRUPO RAIA?)
- PAGUE MENOS
- DPSP (ou DROGARIAS SAO PAULO?)
- SANTA CRUZ
- PROFARMA
- Outras que aparecem...
```

## 📋 Roadmap em Fases

### ✅ Fase 0 — Concluído
- Dashboard sell-in completo
- Card Resumo no sell-in (estrutura pronta)
- Slide PowerPoint do Resumo

### 🔜 Fase 1 — Integrar Sell-out Gerencial no Card Resumo
**Bloqueado em:** estrutura dos dados sell-out gerencial

Quando retomar:
- [ ] Definir mapeamento de colunas (config no topo do Python)
- [ ] Adaptar `gerar_dashboard_v5.py` para ler 2 fontes
- [ ] Embutir dados sell-out gerencial no JSON
- [ ] Atualizar `renderResumo()` para popular linha sell-out
- [ ] Lógica "último mês disponível com aviso"
- [ ] Atualizar slide PPT

### 🔜 Fase 2 — Card Saúde do Canal
- Sell-in vs sell-out por rede
- Identificar estoque empurrado / ruptura / saudável
- Semáforo visual + insight automático

### 🔜 Fase 3 — Dashboard Sell-out Standalone
- Pré-agregação Python (1M+ linhas)
- Filtros: Período, Produto, Canal, Região, UF, Rede
- Top 50 PDVs (modos: geral, por rede, por UF)
- Heatmap UF
- Diagnóstico cliente×produto adaptado

### 🔜 Fase 4 — Apresentação ao Gestor
- Pitch + demo ao vivo
- Pedido específico (tempo + acessos)

## 💡 Pitch Resumido

**Posicionamento:**
> "Plataforma analítica em camadas:
> - SELL-IN (ações Alcon): completo, validado
> - SELL-OUT GERENCIAL (canal): integrado ao sell-in via Card Resumo + Saúde
> - SELL-OUT PDV (drill-down): roadmap próximas semanas"

## ⏳ Status Atual

- [x] Sell-in completo
- [x] Card Resumo (estrutura) implementado
- [x] Slide PPT do Card Resumo
- [x] Decisão de fonte: unificada (M-2) para dashboards oficiais
- [x] Decisão arquitetural: 2 dashboards separados
- [ ] **Aguardando: estrutura exata do f_SELLOUT_GERENCIAL.xlsx**
- [ ] **Aguardando: lista de redes que aparecem nesse arquivo**
- [ ] Implementação Fase 1
- [ ] Dashboard sell-out PDV (Fase 3)

---

**Última atualização:** Aguardando você conferir o arquivo sell-out gerencial e compartilhar:
1. Cabeçalho + linhas de exemplo
2. Lista das redes que aparecem
