"""
Alcon Sales Insights - Gerador de HTML standalone (v5.4)
Dashboard analitico com Sell-in + Sell-out gerencial
============================================================
v5.4: Adiciona leitura de sell-out gerencial (formato wide com meses em colunas)

COMO USAR
---------
1. pip install pandas openpyxl
2. Ajustar PATHs abaixo
3. Ter 'dashboard_template_v5.html' na mesma pasta
4. python gerar_dashboard_v5.py

NOTA: Sell-out e OPCIONAL. Se o arquivo nao existir, o card resumo mostra "—".
"""

from pathlib import Path
import pandas as pd
import json
import re
from datetime import datetime

# =========================================================
# CONFIGURACAO - AJUSTE AQUI
# =========================================================
PATH_XLSX = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\f_SELLIN.xlsx"
PATH_TARGETS = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\Targets.xlsx"
PATH_TARGETS_FIN = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\Targets_Financeiros.xlsx"
PATH_SELLOUT = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\f_SELLOUT_GERENCIAL.xlsx"
PATH_DEPARA = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\DePara_Produtos.xlsx"
PATH_OUTPUT = r"dashboard_alcon_sales_insights.html"
PATH_TEMPLATE = r"dashboard_template_v5.html"
PATH_NAO_MAPEADOS = r"produtos_nao_mapeados.csv"

COLS = {
    "ANO": "ANO",
    "MES_NUM": "MES_NUM",
    "CLIENTE": "GRUPO_CLIENTE_FINAL",
    "TIPO_CLIENTE": "TIPO_CLIENTE_FINAL",
    "FRANQUIA": "FRANQUIA",
    "PRODUTO": "PRODUTO",
    "FONTE": "FONTE",
    "VALOR_UNID": "Vendas_Unid",
    "VALOR_BRL": "Vendas_BRL",
    "VALOR_USD": "Vendas_USD",
    "MOEDA": None,
}

# =========================================================
# CONFIGURACAO SELL-OUT - AJUSTE OS NOMES SE FOREM DIFERENTES
# =========================================================
COLS_SELLOUT = {
    "GRUPO_PAINEL": "GRUPO_PAINEL",
    "FRANQUIA": "FRANQUIA",
    "TIPO_CLIENTE": "TIPO_CLIENTE",
    "PRODUTO": "PROD_DESC",
    "MEDIDA": "MEDIDA",
    "CHAN_DESC": "CHAN_DESC",  # canal: Farmacia / Hospitalar / Outros / Transferencias
    "UF": "UF",  # estado: SP, RJ, MG, etc.
}
MEDIDA_REAIS = "Reais_PPP"
MEDIDA_UNID = "Unidades"

# Filtro de canal - aplicado SEMPRE no agregado de sell-out
# Mantem apenas linhas onde CHAN_DESC contem 'farmacia' (case insensitive)
#
# OBSERVACAO IMPORTANTE para Distribuidores (Profarma, Santa Cruz):
#   - Por padrao, sell-out so tras canal Farmacia
#   - Distribuidores vendem para outros canais (Hospitalar, Outros)
#   - Para analise especifica de distribuidor com TODOS os canais:
#     usuario filtra na mao depois (no dashboard ou direto na base)
#
# Decisao: priorizamos performance (~60k linhas) vs flexibilidade (~207k linhas)
FILTRO_CANAL_PADRAO = "farmacia"


def ler_depara(path):
    """Le tabela De-Para Produtos.

    Espera colunas:
    - PRODUTO_SELLIN: nome canonico do produto no sell-in (pode ter multiplos sell-out apontando para um sell-in)
    - PRODUTO_SELLOUT: nome do produto no sell-out (pode ter multiplos separados por ;)
    - INCLUIR_DASHBOARD: 'SIM' ou 'NAO' (default: SIM)

    Retorna:
    - dict {produto_sellout_norm -> produto_sellin_canonico}
    - set de produtos a IGNORAR (INCLUIR_DASHBOARD = NAO)
    """
    if not Path(path).exists():
        print(f"[AVISO] DePara nao encontrado em {path}")
        print(f"        Sera usado match aproximado (resultado pode ter gaps).")
        print(f"        Para evitar isso, crie um arquivo DePara_Produtos.xlsx com colunas:")
        print(f"        PRODUTO_SELLIN | PRODUTO_SELLOUT | INCLUIR_DASHBOARD")
        return {}, set()

    print(f"[OK] Lendo DePara: {path}")
    df = pd.read_excel(path)

    cols_esperadas = ["PRODUTO_SELLIN", "PRODUTO_SELLOUT", "INCLUIR_DASHBOARD"]
    for c in cols_esperadas[:2]:  # primeiras 2 sao obrigatorias
        if c not in df.columns:
            print(f"[ERRO] Coluna obrigatoria '{c}' ausente no DePara. Coluna sera ignorada.")
            return {}, set()

    if "INCLUIR_DASHBOARD" not in df.columns:
        df["INCLUIR_DASHBOARD"] = "SIM"

    # Normalizacao basica do nome (apenas trim + upper) - sem remover sufixos!
    def norm(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    mapeamento = {}
    ignorar = set()
    sellins_unicos = set()
    duplicatas = []  # PRODUTO_SELLOUT que apareceu mais de uma vez

    for _, row in df.iterrows():
        sellin = norm(row["PRODUTO_SELLIN"])
        sellout = norm(row["PRODUTO_SELLOUT"])
        incluir = norm(row.get("INCLUIR_DASHBOARD", "SIM"))

        if not sellout:
            continue

        # Formato: 1 par por linha (PRODUTO_SELLIN | PRODUTO_SELLOUT)
        # DePara puramente 1:1 - cada PRODUTO_SELLOUT mapeia para 1 PRODUTO_SELLIN
        if incluir == "NAO":
            ignorar.add(sellout)
            continue

        # Validar duplicatas (mesmo PRODUTO_SELLOUT em multiplas linhas)
        if sellout in mapeamento:
            sellin_anterior = mapeamento[sellout]
            duplicatas.append({
                "PRODUTO_SELLOUT": sellout,
                "MAPEAMENTO_USADO": sellin_anterior,
                "MAPEAMENTO_IGNORADO": sellin or "(vazio)",
            })
            continue  # mantem o primeiro mapeamento, ignora duplicatas

        if sellin:
            mapeamento[sellout] = sellin
            sellins_unicos.add(sellin)
        else:
            # sellout sem sellin -> incluir mas nao mapear (mantem nome SO)
            mapeamento[sellout] = sellout

    print(f"     Produtos sell-in unicos: {len(sellins_unicos)}")
    print(f"     Pares sell-in <-> sell-out: {len(mapeamento)}")
    print(f"     Produtos a ignorar: {len(ignorar)}")

    # Avisar duplicatas + gerar CSV
    if duplicatas:
        print()
        print(f"     [AVISO] {len(duplicatas)} duplicatas detectadas em PRODUTO_SELLOUT")
        print(f"     (mantido o primeiro mapeamento, demais ignorados)")
        print(f"     Top 5:")
        for d in duplicatas[:5]:
            print(f"       - '{d['PRODUTO_SELLOUT']}' usado: '{d['MAPEAMENTO_USADO']}' | ignorado: '{d['MAPEAMENTO_IGNORADO']}'")

        # Salvar CSV
        try:
            csv_path = str(Path(path).parent / "DePara_duplicatas.csv")
            pd.DataFrame(duplicatas).to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"     CSV de duplicatas: {csv_path}")
        except Exception as e:
            print(f"     [WARN] Nao foi possivel salvar CSV: {e}")

    return mapeamento, ignorar


def aplicar_depara_e_relatorio(df_sellout, mapeamento, ignorar, path_csv_nao_mapeados):
    """Aplica de-para no DataFrame de sell-out.

    1. Filtra produtos da lista de ignorar
    2. Adiciona coluna PRODUTO_CANONICO (nome do sell-in para cruzamento)
    3. Salva produtos sem mapeamento em CSV
    """
    if df_sellout is None or len(df_sellout) == 0:
        return df_sellout

    if not mapeamento and not ignorar:
        # Sem De-Para configurado, mantem nome original como canonico
        df_sellout["PRODUTO_CANONICO"] = df_sellout["PRODUTO"]
        return df_sellout

    def normalizar(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    df_sellout["_PROD_NORM"] = df_sellout["PRODUTO"].apply(normalizar)

    # 1. Filtrar produtos a ignorar
    antes = len(df_sellout)
    df_sellout = df_sellout[~df_sellout["_PROD_NORM"].isin(ignorar)].copy()
    depois = len(df_sellout)
    if antes != depois:
        print(f"     Linhas filtradas (INCLUIR_DASHBOARD=NAO): {antes - depois:,}")

    # 2. Aplicar mapeamento
    df_sellout["PRODUTO_CANONICO"] = df_sellout["_PROD_NORM"].map(mapeamento).fillna(df_sellout["PRODUTO"])

    # 3. Identificar nao-mapeados
    nao_mapeados_mask = ~df_sellout["_PROD_NORM"].isin(mapeamento.keys()) & ~df_sellout["_PROD_NORM"].isin(ignorar)
    nao_mapeados = df_sellout[nao_mapeados_mask]

    if len(nao_mapeados) > 0:
        # Agrupar por produto e somar volume
        agg = nao_mapeados.groupby("PRODUTO").agg(
            BRL=("BRL", "sum"),
            UNID=("UNID", "sum"),
            LINHAS=("PRODUTO", "size")
        ).reset_index().sort_values("UNID", ascending=False)

        agg["PRODUTO_SELLIN_SUGESTAO"] = ""
        agg["INCLUIR_DASHBOARD"] = "SIM"

        # Reordenar colunas pra facilitar copiar para o DePara
        agg = agg[["PRODUTO", "PRODUTO_SELLIN_SUGESTAO", "INCLUIR_DASHBOARD", "BRL", "UNID", "LINHAS"]]
        agg = agg.rename(columns={"PRODUTO": "PRODUTO_SELLOUT"})

        agg.to_csv(path_csv_nao_mapeados, index=False, encoding="utf-8-sig")

        print()
        print("=" * 60)
        print(f"  PRODUTOS SELL-OUT NAO MAPEADOS NO DEPARA")
        print("=" * 60)
        print(f"  Total: {len(agg)} produtos | {len(nao_mapeados):,} linhas")
        print(f"  CSV gerado: {path_csv_nao_mapeados}")
        print()
        print(f"  Top 10 por volume (UNID):")
        for _, row in agg.head(10).iterrows():
            print(f"    {row['UNID']:>10,.0f} unid - {row['PRODUTO_SELLOUT']}")
        print()
        print(f"  ACAO RECOMENDADA: Abra o CSV e o DePara_Produtos.xlsx,")
        print(f"  classifique cada produto e rode o script novamente.")
        print(f"  Esses produtos NAO aparecerao no dashboard ate serem mapeados.")
        print("=" * 60)

        # Filtrar nao-mapeados (decisao A: nao exibe no dashboard)
        df_sellout = df_sellout[~nao_mapeados_mask].copy()
        print(f"     Linhas removidas (nao mapeadas): {len(nao_mapeados):,}")

    df_sellout = df_sellout.drop(columns=["_PROD_NORM"])
    print(f"     Linhas finais sell-out: {len(df_sellout):,}")
    return df_sellout


def ler_sellout_gerencial(path):
    """Le arquivo wide e converte para long agregado.

    IMPORTANTE: aplica filtro CHAN_DESC contendo 'farmacia' (default fixo).
    Hospitalar, Outros canais e Transferencias sao excluidos.
    """
    if not Path(path).exists():
        print(f"[INFO] Sell-out nao encontrado em {path} - sera ignorado")
        return [], None

    print(f"[OK] Lendo sell-out: {path}")
    df = pd.read_excel(path)
    print(f"     Linhas raw: {len(df)}")

    cols_meses = [c for c in df.columns if re.match(r"^\d{4}_\d{2}_\d{2}$", str(c))]
    print(f"     Colunas de mes encontradas: {len(cols_meses)}")

    if len(cols_meses) == 0:
        print("[ERRO] Nenhuma coluna no formato YYYY_MM_DD. Sell-out ignorado.")
        return [], None

    for k, col_name in COLS_SELLOUT.items():
        if col_name not in df.columns:
            # Colunas opcionais - apenas avisa
            if k in ("CHAN_DESC", "UF"):
                print(f"[AVISO] Coluna '{col_name}' nao encontrada - feature relacionada nao funcionara")
                continue
            print(f"[ERRO] Coluna '{col_name}' nao encontrada. Ajuste COLS_SELLOUT no codigo.")
            return [], None

    # NAO filtra mais por canal no Python (decisao: filtro dinamico no JS)
    # Mantemos CHAN_DESC para o JS aplicar a logica condicional
    if FILTRO_CANAL_PADRAO and COLS_SELLOUT["CHAN_DESC"] in df.columns:
        antes = len(df)
        df = df[df[COLS_SELLOUT["CHAN_DESC"]].astype(str).str.contains(FILTRO_CANAL_PADRAO, case=False, na=False)].copy()
        depois = len(df)
        print(f"     Filtro canal padrao '{FILTRO_CANAL_PADRAO}': {antes:,} -> {depois:,} linhas ({100*depois/antes:.1f}%)")

    cols_id = [
        COLS_SELLOUT["GRUPO_PAINEL"],
        COLS_SELLOUT["FRANQUIA"],
        COLS_SELLOUT["TIPO_CLIENTE"],
        COLS_SELLOUT["PRODUTO"],
        COLS_SELLOUT["MEDIDA"],
    ]
    # Adiciona CHAN_DESC e UF se existirem (para o JS usar)
    if COLS_SELLOUT["CHAN_DESC"] in df.columns:
        cols_id.append(COLS_SELLOUT["CHAN_DESC"])
    if COLS_SELLOUT["UF"] in df.columns:
        cols_id.append(COLS_SELLOUT["UF"])
    df = df[cols_id + cols_meses].copy()

    df_long = df.melt(
        id_vars=cols_id,
        value_vars=cols_meses,
        var_name="DATA_RAW",
        value_name="VALOR",
    )

    df_long["VALOR"] = pd.to_numeric(df_long["VALOR"], errors="coerce").fillna(0)
    df_long = df_long[df_long["VALOR"] > 0]

    df_long["ANO"] = df_long["DATA_RAW"].str[:4].astype(int)
    df_long["MES"] = df_long["DATA_RAW"].str[5:7].astype(int)
    df_long = df_long.drop(columns=["DATA_RAW"])

    # Index do pivot: inclui CHAN_DESC e UF se existirem
    idx_cols = [
        COLS_SELLOUT["GRUPO_PAINEL"],
        COLS_SELLOUT["FRANQUIA"],
        COLS_SELLOUT["TIPO_CLIENTE"],
        COLS_SELLOUT["PRODUTO"],
    ]
    tem_canal = COLS_SELLOUT["CHAN_DESC"] in df_long.columns
    tem_uf = COLS_SELLOUT["UF"] in df_long.columns
    if tem_canal:
        idx_cols.append(COLS_SELLOUT["CHAN_DESC"])
    if tem_uf:
        idx_cols.append(COLS_SELLOUT["UF"])
    idx_cols += ["ANO", "MES"]

    df_pivot = df_long.pivot_table(
        index=idx_cols,
        columns=COLS_SELLOUT["MEDIDA"],
        values="VALOR",
        aggfunc="sum",
    ).reset_index()

    df_pivot.columns.name = None
    if MEDIDA_REAIS in df_pivot.columns:
        df_pivot = df_pivot.rename(columns={MEDIDA_REAIS: "BRL"})
    else:
        df_pivot["BRL"] = 0
    if MEDIDA_UNID in df_pivot.columns:
        df_pivot = df_pivot.rename(columns={MEDIDA_UNID: "UNID"})
    else:
        df_pivot["UNID"] = 0

    df_pivot["BRL"] = df_pivot["BRL"].fillna(0)
    df_pivot["UNID"] = df_pivot["UNID"].fillna(0)

    rename_map = {
        COLS_SELLOUT["GRUPO_PAINEL"]: "GRUPO_PAINEL",
        COLS_SELLOUT["FRANQUIA"]: "FRANQUIA",
        COLS_SELLOUT["TIPO_CLIENTE"]: "TIPO_CLIENTE",
        COLS_SELLOUT["PRODUTO"]: "PRODUTO",
    }
    if tem_canal:
        rename_map[COLS_SELLOUT["CHAN_DESC"]] = "CHAN_DESC"
    if tem_uf:
        rename_map[COLS_SELLOUT["UF"]] = "UF"
    df_pivot = df_pivot.rename(columns=rename_map)

    # APLICAR DE-PARA DE PRODUTOS
    # - Substitui PRODUTO pelo nome canonico do sell-in (mantendo o original em PRODUTO_ORIGINAL)
    # - Filtra produtos da lista 'ignorar'
    # - Salva nao-mapeados em CSV para revisao
    mapeamento, ignorar = ler_depara(PATH_DEPARA)
    df_pivot["PRODUTO_ORIGINAL"] = df_pivot["PRODUTO"]
    df_pivot = aplicar_depara_e_relatorio(df_pivot, mapeamento, ignorar, PATH_NAO_MAPEADOS)
    if df_pivot is None or len(df_pivot) == 0:
        print("[AVISO] Sell-out vazio apos aplicar DePara")
        return [], None

    # Substituir PRODUTO pelo canonico (que sera usado no dashboard)
    df_pivot["PRODUTO"] = df_pivot["PRODUTO_CANONICO"]
    df_pivot = df_pivot.drop(columns=["PRODUTO_CANONICO"])

    # =================================================================
    # CRIAR 2 ESTRUTURAS SEPARADAS PARA OTIMIZAR PERFORMANCE:
    #
    # 1. PRINCIPAL (sem UF, sem CHAN_DESC): usado por 95% do dashboard
    #    Agrega tudo no nivel cliente x produto x mes (~60k linhas tipico)
    #
    # 2. UF (com UF): usado APENAS pelo card "Diagnostico por UF"
    #    Mantem granularidade por estado (~150k+ linhas)
    #
    # Isso evita inflar todos os cards com dimensoes que so um usa.
    # =================================================================

    # Estrutura PRINCIPAL: agrega UF e CHAN_DESC (vira 1 linha por cliente x produto x mes)
    df_principal = df_pivot.groupby(
        ["GRUPO_PAINEL", "FRANQUIA", "TIPO_CLIENTE", "PRODUTO", "ANO", "MES"],
        dropna=False
    ).agg(BRL=("BRL", "sum"), UNID=("UNID", "sum")).reset_index()

    # Estrutura UF: mantem UF como dimensao (so usado pelo card UF)
    df_uf = None
    if tem_uf:
        df_uf = df_pivot.groupby(
            ["GRUPO_PAINEL", "FRANQUIA", "TIPO_CLIENTE", "PRODUTO", "UF", "ANO", "MES"],
            dropna=False
        ).agg(BRL=("BRL", "sum"), UNID=("UNID", "sum")).reset_index()

    ultimo_mes = df_principal.sort_values(["ANO", "MES"], ascending=False).iloc[0]
    ultimo_periodo = {"ano": int(ultimo_mes["ANO"]), "mes": int(ultimo_mes["MES"])}

    print(f"     Linhas estrutura PRINCIPAL: {len(df_principal):,}")
    if df_uf is not None:
        print(f"     Linhas estrutura UF (separada): {len(df_uf):,}")
    print(f"     Ultimo periodo: {ultimo_periodo['mes']:02d}/{ultimo_periodo['ano']}")
    print(f"     Total BRL: R$ {df_principal['BRL'].sum():,.0f}")
    print(f"     Total Unid: {df_principal['UNID'].sum():,.0f}")

    registros = df_principal.to_dict("records")
    registros_uf = df_uf.to_dict("records") if df_uf is not None else []
    return registros, ultimo_periodo, registros_uf


def ler_sellin(path):
    """Le sell-in"""
    print(f"[OK] Lendo sell-in: {path}")
    df = pd.read_excel(path)
    print(f"     Linhas: {len(df):,}")

    obrig = ["ANO", "MES_NUM"]
    for c in obrig:
        if c not in df.columns:
            raise ValueError(f"Coluna obrigatoria '{c}' ausente em {path}")

    for c in ["Vendas_Unid", "Vendas_BRL", "Vendas_USD"]:
        if c not in df.columns:
            df[c] = 0

    for c, default in [
        ("GRUPO_CLIENTE_FINAL", "—"),
        ("TIPO_CLIENTE_FINAL", "—"),
        ("FRANQUIA", "—"),
        ("PRODUTO", "—"),
        ("FONTE", "—"),
    ]:
        if c not in df.columns:
            df[c] = default

    return df


def ler_targets(path):
    if not Path(path).exists():
        return {}
    df = pd.read_excel(path)
    if "PRODUTO" not in df.columns or "TARGET_PCT" not in df.columns:
        return {}
    if "FLAG" in df.columns:
        df_foco = df[df["FLAG"].astype(str).str.upper() == "FOCO"]
    else:
        df_foco = df
    targets = {}
    for _, row in df_foco.iterrows():
        produto = str(row["PRODUTO"]).strip()
        targets[produto] = float(row["TARGET_PCT"])
    print(f"[OK] {len(targets)} targets FOCO carregados")
    return targets


def ler_targets_fin(path):
    if not Path(path).exists():
        return []
    df = pd.read_excel(path)
    cols_obrig = ["FRANQUIA", "PRODUTO", "ANO", "MES_NUM"]
    if not all(c in df.columns for c in cols_obrig):
        return []
    for c in ["TARGET_BRL", "TARGET_UNID"]:
        if c not in df.columns:
            df[c] = 0
    print(f"[OK] {len(df)} targets financeiros carregados")
    return df.to_dict("records")


def gerar_html():
    df = ler_sellin(PATH_XLSX)
    targets = ler_targets(PATH_TARGETS)
    targets_fin = ler_targets_fin(PATH_TARGETS_FIN)
    sellout, ultimo_sellout, sellout_uf = ler_sellout_gerencial(PATH_SELLOUT)

    depara_ok = Path(PATH_DEPARA).exists()

    # Mapeamento cliente_sellin -> TIPO_CLIENTE_FINAL
    # Usado pelo JS pra identificar distribuidores no sell-out
    # (matching aproximado entre nomes do sell-in e GRUPO_PAINEL do sell-out)
    cliente_tipo_map = {}
    for cli, tipo in df.groupby([COLS["CLIENTE"], COLS["TIPO_CLIENTE"]]).groups.keys():
        if cli and tipo:
            # Normaliza nome (UPPER + trim) pra facilitar lookup no JS
            nome_norm = str(cli).strip().upper()
            tipo_norm = str(tipo).strip().upper()
            cliente_tipo_map[nome_norm] = tipo_norm
    print(f"\nMapeamento cliente->tipo: {len(cliente_tipo_map)} clientes mapeados")

    template = Path(PATH_TEMPLATE).read_text(encoding="utf-8")

    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
    dados_json = json.dumps(df.to_dict("records"), ensure_ascii=False, default=str)
    cols_json = json.dumps(COLS, ensure_ascii=False)
    targets_json = json.dumps(targets, ensure_ascii=False)
    targets_fin_json = json.dumps(targets_fin, ensure_ascii=False, default=str)
    sellout_json = json.dumps(sellout, ensure_ascii=False, default=str)
    sellout_uf_json = json.dumps(sellout_uf, ensure_ascii=False, default=str) if sellout_uf else "[]"
    ultimo_sellout_json = json.dumps(ultimo_sellout, ensure_ascii=False) if ultimo_sellout else "null"
    cliente_tipo_json = json.dumps(cliente_tipo_map, ensure_ascii=False)

    print(f"\nJSON sell-in: {len(dados_json)/1024:.1f} KB")
    print(f"JSON sell-out: {len(sellout_json)/1024:.1f} KB")
    print(f"JSON sell-out UF: {len(sellout_uf_json)/1024:.1f} KB")

    padrao_cols = re.compile(r"const\s+COLS\s*=\s*\{[^}]+\};", re.DOTALL)
    novo_template, n_cols = padrao_cols.subn(
        "const COLS = window.__ALCON_COLS_EMBED__;",
        template, count=1
    )
    if n_cols == 0:
        raise ValueError("Nao achei o bloco 'const COLS = {...}' no template.")
    template = novo_template

    marcador = "const COLS = window.__ALCON_COLS_EMBED__;"
    script_injecao = (
        f"window.__ALCON_DATA_EMBED__ = {dados_json};\n"
        f"window.__ALCON_COLS_EMBED__ = {cols_json};\n"
        f"window.__ALCON_TARGETS_EMBED__ = {targets_json};\n"
        f"window.__ALCON_TARGETS_FIN_EMBED__ = {targets_fin_json};\n"
        f"window.__ALCON_SELLOUT_EMBED__ = {sellout_json};\n"
        f"window.__ALCON_SELLOUT_UF_EMBED__ = {sellout_uf_json};\n"
        f"window.__ALCON_SELLOUT_ULTIMO__ = {ultimo_sellout_json};\n"
        f"window.__ALCON_CLIENTE_TIPO_MAP__ = {cliente_tipo_json};\n"
        f"window.__ALCON_DEPARA_OK__ = {str(depara_ok).lower()};\n"
        f"window.__ALCON_META_EMBED__ = {{ registros: {len(df)}, sellout_registros: {len(sellout)}, depara_ok: {str(depara_ok).lower()}, timestamp: \"{timestamp}\" }};\n"
        f"{marcador}"
    )
    template = template.replace(marcador, script_injecao, 1)

    padrao_init = re.compile(r"rawData\s*=\s*gerarDadosExemplo\(\)\s*;", re.DOTALL)
    init_novo = (
        "rawData = window.__ALCON_DATA_EMBED__.map(function(r){"
        "return Object.assign({}, r, {"
        "[COLS.ANO]: +r[COLS.ANO],"
        "[COLS.MES_NUM]: +r[COLS.MES_NUM],"
        "[COLS.VALOR_UNID]: +r[COLS.VALOR_UNID] || 0,"
        "[COLS.VALOR_BRL]: +r[COLS.VALOR_BRL] || 0,"
        "[COLS.VALOR_USD]: +r[COLS.VALOR_USD] || 0,"
        "[COLS.TIPO_CLIENTE]: r[COLS.TIPO_CLIENTE] || '-'"
        "});});"
        "if(window.__ALCON_TARGETS_EMBED__){mapTargets = window.__ALCON_TARGETS_EMBED__;}"
        "if(window.__ALCON_TARGETS_FIN_EMBED__){targetsFinanceiros = window.__ALCON_TARGETS_FIN_EMBED__;}"
        "if(window.__ALCON_SELLOUT_EMBED__){rawDataSellout = window.__ALCON_SELLOUT_EMBED__.map(function(r){return Object.assign({},r,{ANO:+r.ANO,MES:+r.MES,BRL:+r.BRL||0,UNID:+r.UNID||0});});}"
        "if(window.__ALCON_SELLOUT_UF_EMBED__){window.__rawDataSelloutUF = window.__ALCON_SELLOUT_UF_EMBED__.map(function(r){return Object.assign({},r,{ANO:+r.ANO,MES:+r.MES,BRL:+r.BRL||0,UNID:+r.UNID||0});});}"
        "if(window.__ALCON_SELLOUT_ULTIMO__){selloutUltimoPeriodo = window.__ALCON_SELLOUT_ULTIMO__;}"
        "if(window.__ALCON_CLIENTE_TIPO_MAP__){window.__clienteTipoMap = window.__ALCON_CLIENTE_TIPO_MAP__;}"
        "if(window.__ALCON_DEPARA_OK__ === false && rawDataSellout && rawDataSellout.length > 0){var b=document.getElementById('deparaBanner');if(b)b.style.display='block';}"
    )
    novo_template2, n2 = padrao_init.subn(init_novo, template, count=1)

    if n2 == 0:
        raise ValueError("Nao achei 'rawData = gerarDadosExemplo()'.")
    template = novo_template2

    padrao_func = re.compile(
        r"function\s+gerarDadosExemplo\s*\(\s*\)\s*\{.*?^\}",
        re.DOTALL | re.MULTILINE
    )
    template = padrao_func.sub("// gerarDadosExemplo() removida - dados embutidos via Python", template, count=1)

    status_msg = f"{len(df):,} registros sell-in"
    if sellout:
        status_msg += f" + {len(sellout):,} sell-out"
    if targets:
        status_msg += f" + {len(targets)} targets FOCO"
    status_msg += f" - {timestamp}"

    template = template.replace(
        "<span id=\"statusText\">Carregando...</span>",
        f"<span id=\"statusText\">{status_msg}</span>"
    )

    Path(PATH_OUTPUT).write_text(template, encoding="utf-8")
    print(f"\n[OK] Gerado: {PATH_OUTPUT} ({len(template)/1024/1024:.2f} MB)")
    print(f"=" * 60)
    print(f"Pronto! Arquivo: {PATH_OUTPUT}")
    print(f"=" * 60)


if __name__ == "__main__":
    gerar_html()
