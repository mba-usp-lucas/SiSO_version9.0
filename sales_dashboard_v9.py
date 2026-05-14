"""
Alcon Sales Insights - Gerador de HTML standalone (v5.5)
Dashboard analitico com Sell-in + Sell-out gerencial
============================================================
v5.5: - Embutir bibliotecas CDN (Chart.js, SheetJS, PptxGenJS, DataLabels)
      - Substituir Google Fonts por Segoe UI
      - DePara de Clientes (DePara_Clientes.xlsx)
      - Produtos sem DePara NAO sao mais excluidos (mantidos com nome original)
      - HTML 100% self-contained, pronto para deploy IIH

v5.4: Adiciona leitura de sell-out gerencial (formato wide com meses em colunas)

COMO USAR
---------
1. pip install pandas openpyxl
2. Ajustar PATHs abaixo
3. Ter 'dashboard_template_v9.html' na mesma pasta
4. python sales_dashboard_v9.py

NOTA: Sell-out e OPCIONAL. Se o arquivo nao existir, o card resumo mostra "-".
"""

from pathlib import Path
import pandas as pd
import json
import re
from datetime import datetime
import urllib.request
import ssl

# =========================================================
# CONFIGURACAO - AJUSTE AQUI
# =========================================================
PATH_XLSX = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\f_SELLIN.xlsx"
PATH_TARGETS = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\Targets.xlsx"
PATH_TARGETS_FIN = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\Targets_Financeiros.xlsx"
PATH_SELLOUT = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\f_SELLOUT_GERENCIAL.xlsx"
PATH_DEPARA = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\DePara_Produtos.xlsx"
PATH_DEPARA_CLIENTES = r"C:\Users\PEREILU3\OneDrive - Alcon\AdHoc_IC\Projeto\INSIGHTS\DePara_Clientes.xlsx"
PATH_OUTPUT = r"dashboard_alcon_sales_insights.html"
PATH_TEMPLATE = r"dashboard_template_v9.html"
PATH_NAO_MAPEADOS = r"produtos_nao_mapeados.csv"

# =========================================================
# v5.5 - EMBEDDING DE LIBS CDN (zero dependencias externas)
# =========================================================
# Se True, baixa as libs do CDN e embutde inline no HTML final.
# Se False, mantem as tags <script src="...cdn..."> originais.
EMBUTIR_LIBS = True

# Bibliotecas externas para embutir
CDN_LIBS = [
    {
        "name": "Chart.js",
        "version": "4.4.0",
        "url": "https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js",
        "pattern": r'<script\s+src="https://cdn\.jsdelivr\.net/npm/chart\.js@4\.4\.0/dist/chart\.umd\.min\.js"\s*>\s*</script>',
    },
    {
        "name": "SheetJS (xlsx)",
        "version": "0.18.5",
        "url": "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
        "pattern": r'<script\s+src="https://cdn\.jsdelivr\.net/npm/xlsx@0\.18\.5/dist/xlsx\.full\.min\.js"\s*>\s*</script>',
    },
    {
        "name": "PptxGenJS",
        "version": "3.12.0",
        "url": "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js",
        "pattern": r'<script\s+src="https://cdn\.jsdelivr\.net/npm/pptxgenjs@3\.12\.0/dist/pptxgen\.bundle\.js"\s*>\s*</script>',
    },
    {
        "name": "ChartJS DataLabels",
        "version": "2.2.0",
        "url": "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js",
        "pattern": r'<script\s+src="https://cdn\.jsdelivr\.net/npm/chartjs-plugin-datalabels@2\.2\.0/dist/chartjs-plugin-datalabels\.min\.js"\s*>\s*</script>',
    },
]

# Google Fonts patterns para remover
GOOGLE_FONTS_PATTERNS = [
    r'<link\s+rel="preconnect"\s+href="https://fonts\.googleapis\.com"\s*/?>',
    r'<link\s+rel="preconnect"\s+href="https://fonts\.gstatic\.com"\s+crossorigin\s*/?>',
    r'<link\s+href="https://fonts\.googleapis\.com/css2\?family=Open\+Sans[^"]*"\s+rel="stylesheet"\s*/?>',
]

FONT_FALLBACK_CSS = """<style>
  /* === Font Fallback: Google Fonts (Open Sans) -> Segoe UI === */
  body, * {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
  }
</style>"""

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
    "CHAN_DESC": "CHAN_DESC",
    "UF": "UF",
}
MEDIDA_REAIS = "Reais_PPP"
MEDIDA_UNID = "Unidades"

# Filtro de canal - aplicado SEMPRE no agregado de sell-out
# Mantem apenas linhas onde CHAN_DESC contem 'farmacia' (case insensitive)
FILTRO_CANAL_PADRAO = "farmacia"


# =========================================================
# v5.5 - FUNCOES DE EMBEDDING
# =========================================================

def baixar_lib(url, name):
    """Baixa o conteudo de uma biblioteca JS via URL."""
    print(f"  \u2b07\ufe0f  Baixando {name}...")
    ctx = ssl.create_default_context()
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, context=ctx, timeout=60) as response:
            content = response.read().decode("utf-8")
            print(f"     \u2705 {name}: {len(content):,} bytes ({len(content)/1024:.0f} KB)")
            return content
    except Exception as e:
        print(f"     \u274c Erro ao baixar {name}: {e}")
        print(f"        A tag <script src=...> original sera mantida.")
        return None


def embutir_libs_externas(html_content):
    """Embutde todas as bibliotecas CDN inline no HTML.

    Usa re.search() + str.replace() para evitar erros com \\s, \\d no JS.

    1. Baixa cada lib JS do CDN e substitui <script src=...> por <script>conteudo</script>
    2. Remove Google Fonts (preconnect + stylesheet)
    3. Insere CSS fallback com Segoe UI

    Retorna: (html_modificado, libs_embutidas, erros)
    """
    print("\n" + "=" * 60)
    print("  v5.5 - EMBEDDING DE BIBLIOTECAS EXTERNAS")
    print("=" * 60)

    html = html_content
    libs_embutidas = 0
    erros = 0

    # --- Etapa 1: Embutir bibliotecas JS ---
    print(f"\n\U0001f4e6 Etapa 1/3 \u2014 Embutindo {len(CDN_LIBS)} bibliotecas JS...")
    for lib in CDN_LIBS:
        js_content = baixar_lib(lib["url"], lib["name"])
        if js_content:
            inline_tag = f'<script>\n/* === {lib["name"]} v{lib["version"]} (embedded by v5.5) === */\n{js_content}\n</script>'
            # FIX: usa re.search + str.replace para evitar erro \\s no JS
            match = re.search(lib["pattern"], html, flags=re.IGNORECASE)
            if match:
                html = html.replace(match.group(0), inline_tag)
                libs_embutidas += 1
                print(f"     \U0001f504 Substituido: {lib['name']}")
            else:
                print(f"     \u26a0\ufe0f  Pattern nao encontrado para {lib['name']} \u2014 inserindo antes de </head>")
                html = html.replace("</head>", f"{inline_tag}\n</head>", 1)
                libs_embutidas += 1
        else:
            erros += 1

    # --- Etapa 2: Remover Google Fonts ---
    print(f"\n\U0001f524 Etapa 2/3 \u2014 Substituindo Google Fonts por Segoe UI...")
    fonts_removed = 0
    font_fallback_inserted = False
    for pattern in GOOGLE_FONTS_PATTERNS:
        match = re.search(pattern, html, flags=re.IGNORECASE)
        while match:
            if not font_fallback_inserted:
                html = html.replace(match.group(0), FONT_FALLBACK_CSS)
                font_fallback_inserted = True
            else:
                html = html.replace(match.group(0), "")
            fonts_removed += 1
            match = re.search(pattern, html, flags=re.IGNORECASE)
    if not font_fallback_inserted:
        html = html.replace("</head>", f"{FONT_FALLBACK_CSS}\n</head>", 1)
    print(f"     \u2705 {fonts_removed} tag(s) Google Fonts removida(s)")
    print(f"     \u2705 Fallback CSS (Segoe UI) inserido")

    # --- Etapa 3: Verificacao final ---
    print(f"\n\U0001f50d Etapa 3/3 \u2014 Verificacao de referencias externas...")
    cdn_refs = len(re.findall(r'cdn\.jsdelivr\.net', html))
    google_refs = len(re.findall(r'fonts\.googleapis\.com', html))
    gstatic_refs = len(re.findall(r'fonts\.gstatic\.com', html))
    total_ext = cdn_refs + google_refs + gstatic_refs

    s1 = "\u2705" if cdn_refs == 0 else "\u26a0\ufe0f VERIFICAR!"
    s2 = "\u2705" if google_refs == 0 else "\u26a0\ufe0f VERIFICAR!"
    s3 = "\u2705" if gstatic_refs == 0 else "\u26a0\ufe0f VERIFICAR!"
    print(f"     CDN jsdelivr:      {cdn_refs} {s1}")
    print(f"     Google Fonts:      {google_refs} {s2}")
    print(f"     Google Gstatic:    {gstatic_refs} {s3}")

    if total_ext == 0:
        print(f"\n  \U0001f389 HTML 100% self-contained! Zero dependencias externas.")
    else:
        print(f"\n  \u26a0\ufe0f  {total_ext} referencias externas restantes. Revise manualmente.")

    print(f"\n  \U0001f4ca Resumo embedding:")
    print(f"     Libs JS embutidas: {libs_embutidas}")
    print(f"     Google Fonts:      removido \u2192 Segoe UI")
    print(f"     Erros de download: {erros}")
    print("=" * 60)

    return html, libs_embutidas, erros


# =========================================================
# FUNCOES DE DEPARA
# =========================================================

def ler_depara(path):
    """Le tabela De-Para Produtos.

    Espera colunas:
    - PRODUTO_SELLIN: nome canonico do produto no sell-in
    - PRODUTO_SELLOUT: nome do produto no sell-out
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
    for c in cols_esperadas[:2]:
        if c not in df.columns:
            print(f"[ERRO] Coluna obrigatoria '{c}' ausente no DePara. Coluna sera ignorada.")
            return {}, set()

    if "INCLUIR_DASHBOARD" not in df.columns:
        df["INCLUIR_DASHBOARD"] = "SIM"

    def norm(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    mapeamento = {}
    ignorar = set()
    sellins_unicos = set()
    duplicatas = []

    for _, row in df.iterrows():
        sellin = norm(row["PRODUTO_SELLIN"])
        sellout = norm(row["PRODUTO_SELLOUT"])
        incluir = norm(row.get("INCLUIR_DASHBOARD", "SIM"))

        if not sellout:
            continue

        if incluir == "NAO":
            ignorar.add(sellout)
            continue

        if sellout in mapeamento:
            sellin_anterior = mapeamento[sellout]
            duplicatas.append({
                "PRODUTO_SELLOUT": sellout,
                "MAPEAMENTO_USADO": sellin_anterior,
                "MAPEAMENTO_IGNORADO": sellin or "(vazio)",
            })
            continue

        if sellin:
            mapeamento[sellout] = sellin
            sellins_unicos.add(sellin)
        else:
            mapeamento[sellout] = sellout

    print(f"     Produtos sell-in unicos: {len(sellins_unicos)}")
    print(f"     Pares sell-in <-> sell-out: {len(mapeamento)}")
    print(f"     Produtos a ignorar: {len(ignorar)}")

    if duplicatas:
        print()
        print(f"     [AVISO] {len(duplicatas)} duplicatas detectadas em PRODUTO_SELLOUT")
        print(f"     (mantido o primeiro mapeamento, demais ignorados)")
        print(f"     Top 5:")
        for d in duplicatas[:5]:
            print(f"       - '{d['PRODUTO_SELLOUT']}' usado: '{d['MAPEAMENTO_USADO']}' | ignorado: '{d['MAPEAMENTO_IGNORADO']}'")

        try:
            csv_path = str(Path(path).parent / "DePara_duplicatas.csv")
            pd.DataFrame(duplicatas).to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"     CSV de duplicatas: {csv_path}")
        except Exception as e:
            print(f"     [WARN] Nao foi possivel salvar CSV: {e}")

    return mapeamento, ignorar


def ler_depara_clientes(path):
    """Le tabela De-Para de Clientes.

    Espera colunas:
    - CLIENTE_SELLOUT: nome do cliente no sell-out (GRUPO_PAINEL)
    - CLIENTE_SELLIN_SUGESTAO: nome correspondente no sell-in (GRUPO_CLIENTE_FINAL)

    Retorna dict {cliente_sellout_norm -> cliente_sellin}
    Clientes NAO listados no DePara mantem o nome original (nao sao excluidos).
    """
    if not Path(path).exists():
        print(f"[INFO] DePara de Clientes nao encontrado em {path}")
        print(f"       Clientes sell-out manterao o nome original para match.")
        return {}

    print(f"[OK] Lendo DePara de Clientes: {path}")
    df = pd.read_excel(path)

    if "CLIENTE_SELLOUT" not in df.columns or "CLIENTE_SELLIN_SUGESTAO" not in df.columns:
        print(f"[ERRO] Colunas esperadas: CLIENTE_SELLOUT, CLIENTE_SELLIN_SUGESTAO")
        print(f"       Colunas encontradas: {list(df.columns)}")
        return {}

    def norm(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    mapeamento = {}
    for _, row in df.iterrows():
        so = norm(row["CLIENTE_SELLOUT"])
        si = norm(row["CLIENTE_SELLIN_SUGESTAO"])
        if so and si:
            mapeamento[so] = si

    print(f"     Pares mapeados: {len(mapeamento)}")
    return mapeamento


def aplicar_depara_clientes(df_sellout, mapeamento_clientes):
    """Aplica De-Para de Clientes no DataFrame de sell-out.

    - Clientes COM mapeamento: GRUPO_PAINEL recebe o nome do sell-in
    - Clientes SEM mapeamento: GRUPO_PAINEL mantem o nome original (NAO exclui)
    """
    if df_sellout is None or len(df_sellout) == 0:
        return df_sellout

    if not mapeamento_clientes:
        print("     DePara de Clientes vazio - nomes originais mantidos")
        return df_sellout

    def norm(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    df_sellout["CLIENTE_ORIGINAL"] = df_sellout["GRUPO_PAINEL"]
    df_sellout["_CLI_NORM"] = df_sellout["GRUPO_PAINEL"].apply(norm)
    df_sellout["GRUPO_PAINEL"] = df_sellout["_CLI_NORM"].map(mapeamento_clientes).fillna(df_sellout["GRUPO_PAINEL"])

    mapeados = df_sellout["_CLI_NORM"].isin(mapeamento_clientes.keys()).sum()
    nao_mapeados = len(df_sellout) - mapeados
    clientes_unicos_mapeados = df_sellout[df_sellout["_CLI_NORM"].isin(mapeamento_clientes.keys())]["_CLI_NORM"].nunique()
    clientes_unicos_nao = df_sellout[~df_sellout["_CLI_NORM"].isin(mapeamento_clientes.keys())]["_CLI_NORM"].nunique()

    print(f"     Clientes mapeados:      {clientes_unicos_mapeados} ({mapeados:,} linhas)")
    print(f"     Clientes sem DePara:    {clientes_unicos_nao} ({nao_mapeados:,} linhas) - mantidos com nome original")

    df_sellout = df_sellout.drop(columns=["_CLI_NORM"])
    return df_sellout


def aplicar_depara_e_relatorio(df_sellout, mapeamento, ignorar, path_csv_nao_mapeados):
    """Aplica de-para no DataFrame de sell-out.

    v5.5: Produtos sem mapeamento NAO sao mais excluidos.
    Eles mantem o nome original do sell-out e aparecem no dashboard.
    O CSV de nao-mapeados e gerado apenas como REFERENCIA.

    1. Filtra produtos da lista de ignorar (INCLUIR_DASHBOARD=NAO)
    2. Adiciona coluna PRODUTO_CANONICO (nome do sell-in para cruzamento)
    3. Salva produtos sem mapeamento em CSV (informativo, nao exclui)
    """
    if df_sellout is None or len(df_sellout) == 0:
        return df_sellout

    if not mapeamento and not ignorar:
        df_sellout["PRODUTO_CANONICO"] = df_sellout["PRODUTO"]
        return df_sellout

    def normalizar(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    df_sellout["_PROD_NORM"] = df_sellout["PRODUTO"].apply(normalizar)

    # 1. Filtrar APENAS produtos explicitamente marcados como NAO
    antes = len(df_sellout)
    df_sellout = df_sellout[~df_sellout["_PROD_NORM"].isin(ignorar)].copy()
    depois = len(df_sellout)
    if antes != depois:
        print(f"     Linhas filtradas (INCLUIR_DASHBOARD=NAO): {antes - depois:,}")

    # 2. Aplicar mapeamento - produtos SEM mapeamento mantem nome original
    df_sellout["PRODUTO_CANONICO"] = df_sellout["_PROD_NORM"].map(mapeamento).fillna(df_sellout["PRODUTO"])

    # 3. Identificar nao-mapeados (apenas para CSV informativo - NAO exclui!)
    nao_mapeados_mask = ~df_sellout["_PROD_NORM"].isin(mapeamento.keys()) & ~df_sellout["_PROD_NORM"].isin(ignorar)
    nao_mapeados = df_sellout[nao_mapeados_mask]

    prod_mapeados = df_sellout[~nao_mapeados_mask]["_PROD_NORM"].nunique()
    prod_nao_mapeados = nao_mapeados["_PROD_NORM"].nunique()

    print(f"     Produtos mapeados:      {prod_mapeados}")
    print(f"     Produtos sem DePara:    {prod_nao_mapeados} - mantidos com nome original")

    if len(nao_mapeados) > 0:
        agg = nao_mapeados.groupby("PRODUTO").agg(
            BRL=("BRL", "sum"),
            UNID=("UNID", "sum"),
            LINHAS=("PRODUTO", "size")
        ).reset_index().sort_values("UNID", ascending=False)

        agg["PRODUTO_SELLIN_SUGESTAO"] = ""
        agg["INCLUIR_DASHBOARD"] = "SIM"

        agg = agg[["PRODUTO", "PRODUTO_SELLIN_SUGESTAO", "INCLUIR_DASHBOARD", "BRL", "UNID", "LINHAS"]]
        agg = agg.rename(columns={"PRODUTO": "PRODUTO_SELLOUT"})

        agg.to_csv(path_csv_nao_mapeados, index=False, encoding="utf-8-sig")

        print()
        print("=" * 60)
        print(f"  PRODUTOS SELL-OUT SEM DEPARA (mantidos no dashboard)")
        print("=" * 60)
        print(f"  Total: {len(agg)} produtos | {len(nao_mapeados):,} linhas")
        print(f"  Volume BRL: R$ {nao_mapeados['BRL'].sum():,.0f}")
        print(f"  CSV referencia: {path_csv_nao_mapeados}")
        print()
        print(f"  Top 10 por volume (UNID):")
        for _, row in agg.head(10).iterrows():
            print(f"    {row['UNID']:>10,.0f} unid | R$ {row['BRL']:>12,.0f} - {row['PRODUTO_SELLOUT']}")
        print()
        print(f"  NOTA: Esses produtos JA aparecem no dashboard com o nome do sell-out.")
        print(f"  Para melhorar o cruzamento SI x SO, preencha o DePara e rode novamente.")
        print("=" * 60)

        # v5.5: NAO remove mais! Apenas informa.
        # (linha removida: df_sellout = df_sellout[~nao_mapeados_mask].copy())

    df_sellout = df_sellout.drop(columns=["_PROD_NORM"])
    print(f"     Linhas finais sell-out: {len(df_sellout):,}")
    return df_sellout


# =========================================================
# FUNCOES DE LEITURA DE DADOS
# =========================================================

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
            if k in ("CHAN_DESC", "UF"):
                print(f"[AVISO] Coluna '{col_name}' nao encontrada - feature relacionada nao funcionara")
                continue
            print(f"[ERRO] Coluna '{col_name}' nao encontrada. Ajuste COLS_SELLOUT no codigo.")
            return [], None

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
    mapeamento, ignorar = ler_depara(PATH_DEPARA)
    df_pivot["PRODUTO_ORIGINAL"] = df_pivot["PRODUTO"]
    df_pivot = aplicar_depara_e_relatorio(df_pivot, mapeamento, ignorar, PATH_NAO_MAPEADOS)
    if df_pivot is None or len(df_pivot) == 0:
        print("[AVISO] Sell-out vazio apos aplicar DePara")
        return [], None

    df_pivot["PRODUTO"] = df_pivot["PRODUTO_CANONICO"]
    df_pivot = df_pivot.drop(columns=["PRODUTO_CANONICO"])

    # v5.5: APLICAR DE-PARA DE CLIENTES
    mapeamento_clientes = ler_depara_clientes(PATH_DEPARA_CLIENTES)
    df_pivot = aplicar_depara_clientes(df_pivot, mapeamento_clientes)

    # =================================================================
    # CRIAR 2 ESTRUTURAS SEPARADAS PARA OTIMIZAR PERFORMANCE
    # =================================================================

    df_principal = df_pivot.groupby(
        ["GRUPO_PAINEL", "FRANQUIA", "TIPO_CLIENTE", "PRODUTO", "ANO", "MES"],
        dropna=False
    ).agg(BRL=("BRL", "sum"), UNID=("UNID", "sum")).reset_index()

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
        ("GRUPO_CLIENTE_FINAL", "\u2014"),
        ("TIPO_CLIENTE_FINAL", "\u2014"),
        ("FRANQUIA", "\u2014"),
        ("PRODUTO", "\u2014"),
        ("FONTE", "\u2014"),
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


# =========================================================
# FUNCAO PRINCIPAL
# =========================================================

def gerar_html():
    df = ler_sellin(PATH_XLSX)
    targets = ler_targets(PATH_TARGETS)
    targets_fin = ler_targets_fin(PATH_TARGETS_FIN)
    sellout, ultimo_sellout, sellout_uf = ler_sellout_gerencial(PATH_SELLOUT)

    depara_ok = Path(PATH_DEPARA).exists()

    # Mapeamento cliente_sellin -> TIPO_CLIENTE_FINAL
    cliente_tipo_map = {}
    for cli, tipo in df.groupby([COLS["CLIENTE"], COLS["TIPO_CLIENTE"]]).groups.keys():
        if cli and tipo:
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
        '<span id="statusText">Carregando...</span>',
        f'<span id="statusText">{status_msg}</span>'
    )

    # =========================================================
    # v5.5 - EMBUTIR LIBS EXTERNAS (se habilitado)
    # =========================================================
    if EMBUTIR_LIBS:
        template, libs_ok, libs_err = embutir_libs_externas(template)
        if libs_err > 0:
            print(f"\n[AVISO] {libs_err} lib(s) nao foram embutidas (erro de download).")
            print(f"        O dashboard ainda funciona, mas precisa de internet para essas libs.")
    else:
        print("\n[INFO] EMBUTIR_LIBS = False - libs CDN mantidas como referencia externa.")

    Path(PATH_OUTPUT).write_text(template, encoding="utf-8")
    tamanho_mb = len(template.encode("utf-8")) / 1024 / 1024
    print(f"\n[OK] Gerado: {PATH_OUTPUT} ({tamanho_mb:.2f} MB)")

    if EMBUTIR_LIBS:
        print(f"\n{'='*60}")
        print(f"  \U0001f389 DASHBOARD STANDALONE GERADO COM SUCESSO!")
        print(f"{'='*60}")
        print(f"  Arquivo: {PATH_OUTPUT}")
        print(f"  Tamanho: {tamanho_mb:.2f} MB")
        print(f"  Status:  100% self-contained (zero dependencias externas)")
        print(f"  Teste:   Abra o arquivo com internet DESLIGADA")
        print(f"  Deploy:  Pronto para request IIH!")
        print(f"{'='*60}")
    else:
        print(f"\n{'='*60}")
        print(f"Pronto! Arquivo: {PATH_OUTPUT}")
        print(f"{'='*60}")


if __name__ == "__main__":
    gerar_html()
