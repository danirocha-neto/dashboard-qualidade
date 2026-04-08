import pandas as pd
import numpy as np
from datetime import date, timedelta
from io import StringIO
import urllib.request
import re

# ─────────────────────────────────────────────
#  CONFIGURAÇÃO — altere apenas aqui se mudar a planilha
# ─────────────────────────────────────────────
SHEET_ID       = '2PACX-1vQFihVCoJyeNl66YtugzUD2_e7z5SysjO5TEkiWtkmBKb0VF1WNUen65LvPMJeHkfnd0OXMw3ZP0YLX'
GID_AUDITORIAS = 351234331
GID_DADOS      = 0

FIREBASE_CONFIG = """{
  apiKey: "AIzaSyBZQPAI_iae5FxROXfBH3TK3yqwJuRlRws",
  authDomain: "dashboard-qualidade-f96d2.firebaseapp.com",
  projectId: "dashboard-qualidade-f96d2",
  storageBucket: "dashboard-qualidade-f96d2.firebasestorage.app",
  messagingSenderId: "574600100930",
  appId: "1:574600100930:web:b9c5cc2717abd5873541db"
}"""

# ─────────────────────────────────────────────
#  FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────
def fetch_csv(gid):
    url = (f'https://docs.google.com/spreadsheets/d/e/{SHEET_ID}'
           f'/pub?output=csv&gid={gid}')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as r:
        return pd.read_csv(StringIO(r.read().decode('utf-8')))

def d_util_atras(ref, n):
    d, c = ref, 0
    while c < n:
        d -= timedelta(days=1)
        if d.weekday() < 5:
            c += 1
    return d

def normalizar_marca(m):
    if pd.isna(m): return None
    m = str(m).strip()
    mapa = {
        'fast tennis':'Fast Tennis','airlocker ':'Airlocker','airlocker':'Airlocker',
        'ecoville ':'Ecoville','ecoville':'Ecoville','shelf':'Shelf','shelf ':'Shelf',
        'locar-x':'Locar X','locar x ':'Locar X','la bolaria':'La Bolaria',
        'brumed ':'Brumed','sua hora unha':'Sua Hora Unha','mentoria ':'Mentoria',
        'saúde livre ':'Saúde Livre','saúde livre vacinas':'Saúde Livre',
        'lypedepil':'Lypedepil','lypedepyl':'Lypedepil','4beach':'4 Beach',
    }
    return mapa.get(m.lower(), m)

def safe_key(s):
    return re.sub(r'[^a-zA-Z0-9]', '', str(s))

def achar_coluna(colunas, possiveis):
    for p in possiveis:
        if p in colunas: return p
    for p in possiveis:
        found = next((c for c in colunas if p.lower() in c.lower()), None)
        if found: return found
    return None

# ─────────────────────────────────────────────
#  LÓGICA PRINCIPAL
# ─────────────────────────────────────────────
def gerar():

    # ── 1. Carregar Auditorias ──────────────────
    df = fetch_csv(GID_AUDITORIAS)
    df.columns = [c.strip() for c in df.columns]
    print('Colunas Auditorias:', df.columns.tolist())

    if not any('tipo' in c.lower() for c in df.columns):
        raise ValueError(f'Aba errada carregada. Colunas: {df.columns.tolist()}. Verifique GID_AUDITORIAS.')

    col_data  = achar_coluna(df.columns, ['Data da Auditoria','Data','data'])
    # col E = Checklist de Processo (tipos especiais)
    col_E     = achar_coluna(df.columns, ['Checklist de Processo'])
    # col F = Checklist da Qualificação (Ligação)
    col_F     = achar_coluna(df.columns, ['Checklist da Qualificação','Checklist da Qualidade','Checklist da Qualificacao'])
    col_tipo  = achar_coluna(df.columns, ['Tipo de Auditoria','Tipo'])
    col_marca = achar_coluna(df.columns, ['Marca','marca'])

    print(f'col_data={col_data} | col_E={col_E} | col_F={col_F} | col_tipo={col_tipo} | col_marca={col_marca}')

    for col, nome in [(col_data,'Data'),(col_E,'Checklist de Processo'),(col_F,'Checklist da Qualificação'),(col_tipo,'Tipo'),(col_marca,'Marca')]:
        if not col:
            raise ValueError(f'Coluna "{nome}" não encontrada. Colunas disponíveis: {df.columns.tolist()}')

    df[col_data]  = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
    df = df.dropna(subset=[col_data, col_tipo])
    df[col_tipo]  = df[col_tipo].str.strip().replace({'Treinamentos':'Treinamento'})
    df[col_marca] = df[col_marca].apply(normalizar_marca)

    # ── 2. Datas de referência (ignora fins de semana) ──
    hoje       = df[col_data].dt.date.max()
    d1_data    = hoje if hoje.weekday() < 5 else d_util_atras(hoje, 1)
    d7_inicio  = d_util_atras(hoje, 7)
    d30_inicio = d_util_atras(hoje, 30)

    ref_str = hoje.strftime('%d/%m/%Y')
    d1_str  = d1_data.strftime('%d/%m')
    d7_str  = d7_inicio.strftime('%d/%m')
    d30_str = d30_inicio.strftime('%d/%m')

    r1  = df[df[col_data].dt.date == d1_data]
    r7  = df[df[col_data].dt.date >= d7_inicio]
    r30 = df[df[col_data].dt.date >= d30_inicio]
    total = len(df)

    print(f'Total={total} | D1={len(r1)} ({d1_str}) | D7={len(r7)} | D30={len(r30)}')

    # ── 3. Tipos ────────────────────────────────
    excluir = {'Treinamento','Treinamentos','Gestão á Vista','Gestão à Vista'}
    tipos = sorted([t for t in df[col_tipo].dropna().unique()
                    if t not in excluir and str(t).strip()])
    tipos_esp = {'Spread.Chat','Volumetria','Lead Time','Cadência'}

    # ── 4. Quantidade por tipo ──────────────────
    linhas_tipo = ''
    tot1=tot7=tot30=0
    for t in tipos:
        c1  = len(r1[r1[col_tipo]==t])
        c7  = len(r7[r7[col_tipo]==t])
        c30 = len(r30[r30[col_tipo]==t])
        tot1+=c1; tot7+=c7; tot30+=c30
        s1  = f'<span class="tr">{c1}</span>' if c1 else '<span class="tz">—</span>'
        s7  = f'<span class="tr">{c7}</span>' if c7 else '<span class="tz">—</span>'
        linhas_tipo += f'<tr><td class="tl">{t}</td><td>{s1}</td><td class="sep">{s7}</td><td class="sep tr">{c30}</td></tr>\n'
    linhas_tipo += f'<tr class="tr-total"><td>Total</td><td class="tr">{tot1 or "—"}</td><td class="tr sep">{tot7}</td><td class="tr sep">{tot30}</td></tr>'

    # ── 5. Conformidade por tipo ─────────────────
    # Ligação  → col F (Checklist da Qualificação): % conf, % nao, % cor
    # Especiais → col E (Checklist de Processo): pontos = 100-qtd_NC, % NC, % Cor
    def calc_conf_tipo(r, t):
        if t == 'Ligação':
            gv = r[(r[col_tipo]==t) & r[col_F].notna()]
            tot = len(gv)
            if tot == 0: return None
            conf = round(len(gv[gv[col_F]=='Conforme'])/tot*100, 1)
            nao  = round(len(gv[gv[col_F]=='Não Conforme'])/tot*100, 1)
            cor  = round(len(gv[gv[col_F]=='Corrigido'])/tot*100, 1)
            return {'tot':tot, 'conf':conf, 'nao':nao, 'cor':cor, 'modo':'lig'}
        else:
            gv = r[(r[col_tipo]==t) & r[col_E].notna()]
            tot = len(gv)
            if tot == 0: return None
            nc  = len(gv[gv[col_E]=='Não Conforme'])
            cor = len(gv[gv[col_E]=='Corrigido'])
            pts = max(0, 100 - nc)
            nao_pct = round(nc/tot*100, 1) if tot > 0 else 0
            cor_pct = round(cor/tot*100, 1) if tot > 0 else 0
            return {'tot':tot, 'conf':pts, 'nao':nao_pct, 'cor':cor_pct, 'modo':'esp'}

    def conf_tipo_js(r):
        rows = []
        for t in tipos:
            res = calc_conf_tipo(r, t)
            if res:
                rows.append(f'{{tipo:"{t}",tot:{res["tot"]},conf:{res["conf"]},nao:{res["nao"]},cor:{res["cor"]},modo:"{res["modo"]}"}}'  )
            else:
                rows.append(f'{{tipo:"{t}",tot:0,conf:null,nao:null,cor:null,modo:""}}')
        return '[' + ',\n    '.join(rows) + ']'

    conf_tipo_d1  = conf_tipo_js(r1)
    conf_tipo_d7  = conf_tipo_js(r7)
    conf_tipo_d30 = conf_tipo_js(r30)

    # ── 6. Conformidade por marca ────────────────
    marcas_ord = (r30.groupby(col_marca).size()
                     .sort_values(ascending=False).index.tolist())
    marcas_ord = [m for m in marcas_ord if m and not pd.isna(m)]

    def calc_marca(m, filtro, r):
        tipo_map = {
            'Ligacao':'Ligação','SpreadChat':'Spread.Chat',
            'Volumetria':'Volumetria','LeadTime':'Lead Time',
            'Cadencia':'Cadência'
        }
        if filtro == 'Ligacao':
            g  = r[(r[col_marca]==m) & (r[col_tipo]=='Ligação')]
            gv = g[g[col_F].notna()]
            tot = len(gv)
            if tot == 0: return 'null'
            conf = round(len(gv[gv[col_F]=='Conforme'])/tot*100, 1)
            nao  = round(len(gv[gv[col_F]=='Não Conforme'])/tot*100, 1)
            cor  = round(len(gv[gv[col_F]=='Corrigido'])/tot*100, 1)
            return f'{{conf:{conf},nao:{nao},cor:{cor},modo:"lig"}}'
        elif filtro == 'geral':
            lig = r[(r[col_marca]==m) & (r[col_tipo]=='Ligação')]
            out = r[(r[col_marca]==m) & (r[col_tipo].isin(tipos_esp))]
            gl  = lig[lig[col_F].notna()]
            go  = out[out[col_E].notna()]
            tot = len(gl) + len(go)
            if tot == 0: return 'null'
            ct  = len(gl[gl[col_F]=='Conforme'])  + len(go[go[col_E]=='Conforme'])
            nt  = len(gl[gl[col_F]=='Não Conforme']) + len(go[go[col_E]=='Não Conforme'])
            cot = len(gl[gl[col_F]=='Corrigido'])  + len(go[go[col_E]=='Corrigido'])
            return f'{{conf:{round(ct/tot*100,1)},nao:{round(nt/tot*100,1)},cor:{round(cot/tot*100,1)},modo:"geral"}}'
        else:
            t  = tipo_map.get(filtro, filtro)
            g  = r[(r[col_marca]==m) & (r[col_tipo]==t)]
            gv = g[g[col_E].notna()]
            tot = len(gv)
            if tot == 0: return 'null'
            nc  = len(gv[gv[col_E]=='Não Conforme'])
            cor = len(gv[gv[col_E]=='Corrigido'])
            nao_pct = round(nc/tot*100, 1)
            cor_pct = round(cor/tot*100, 1)
            return f'{{conf:{max(0,100-nc)},nao:{nao_pct},cor:{cor_pct},modo:"esp"}}'

    def js_mapa(filtro):
        return '{' + ','.join(f'"{m}":{calc_marca(m,filtro,r30)}' for m in marcas_ord) + '}'

    conf_js = (
        f'const confLigacao    = {js_mapa("Ligacao")};\n'
        f'const confSpreadChat = {js_mapa("SpreadChat")};\n'
        f'const confVolumetria = {js_mapa("Volumetria")};\n'
        f'const confLeadTime   = {js_mapa("LeadTime")};\n'
        f'const confCadencia   = {js_mapa("Cadencia")};\n'
        f'const confgeral      = {js_mapa("geral")};\n'
    )

    marcas_js = 'const marcasData = [\n'
    for m in marcas_ord:
        mc1  = len(r1[r1[col_marca]==m])
        mc7  = len(r7[r7[col_marca]==m])
        mc30 = len(r30[r30[col_marca]==m])
        marcas_js += f'  {{nome:"{m}",d1:{mc1},d7:{mc7},d30:{mc30},key:"pa-{safe_key(m)}"}},\n'
    marcas_js += '];'

    # ── 7. Diretorias ───────────────────────────
    df2 = fetch_csv(GID_DADOS)
    df2.columns = [c.strip() for c in df2.columns]
    print('Colunas Dados:', df2.columns.tolist())

    col_dia   = achar_coluna(df2.columns, ['Dia','Data','data','dia'])
    col_media = achar_coluna(df2.columns, ['Média','Media','média','media'])
    col_dir   = achar_coluna(df2.columns, ['Diretoria','diretoria'])
    col_meta  = achar_coluna(df2.columns, ['Meta','meta'])

    print(f'col_dia={col_dia} | col_media={col_media} | col_dir={col_dir}')

    df2[col_dia]   = pd.to_datetime(df2[col_dia], dayfirst=True, errors='coerce')
    df2[col_media] = pd.to_numeric(df2[col_media], errors='coerce')
    df2 = df2.dropna(subset=[col_media, col_dia])

    dr7  = df2[df2[col_dia].dt.date >= d7_inicio]
    dr30 = df2[df2[col_dia].dt.date >= d30_inicio]
    dirs = sorted(df2[col_dir].dropna().unique())
    meta = int(df2[col_meta].iloc[0]) if col_meta else 200

    dir_js = 'const dirData = {\n'
    for d in dirs:
        m7  = dr7[dr7[col_dir]==d][col_media].mean()
        m30 = dr30[dr30[col_dir]==d][col_media].mean()
        v7  = round(float(m7), 1) if not pd.isna(m7) else 'null'
        v30 = round(float(m30), 1) if not pd.isna(m30) else 'null'
        dir_js += f'  "{d}":{{d7:{v7},d30:{v30}}},\n'
    dir_js += '};'

    # ── 8. Montar HTML ──────────────────────────
    html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard de Qualidade</title>
<style>
:root {{
  --bg:#161614;--surface:#202020;--surface2:#1a1a18;
  --border:rgba(255,255,255,.08);--text:#eeece6;--muted:#86847c;
  --green:#1a8f68;--blue:#1558a0;--amber:#a86a10;--red:#922828;
  --purple:#4a42a8;--radius:12px;--qual-bg:#1e1a2a;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:var(--bg);color:var(--text);font-family:-apple-system,'Segoe UI',system-ui,sans-serif;
  font-size:11px;line-height:1.3;padding:8px 14px;max-width:1400px;margin:0 auto;}}
header{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;margin-bottom:8px;}}
h1{{font-size:14px;font-weight:700;letter-spacing:-.3px;}}
.sub{{font-size:9.5px;color:var(--muted);margin-top:1px;}}
.badge{{font-size:9.5px;background:var(--surface);border:.5px solid var(--border);border-radius:20px;padding:2px 9px;color:var(--muted);}}
.section-label{{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:5px;padding-left:2px;display:flex;align-items:center;gap:8px;}}
.section-label::after{{content:'';flex:1;height:.5px;background:var(--border);}}
.section-quantitativa{{margin-bottom:8px;}}
.section-qualitativa{{background:var(--qual-bg);border-radius:12px;padding:8px 12px 10px;margin-bottom:8px;}}
.metrics-row{{display:grid;grid-template-columns:repeat(7,1fr);gap:6px;margin-bottom:8px;}}
.metric{{background:var(--surface);border:.5px solid var(--border);border-radius:var(--radius);padding:5px 9px;}}
.metric-label{{font-size:7.5px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:1px;}}
.metric-value{{font-size:17px;font-weight:700;letter-spacing:-1px;line-height:1;}}
.metric-sub{{font-size:8px;color:var(--muted);margin-top:1px;}}
.m-blue .metric-value{{color:var(--blue);}}.m-green .metric-value{{color:var(--green);}}
.m-purple .metric-value{{color:var(--purple);}}.m-amber .metric-value{{color:var(--amber);}}
.m-pos .metric-value{{color:var(--green);}}.m-neg .metric-value{{color:var(--red);}}
.main-grid{{display:grid;grid-template-columns:1fr 1.2fr;gap:8px;align-items:start;}}
.col-left{{display:flex;flex-direction:column;gap:8px;}}
.col-right{{display:flex;flex-direction:column;}}
.card{{background:var(--surface);border:.5px solid var(--border);border-radius:var(--radius);overflow:hidden;}}
.card-header{{display:flex;align-items:center;gap:5px;padding:5px 10px;border-bottom:.5px solid var(--border);background:var(--surface2);flex-wrap:wrap;}}
.dot{{width:6px;height:6px;border-radius:50%;flex-shrink:0;}}
.d-blue{{background:var(--blue);}}.d-green{{background:var(--green);}}
.d-amber{{background:var(--amber);}}.d-purple{{background:var(--purple);}}
.card-title{{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);}}
.periodo-filter{{display:flex;gap:0;margin-left:auto;}}
.pf-btn{{padding:2px 7px;font-size:9px;font-weight:600;border:.5px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;font-family:inherit;transition:all .15s;}}
.pf-btn:first-child{{border-radius:4px 0 0 4px;}}.pf-btn:last-child{{border-radius:0 4px 4px 0;}}
.pf-btn:not(:first-child){{border-left:none;}}.pf-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue);}}
.tipo-filter{{display:flex;gap:3px;flex-wrap:wrap;margin-left:auto;}}
.tf-btn{{padding:2px 6px;font-size:9px;font-weight:600;border:.5px solid var(--border);border-radius:20px;background:transparent;color:var(--muted);cursor:pointer;font-family:inherit;transition:all .15s;white-space:nowrap;}}
.tf-btn.active{{background:var(--amber);color:#fff;border-color:var(--amber);}}
.meta-selector{{display:flex;align-items:center;gap:5px;margin-left:auto;}}
.meta-selector label{{font-size:8.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.3px;}}
.meta-selector select{{padding:2px 6px;font-size:9px;border:.5px solid var(--border);border-radius:4px;background:var(--surface);color:var(--text);font-family:inherit;cursor:pointer;outline:none;}}
.tbl-wrap{{overflow-x:auto;}}
table{{width:100%;border-collapse:collapse;font-size:10px;}}
thead{{background:var(--surface2);}}
th{{padding:3px 7px;text-align:left;font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;color:var(--muted);white-space:nowrap;}}
th.r{{text-align:right;}}th.sep{{border-left:.5px solid var(--border);}}
td{{padding:4px 7px;border-top:.5px solid var(--border);white-space:nowrap;}}
tr:hover td{{background:var(--surface2);}}
.tl{{font-weight:500;}}.tr{{text-align:right;font-variant-numeric:tabular-nums;}}
.tz{{color:var(--muted);text-align:right;}}.sep{{border-left:.5px solid var(--border);}}
.c-g{{color:var(--green);font-weight:600;}}.c-r{{color:var(--red);font-weight:600;}}
.c-a{{color:var(--amber);font-weight:600;}}
.tr-total td{{font-weight:600;background:var(--surface2);border-top:1px solid var(--border);}}
.tbl-marcas-wrap{{overflow-x:auto;overflow-y:auto;}}
.tbl-marcas-wrap thead th{{position:sticky;top:0;z-index:2;background:var(--surface2);}}
.tbl-marcas-wrap .tr-total td{{position:sticky;bottom:0;z-index:1;background:var(--surface2);}}
.conf-cell{{display:flex;align-items:center;gap:4px;}}
.bola{{width:7px;height:7px;border-radius:50%;flex-shrink:0;}}
.bola-red{{background:var(--red);}}.bola-ok{{background:transparent;border:1.5px solid var(--green);}}
.conf-pct{{font-size:10px;font-weight:600;}}.conf-red{{color:var(--red);}}.conf-ok{{color:var(--green);}}
.conf-na{{color:var(--muted);font-size:9px;font-weight:400;}}
.plano-cell{{min-width:120px;max-width:155px;white-space:normal;}}
.plano-input{{width:100%;padding:2px 5px;border:.5px solid var(--border);border-radius:4px;background:var(--surface2);color:var(--text);font-family:inherit;font-size:10px;line-height:1.3;resize:none;outline:none;min-height:22px;transition:border-color .15s;}}
.plano-input:focus{{border-color:var(--blue);background:var(--surface);}}
.plano-input::placeholder{{color:var(--muted);}}
.bar-wrap{{display:flex;align-items:center;gap:4px;}}
.bar-bg{{flex:1;height:3px;background:var(--surface2);border-radius:3px;overflow:hidden;min-width:30px;}}
.bar-fill{{height:100%;border-radius:3px;transition:width .3s,background .3s;}}
.bar-pct{{font-size:9px;font-weight:600;min-width:30px;text-align:right;}}
.qual-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px;}}
.qual-bloco{{display:flex;flex-direction:column;gap:5px;}}
.qual-bloco-title{{font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:4px;padding-bottom:4px;border-bottom:.5px solid var(--border);}}
.qual-pos-title{{color:var(--green);}}.qual-neg-title{{color:var(--red);}}
.qual-filters{{display:flex;gap:5px;align-items:center;flex-wrap:wrap;}}
.qual-filter-input{{padding:2px 6px;border:.5px solid var(--border);border-radius:4px;background:var(--surface);color:var(--text);font-family:inherit;font-size:10px;outline:none;height:22px;}}
.qual-filter-input::placeholder{{color:var(--muted);}}
.qual-filter-date{{width:115px;}}.qual-filter-name{{flex:1;min-width:90px;}}
.qual-filter-label{{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;color:var(--muted);white-space:nowrap;}}
.behavior-list{{overflow-y:auto;max-height:130px;display:flex;flex-direction:column;gap:4px;padding-right:2px;}}
.behavior-entry{{background:var(--surface);border:.5px solid var(--border);border-radius:6px;padding:5px 7px;display:flex;flex-direction:column;gap:4px;flex-shrink:0;}}
.behavior-meta{{display:grid;grid-template-columns:1fr 110px auto;gap:4px;align-items:center;}}
.behavior-field{{padding:2px 6px;border:.5px solid var(--border);border-radius:4px;background:var(--surface2);color:var(--text);font-family:inherit;font-size:10px;outline:none;height:22px;width:100%;}}
.behavior-field::placeholder{{color:var(--muted);}}
.behavior-date{{padding:2px 6px;border:.5px solid var(--border);border-radius:4px;background:var(--surface2);color:var(--text);font-family:inherit;font-size:10px;outline:none;height:22px;width:100%;}}
.remove-btn{{padding:1px 6px;font-size:10px;border:.5px solid var(--border);border-radius:4px;background:transparent;color:var(--muted);cursor:pointer;font-family:inherit;white-space:nowrap;height:22px;}}
.remove-btn:hover{{border-color:var(--red);color:var(--red);}}
.behavior-desc{{width:100%;padding:3px 6px;border:.5px solid var(--border);border-radius:4px;background:var(--surface2);color:var(--text);font-family:inherit;font-size:10px;resize:none;outline:none;min-height:28px;line-height:1.3;}}
.behavior-desc::placeholder{{color:var(--muted);}}
.add-btn{{align-self:flex-start;padding:2px 8px;font-size:10px;font-weight:600;color:var(--blue);background:transparent;border:.5px solid var(--blue);border-radius:4px;cursor:pointer;font-family:inherit;}}
.add-btn.neg{{color:var(--red);border-color:var(--red);}}
.obs-label{{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);margin-bottom:2px;}}
.obs-box{{width:100%;padding:4px 7px;border:.5px solid var(--border);border-radius:5px;background:var(--surface);color:var(--text);font-family:inherit;font-size:10px;resize:vertical;outline:none;min-height:28px;line-height:1.4;}}
.obs-box::placeholder{{color:var(--muted);}}
.no-results{{font-size:10px;color:var(--muted);padding:6px 4px;text-align:center;}}
footer{{font-size:9px;color:var(--muted);text-align:center;padding:5px 0 2px;}}
</style>
</head>
<body>

<header>
  <div>
    <h1>Dashboard de Qualidade</h1>
    <div class="sub">Referência: <strong>{ref_str}</strong> &nbsp;·&nbsp; D-1 = {d1_str} &nbsp;·&nbsp; D-7 desde {d7_str} &nbsp;·&nbsp; D-30 desde {d30_str} &nbsp;·&nbsp; <em>fins de semana ignorados</em></div>
  </div>
  <div class="badge">📊 {total} auditorias · {ref_str}</div>
</header>

<div class="section-quantitativa">
<div class="section-label">Dados Quantitativos</div>
<div class="metrics-row">
  <div class="metric m-blue"><div class="metric-label">Auditorias D-1</div><div class="metric-value">{len(r1)}</div><div class="metric-sub">{d1_str}</div></div>
  <div class="metric m-green"><div class="metric-label">Auditorias D-7</div><div class="metric-value">{len(r7)}</div><div class="metric-sub">Últimos 7 dias úteis</div></div>
  <div class="metric m-purple"><div class="metric-label">Auditorias D-30</div><div class="metric-value">{len(r30)}</div><div class="metric-sub">Últimos 30 dias úteis</div></div>
  <div class="metric m-amber"><div class="metric-label">Total Auditorias 2026</div><div class="metric-value">{total}</div><div class="metric-sub">Aba Auditorias</div></div>
  <div class="metric"><div class="metric-label">Planos de Ação</div><div class="metric-value" id="contador-planos">0</div><div class="metric-sub">preenchidos</div></div>
  <div class="metric m-pos"><div class="metric-label">Comportamentos +</div><div class="metric-value" id="contador-pos">0</div><div class="metric-sub">positivos</div></div>
  <div class="metric m-neg"><div class="metric-label">Comportamentos −</div><div class="metric-value" id="contador-neg">0</div><div class="metric-sub">negativos</div></div>
</div>

<div class="main-grid">
<div class="col-left">
  <div class="card">
    <div class="card-header"><div class="dot d-blue"></div><div class="card-title">Quantidade por Tipo de Auditoria</div></div>
    <div class="tbl-wrap"><table>
      <thead><tr><th>Tipo</th><th class="r">D-1</th><th class="r sep">D-7</th><th class="r sep">D-30</th></tr></thead>
      <tbody>{linhas_tipo}</tbody>
    </table></div>
  </div>

  <div class="card">
    <div class="card-header">
      <div class="dot d-green"></div><div class="card-title">Conformidade por Tipo</div>
      <div class="periodo-filter">
        <button class="pf-btn" onclick="setConf('d1')" id="btn-d1">D-1</button>
        <button class="pf-btn" onclick="setConf('d7')" id="btn-d7">D-7</button>
        <button class="pf-btn active" onclick="setConf('d30')" id="btn-d30">D-30</button>
      </div>
    </div>
    <div class="tbl-wrap"><table>
      <thead><tr>
        <th>Tipo</th><th class="r">Total</th>
        <th class="r sep" style="color:var(--green)">Conforme</th>
        <th class="r sep" style="color:var(--red)">Não Conf.</th>
        <th class="r sep" style="color:var(--amber)">Corrigido</th>
      </tr></thead>
      <tbody id="tbody-conf"></tbody>
    </table></div>
  </div>

  <div class="card">
    <div class="card-header">
      <div class="dot d-purple"></div>
      <div class="card-title">Média de Ligações por Diretoria</div>
      <div class="meta-selector">
        <label>Meta:</label>
        <select id="select-meta" onchange="setMeta(this.value)">
          <option value="120">120 ligações</option>
          <option value="{meta}" selected>{meta} ligações</option>
        </select>
      </div>
    </div>
    <div class="tbl-wrap"><table>
      <thead><tr>
        <th>Diretoria</th>
        <th class="r">Méd. D-7</th><th class="r">% Meta</th>
        <th class="r sep">Méd. D-30</th><th class="r">% Meta</th>
      </tr></thead>
      <tbody id="tbody-dir"></tbody>
    </table></div>
  </div>
</div>

<div class="col-right">
  <div class="card" id="card-marcas">
    <div class="card-header">
      <div class="dot d-amber"></div>
      <div class="card-title">Auditorias por Marca · Conformidade D-30 &nbsp;·&nbsp; 🔴 &lt;90%</div>
      <div class="tipo-filter">
        <button class="tf-btn active" onclick="setTipo('Ligacao')" id="tf-Ligacao">Ligação</button>
        <button class="tf-btn" onclick="setTipo('SpreadChat')" id="tf-SpreadChat">Spread.Chat</button>
        <button class="tf-btn" onclick="setTipo('Volumetria')" id="tf-Volumetria">Volumetria</button>
        <button class="tf-btn" onclick="setTipo('LeadTime')" id="tf-LeadTime">Lead Time</button>
        <button class="tf-btn" onclick="setTipo('Cadencia')" id="tf-Cadencia">Cadência</button>
        <button class="tf-btn" onclick="setTipo('geral')" id="tf-geral">Geral</button>
      </div>
    </div>
    <div class="tbl-marcas-wrap" id="tbl-marcas-wrap">
      <table>
        <thead><tr>
          <th>Marca</th>
          <th class="r">D-1</th><th class="r sep">D-7</th><th class="r sep">D-30</th>
          <th class="sep" id="th-conf">Conformidade · Ligação</th>
          <th class="sep" style="color:var(--amber)">Plano de Ação</th>
        </tr></thead>
        <tbody id="tbody-marcas"></tbody>
      </table>
    </div>
  </div>
</div>
</div>
</div>

<div class="section-qualitativa">
<div class="section-label" style="margin-bottom:8px;">Análise Qualitativa de Comportamento</div>
<div class="qual-grid">
  <div class="qual-bloco">
    <div class="qual-bloco-title qual-pos-title">
      <svg width="10" height="10" viewBox="0 0 12 12" fill="none"><circle cx="6" cy="6" r="5.5" stroke="var(--green)" stroke-width="1.2"/><path d="M3 6l2 2 4-4" stroke="var(--green)" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>
      Comportamentos Positivos
    </div>
    <div class="qual-filters">
      <span class="qual-filter-label">Filtrar:</span>
      <input type="text" class="qual-filter-input qual-filter-name" id="f-pos-nome" placeholder="Por colaborador..." oninput="renderEntries('positivos')">
      <input type="date" class="qual-filter-input qual-filter-date" id="f-pos-data" oninput="renderEntries('positivos')">
    </div>
    <div class="behavior-list" id="lista-positivos"></div>
    <div id="no-pos" class="no-results" style="display:none;">Nenhum registro encontrado.</div>
    <button class="add-btn" onclick="addEntry('positivos')">+ Adicionar comportamento</button>
    <div><div class="obs-label">Observações gerais</div>
    <textarea class="obs-box" id="obs-positivo" placeholder="Observações gerais sobre comportamentos positivos..."></textarea></div>
  </div>
  <div class="qual-bloco">
    <div class="qual-bloco-title qual-neg-title">
      <svg width="10" height="10" viewBox="0 0 12 12" fill="none"><circle cx="6" cy="6" r="5.5" stroke="var(--red)" stroke-width="1.2"/><path d="M4 4l4 4M8 4l-4 4" stroke="var(--red)" stroke-width="1.3" stroke-linecap="round"/></svg>
      Comportamentos Negativos
    </div>
    <div class="qual-filters">
      <span class="qual-filter-label">Filtrar:</span>
      <input type="text" class="qual-filter-input qual-filter-name" id="f-neg-nome" placeholder="Por colaborador..." oninput="renderEntries('negativos')">
      <input type="date" class="qual-filter-input qual-filter-date" id="f-neg-data" oninput="renderEntries('negativos')">
    </div>
    <div class="behavior-list" id="lista-negativos"></div>
    <div id="no-neg" class="no-results" style="display:none;">Nenhum registro encontrado.</div>
    <button class="add-btn neg" onclick="addEntry('negativos')">+ Adicionar comportamento</button>
    <div><div class="obs-label">Observações gerais</div>
    <textarea class="obs-box" id="obs-negativo" placeholder="Observações gerais sobre comportamentos negativos..."></textarea></div>
  </div>
</div>
</div>

<footer>Gerado automaticamente em {ref_str} · Dashboard de Qualidade · {total} auditorias · fins de semana ignorados</footer>

<script type="module">
import {{ initializeApp }} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js';
import {{ getFirestore, doc, getDoc, setDoc, onSnapshot, collection }} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

const app = initializeApp({FIREBASE_CONFIG});
const db  = getFirestore(app);
const _c  = {{}};

async function fbSave(k,v){{_c[k]=v;try{{await setDoc(doc(db,'dashboard',k),{{value:v}});}}catch(e){{}}}}
async function fbLoad(k,def=''){{
  if(_c[k]!==undefined)return _c[k];
  try{{const s=await getDoc(doc(db,'dashboard',k));const v=s.exists()?s.data().value:def;_c[k]=v;return v;}}catch(e){{return def;}}
}}

/* ── DADOS ── */
const confTipoPeriodos = {{
  d1: {conf_tipo_d1},
  d7: {conf_tipo_d7},
  d30: {conf_tipo_d30}
}};

{conf_js}
{marcas_js}
{dir_js}

const mapaConf = {{Ligacao:confLigacao,SpreadChat:confSpreadChat,Volumetria:confVolumetria,LeadTime:confLeadTime,Cadencia:confCadencia,geral:confgeral}};
const labelTipo = {{Ligacao:'Ligação',SpreadChat:'Spread.Chat',Volumetria:'Volumetria',LeadTime:'Lead Time',Cadencia:'Cadência',geral:'Geral'}};
let tipoAtual='Ligacao', metaAtual={meta};

/* ── CONFORMIDADE POR TIPO ── */
function renderConf(p){{
  ['d1','d7','d30'].forEach(x=>document.getElementById('btn-'+x).classList.toggle('active',x===p));
  document.getElementById('tbody-conf').innerHTML = confTipoPeriodos[p].map(r=>{{
    if(r.tot===0) return `<tr><td class="tl">${{r.tipo}}</td><td class="tz">—</td><td class="sep tz">—</td><td class="sep tz">—</td><td class="sep tz">—</td></tr>`;
    if(r.modo==='lig'){{
      const cs=`<span class="c-g">${{String(r.conf).replace('.',',')}}%</span>`;
      const ns=r.nao>0?`<span class="c-r">${{String(r.nao).replace('.',',')}}%</span>`:'<span class="tz">—</span>';
      const cors=r.cor>0?`<span class="c-a">${{String(r.cor).replace('.',',')}}%</span>`:'<span class="tz">—</span>';
      return `<tr><td class="tl">${{r.tipo}}</td><td class="tr">${{r.tot}}</td><td class="sep">${{cs}}</td><td class="sep">${{ns}}</td><td class="sep">${{cors}}</td></tr>`;
    }} else {{
      const cs=`<span class="${{r.conf>=90?'c-g':r.conf>=70?'c-a':'c-r'}}">${{r.conf}}%</span>`;
      const ns=r.nao>0?`<span class="c-r">${{String(r.nao).replace('.',',')}}%</span>`:'<span class="tz">—</span>';
      const cors=r.cor>0?`<span class="c-a">${{String(r.cor).replace('.',',')}}%</span>`:'<span class="tz">—</span>';
      return `<tr><td class="tl">${{r.tipo}}</td><td class="tr">${{r.tot}}</td><td class="sep">${{cs}}</td><td class="sep">${{ns}}</td><td class="sep">${{cors}}</td></tr>`;
    }}
  }}).join('');
}}
window.setConf = p => renderConf(p);
renderConf('d30');

/* ── DIRETORIAS ── */
function barDir(val,sep=false){{
  if(!val||val==='null')return`<td class="${{sep?'sep ':''}}tz">—</td><td>—</td>`;
  const pct=Math.min(parseFloat(val)/metaAtual*100,100);
  const col=pct>=100?'var(--green)':pct>=85?'var(--amber)':'var(--red)';
  return`<td class="tr${{sep?' sep':''}}">${{String(val).replace('.',',')}}</td>
    <td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:${{Math.min(pct,100).toFixed(1)}}%;background:${{col}}"></div></div>
    <span class="bar-pct" style="color:${{col}}">${{pct.toFixed(1).replace('.',',')}}%</span></div></td>`;
}}
function renderDir(){{
  document.getElementById('tbody-dir').innerHTML=Object.entries(dirData).map(([d,v])=>
    `<tr><td class="tl">${{d}}</td>${{barDir(v.d7)}}${{barDir(v.d30,true)}}</tr>`
  ).join('');
}}
window.setMeta=v=>{{metaAtual=parseInt(v);renderDir();}};
renderDir();

/* ── MARCAS ── */
function confCell(nome,tipo){{
  const d=mapaConf[tipo]?.[nome];
  if(d===null||d===undefined)return'<div class="conf-cell"><span class="conf-na">Sem dados</span></div>';
  const ok=d.conf>=90;
  const lbl=`${{d.conf}}% Conf.`;
  return`<div class="conf-cell"><span class="bola ${{ok?'bola-ok':'bola-red'}}"></span><span class="conf-pct ${{ok?'conf-ok':'conf-red'}}">${{lbl}}</span></div>`;
}}
function planoCell(nome,tipo){{
  const d=mapaConf[tipo]?.[nome];
  if(!d||d===null)return`<td class="sep" style="color:var(--muted);font-size:9px;padding-left:8px;">—</td>`;
  if(d.conf<90){{const key=marcasData.find(m=>m.nome===nome)?.key||nome;return`<td class="sep plano-cell"><textarea class="plano-input" data-key="${{key}}" placeholder="Plano de ação..."></textarea></td>`;}}
  return`<td class="sep" style="color:var(--muted);font-size:9px;padding-left:8px;">—</td>`;
}}
function renderMarcas(){{
  const tbody=document.getElementById('tbody-marcas');let t1=0,t7=0,t30=0;
  tbody.innerHTML=marcasData.map(m=>{{t1+=m.d1;t7+=m.d7;t30+=m.d30;
    return`<tr><td class="tl">${{m.nome}}</td>
      <td>${{m.d1>0?m.d1:'<span class="tz">—</span>'}}</td>
      <td class="sep">${{m.d7>0?m.d7:'<span class="tz">—</span>'}}</td>
      <td class="sep tr">${{m.d30}}</td>
      <td class="sep">${{confCell(m.nome,tipoAtual)}}</td>
      ${{planoCell(m.nome,tipoAtual)}}</tr>`;
  }}).join('');
  tbody.innerHTML+=`<tr class="tr-total"><td>Total</td><td class="tr">${{t1||'—'}}</td><td class="tr sep">${{t7}}</td><td class="tr sep">${{t30}}</td><td class="sep"></td><td class="sep"></td></tr>`;
  document.querySelectorAll('.plano-input').forEach(async el=>{{
    const k=el.dataset.key;el.value=await fbLoad(k,'');resize(el);
    let tm;el.addEventListener('input',()=>{{resize(el);clearTimeout(tm);tm=setTimeout(()=>{{fbSave(k,el.value);atualizarContador();}},600);}});
  }});
  atualizarContador();
}}
window.setTipo=function(tipo){{
  tipoAtual=tipo;
  document.querySelectorAll('.tf-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tf-'+tipo)?.classList.add('active');
  document.getElementById('th-conf').textContent='Conformidade · '+labelTipo[tipo];
  renderMarcas();
}};

/* ── UTILS ── */
function resize(el){{el.style.height='auto';el.style.height=Math.max(22,el.scrollHeight)+'px';}}
function atualizarContador(){{let n=0;document.querySelectorAll('.plano-input').forEach(el=>{{if(el.value.trim())n++;}});document.getElementById('contador-planos').textContent=n;}}
function atualizarContadores(){{document.getElementById('contador-pos').textContent=Array.isArray(_c['positivos'])?_c['positivos'].length:0;document.getElementById('contador-neg').textContent=Array.isArray(_c['negativos'])?_c['negativos'].length:0;}}
function escHtml(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}}

/* ── COMPORTAMENTOS ── */
function buildEntry(tipo,idx,data){{
  const div=document.createElement('div');div.className='behavior-entry';div.dataset.idx=idx;
  const today=new Date().toISOString().split('T')[0];
  div.innerHTML=`<div class="behavior-meta">
    <input type="text" class="behavior-field" placeholder="Nome do colaborador..." value="${{escHtml(data.nome)}}" data-field="nome">
    <input type="date" class="behavior-date" value="${{escHtml(data.data||today)}}" data-field="data">
    <button class="remove-btn" onclick="removeEntry('${{tipo}}',${{idx}})">✕ Remover</button>
  </div>
  <textarea class="behavior-desc" placeholder="Descreva o comportamento observado..." data-field="descricao">${{escHtml(data.descricao)}}</textarea>`;
  let tm;div.querySelectorAll('[data-field]').forEach(el=>el.addEventListener('input',()=>{{clearTimeout(tm);tm=setTimeout(()=>saveEntries(tipo),600);}}));
  return div;
}}
function renderEntries(tipo){{
  const lista=document.getElementById('lista-'+tipo);
  const noRes=document.getElementById('no-'+(tipo==='positivos'?'pos':'neg'));
  const entries=_c[tipo]||[];
  const fN=(document.getElementById('f-'+(tipo==='positivos'?'pos':'neg')+'-nome')?.value||'').toLowerCase().trim();
  const fD=document.getElementById('f-'+(tipo==='positivos'?'pos':'neg')+'-data')?.value||'';
  const filtered=entries.map((e,i)=>{{return{{...e,_idx:i}};}}).filter(e=>(!fN||(e.nome||'').toLowerCase().includes(fN))&&(!fD||(e.data||'')===fD));
  lista.innerHTML='';
  if(filtered.length===0&&entries.length>0){{noRes.style.display='block';}}
  else{{noRes.style.display='none';filtered.forEach(e=>lista.appendChild(buildEntry(tipo,e._idx,e)));}}
  if(entries.length===0)noRes.style.display='none';
}}
async function addEntry(tipo){{
  const entries=[...(_c[tipo]||[])];entries.push({{nome:'',descricao:'',data:new Date().toISOString().split('T')[0]}});
  _c[tipo]=entries;await fbSave(tipo,entries);renderEntries(tipo);atualizarContadores();
  const lista=document.getElementById('lista-'+tipo);
  setTimeout(()=>{{lista.scrollTop=lista.scrollHeight;}},50);
  setTimeout(()=>{{const c=lista.querySelectorAll('.behavior-field');if(c.length)c[c.length-1].focus();}},60);
}}
async function removeEntry(tipo,idx){{
  const entries=[...(_c[tipo]||[])];entries.splice(idx,1);
  _c[tipo]=entries;await fbSave(tipo,entries);renderEntries(tipo);atualizarContadores();
}}
async function saveEntries(tipo){{
  const lista=document.getElementById('lista-'+tipo);const entries=[...(_c[tipo]||[])];
  lista.querySelectorAll('.behavior-entry').forEach(div=>{{
    const i=parseInt(div.dataset.idx);
    if(entries[i]){{entries[i].nome=div.querySelector('[data-field="nome"]')?.value||'';entries[i].descricao=div.querySelector('[data-field="descricao"]')?.value||'';entries[i].data=div.querySelector('[data-field="data"]')?.value||'';}}
  }});
  _c[tipo]=entries;await fbSave(tipo,entries);atualizarContadores();
}}
function listenFirebase(){{
  ['positivos','negativos','obs-positivo','obs-negativo'].forEach(key=>{{
    onSnapshot(doc(db,'dashboard',key),snap=>{{
      if(!snap.exists())return;const val=snap.data().value;_c[key]=val;
      if(key==='positivos'||key==='negativos'){{renderEntries(key);atualizarContadores();}}
      else{{const el=document.getElementById(key);if(el&&document.activeElement!==el)el.value=val||'';}}
    }});
  }});
  onSnapshot(collection(db,'dashboard'),snap=>{{
    snap.docChanges().forEach(ch=>{{
      const k=ch.doc.id;if(!k.startsWith('pa-'))return;
      const v=ch.doc.data().value||'';_c[k]=v;
      const el=document.querySelector(`.plano-input[data-key="${{k}}"]`);
      if(el&&document.activeElement!==el){{el.value=v;resize(el);}}
    }});atualizarContador();
  }});
}}
async function init(){{
  renderMarcas();
  for(const id of ['obs-positivo','obs-negativo']){{
    const el=document.getElementById(id);el.value=await fbLoad(id,'');
    let tm;el.addEventListener('input',()=>{{clearTimeout(tm);tm=setTimeout(()=>fbSave(id,el.value),600);}});
  }}
  _c['positivos']=await fbLoad('positivos',[]);if(!Array.isArray(_c['positivos']))_c['positivos']=[];
  _c['negativos']=await fbLoad('negativos',[]);if(!Array.isArray(_c['negativos']))_c['negativos']=[];
  renderEntries('positivos');renderEntries('negativos');atualizarContadores();
  listenFirebase();
}}
window.addEntry=addEntry;window.removeEntry=removeEntry;window.renderEntries=renderEntries;
init();
function ajustarAltura(){{
  const cl=document.querySelector('.col-left'),w=document.getElementById('tbl-marcas-wrap'),h=document.querySelector('#card-marcas .card-header');
  if(!cl||!w||!h)return;w.style.height=Math.max(120,cl.getBoundingClientRect().height-h.getBoundingClientRect().height)+'px';
}}
window.addEventListener('load',()=>setTimeout(ajustarAltura,150));
window.addEventListener('resize',ajustarAltura);
setTimeout(ajustarAltura,300);
</script>
</body>
</html>'''

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'✅ Dashboard gerado! Ref: {ref_str} · {total} auditorias · D-1={d1_str} · D-7 desde {d7_str} · D-30 desde {d30_str}')

if __name__ == '__main__':
    gerar()
