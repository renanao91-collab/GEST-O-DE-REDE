import os
import io
import gc
import math
import time
from datetime import datetime, date
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Panasonic do Brasil - Gestão de Rede & Pós-Venda",
    page_icon="🟦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# IDENTIDADE VISUAL PANASONIC DO BRASIL
# Cores Oficiais: Panasonic Blue (#004098), Dark Navy (#002B49), Accent Cyan (#00A3E0)
# -------------------------------------------------------------
PANASONIC_BLUE = "#004098"
PANASONIC_DARK = "#002B49"
PANASONIC_LIGHT_BG = "#F0F4F8"
PANASONIC_ACCENT = "#00A3E0"
META_10_DIAS = 62.0  # Meta estabelecida de 62% para atendimentos em até 10 dias

st.markdown(f"""
<style>
    .panasonic-header-container {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(135deg, {PANASONIC_BLUE} 0%, {PANASONIC_DARK} 100%);
        padding: 18px 25px;
        border-radius: 8px;
        margin-bottom: 20px;
        color: white;
        box-shadow: 0 4px 12px rgba(0, 64, 152, 0.15);
    }}
    .panasonic-title {{
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 24px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.5px;
    }}
    .panasonic-subtitle {{
        font-size: 13px;
        color: #D1E4FA;
        margin-top: 4px;
    }}
    .panasonic-logo-text {{
        font-family: 'Arial Black', Gadget, sans-serif;
        font-size: 26px;
        font-weight: 900;
        letter-spacing: 2px;
        color: #FFFFFF;
        background-color: rgba(255, 255, 255, 0.12);
        padding: 6px 16px;
        border-radius: 4px;
        border: 1px solid rgba(255, 255, 255, 0.25);
    }}
    div.stButton > button:first-child {{
        background-color: {PANASONIC_BLUE};
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: 600;
    }}
    div.stButton > button:first-child:hover {{
        background-color: {PANASONIC_DARK};
        color: white;
    }}
    [data-testid="stMetricValue"] {{
        color: {PANASONIC_BLUE} !important;
        font-weight: 700 !important;
    }}
    .alerta-meta-ofensor {{
        background-color: #FEE2E2;
        border-left: 5px solid #EF4444;
        padding: 12px 16px;
        border-radius: 6px;
        color: #991B1B;
        margin-bottom: 15px;
        font-weight: 500;
    }}
    .alerta-meta-atingida {{
        background-color: #ECFDF5;
        border-left: 5px solid #10B981;
        padding: 12px 16px;
        border-radius: 6px;
        color: #065F46;
        margin-bottom: 15px;
        font-weight: 500;
    }}
</style>
""", unsafe_allow_html=True)

# Topo com Identidade Panasonic
st.markdown(f"""
<div class="panasonic-header-container">
    <div>
        <div class="panasonic-title">Panasonic do Brasil | Gestão de Rede & Pós-Venda</div>
        <div class="panasonic-subtitle">Acompanhamento Operacional, Controle de SLA (Meta 10d ≥ 62%), Prazo CDC (30d), Diagnóstico de Postos Ofensores e Gestão por Consultoras</div>
    </div>
    <div class="panasonic-logo-text">Panasonic</div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# ÁREA DE UPLOAD (2 ARQUIVOS)
# -------------------------------------------------------------
col_up1, col_up2 = st.columns(2)

with col_up1:
    st.subheader("1. Relatório de Ordens de Serviço")
    file_servico = st.file_uploader(
        "Selecione o relatório de chamados da rede (CSV ou XLSX):",
        type=["csv", "xlsx"],
        key="servico"
    )

with col_up2:
    st.subheader("2. Base de Consultoras Panasonic (De-Para)")
    file_consultoras = st.file_uploader(
        "Selecione a planilha com Posto Autorizado x Consultora Responsável (XLSX ou CSV):",
        type=["xlsx", "csv"],
        key="consultoras"
    )
    with st.expander("ℹ️ Instrução sobre a planilha de Consultoras"):
        st.write("""
        A planilha deve conter 2 colunas simples:
        - **Nome da Autorizada / Posto:** ex. *Unidade*, *Autorizada*, *Posto*, ou *Unidade (Digiteam)*
        - **Consultora Responsável:** ex. *Consultora*, *Responsável*, *Gestora*
        """)

# -------------------------------------------------------------
# FUNÇÕES OTIMIZADAS COM CACHE DE MEMÓRIA
# -------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=2)
def carregar_e_processar_dados(bytes_servico, nome_servico, bytes_cons, nome_cons, data_ref_str):
    buf_s = io.BytesIO(bytes_servico)
    if nome_servico.lower().endswith('.csv'):
        try:
            df_raw = pd.read_csv(buf_s, low_memory=False)
        except Exception:
            buf_s.seek(0)
            df_raw = pd.read_csv(buf_s, sep=';', low_memory=False)
    else:
        df_raw = pd.read_excel(buf_s)

    df_raw.columns = [str(c).replace('\ufeff', '').replace('"', '').strip() for c in df_raw.columns]

    # Tratamento de datas (Vetorizado - ultra rápido)
    df_raw['dt_criacao'] = pd.to_datetime(df_raw.get('Data de Criação', pd.Series(dtype='object')), errors='coerce')
    df_raw['dt_conclusao'] = pd.to_datetime(df_raw.get('Data de Conclusão', pd.Series(dtype='object')), errors='coerce')
    df_raw['dt_finalizacao'] = pd.to_datetime(df_raw.get('Data de Finalização', pd.Series(dtype='object')), errors='coerce')

    # Ordenar do mais antigo para o mais novo
    df_sorted = df_raw.sort_values(by='dt_criacao', ascending=True, na_position='last').reset_index(drop=True)
    del df_raw
    gc.collect()

    df_sorted['Unidade (Digiteam)'] = df_sorted.get('Unidade (Digiteam)', pd.Series(['Não Informada']*len(df_sorted))).fillna('Não Informada').astype(str).str.strip()

    # Mapeamento de Consultoras
    if bytes_cons is not None and nome_cons is not None:
        buf_c = io.BytesIO(bytes_cons)
        if nome_cons.lower().endswith('.csv'):
            try:
                df_cons = pd.read_csv(buf_c, low_memory=False)
            except Exception:
                buf_c.seek(0)
                df_cons = pd.read_csv(buf_c, sep=';', low_memory=False)
        else:
            df_cons = pd.read_excel(buf_c)

        df_cons.columns = [str(c).replace('\ufeff', '').replace('"', '').strip() for c in df_cons.columns]

        col_unid_cons = next((c for c in df_cons.columns if any(k in c.lower() for k in ['unidade', 'autorizada', 'posto', 'nome', 'razão', 'credenciada'])), df_cons.columns[0])
        col_nome_cons = next((c for c in df_cons.columns if any(k in c.lower() for k in ['consultora', 'responsável', 'responsavel', 'gestora', 'atendente', 'consultor'])), df_cons.columns[1] if len(df_cons.columns) > 1 else df_cons.columns[0])

        map_dict = dict(zip(
            df_cons[col_unid_cons].astype(str).str.strip().str.upper(),
            df_cons[col_nome_cons].astype(str).str.strip()
        ))
        df_sorted['Consultora Responsável'] = df_sorted['Unidade (Digiteam)'].str.upper().map(map_dict).fillna('Não Atribuída')
    else:
        df_sorted['Consultora Responsável'] = 'Não Atribuída (Insira a planilha 2)'

    # Cálculo Vetorizado do TAT
    ts_ref = pd.Timestamp(data_ref_str).normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
    dt_fim = df_sorted['dt_conclusao'].fillna(df_sorted['dt_finalizacao']).fillna(ts_ref)
    
    diff_secs = (dt_fim - df_sorted['dt_criacao']).dt.total_seconds()
    tat_dias = np.floor(diff_secs / 86400.0).fillna(0).clip(lower=0).astype(int)
    tat_dias[df_sorted['dt_criacao'].isna()] = 0
    df_sorted['TAT_Dias'] = tat_dias

    # Flags vetorizadas
    st_upper = df_sorted.get('Status Atual', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip().str.upper()
    df_sorted['Status Atual'] = st_upper
    df_sorted['Is_Reparado'] = st_upper.isin(['CONCLUÍDA', 'FINALIZADA'])
    df_sorted['Is_Aberta'] = ~st_upper.isin(['CONCLUÍDA', 'FINALIZADA', 'CANCELADA'])

    # Faixas de TAT vetorizadas
    df_sorted['Faixa_TAT'] = np.select(
        [df_sorted['TAT_Dias'] <= 10, df_sorted['TAT_Dias'] <= 30],
        ["Até 10 dias (No Target)", "11 a 30 dias (Alerta)"],
        default="> 30 dias (Estouro CDC)"
    )

    # Reincidência vetorizada
    dem_str = df_sorted.get('Tipo de Demanda', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
    df_sorted['Tipo de Demanda'] = dem_str
    df_sorted['Is_Reincidente'] = dem_str.str.upper().str.contains('REINCIDENCIA|REINCIDÊNCIA', regex=True)

    # Peça vetorizada
    c_pl = df_sorted.get('Código Peça Lançada', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
    c_pt = df_sorted.get('Codigo Troca de Peças', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
    df_sorted['Tem_Peca'] = (c_pl != '') | (c_pt != '')
    df_sorted['Tem_Peca_Txt'] = np.where(df_sorted['Tem_Peca'], 'SIM', 'NÃO')

    # Prioridade vetorizada
    cond_baixa = (st_upper == 'CANCELADA')
    cond_critica = (df_sorted['Is_Reincidente']) | (df_sorted['TAT_Dias'] > 30)
    cond_alta = ((df_sorted['Tem_Peca']) | (df_sorted['TAT_Dias'] >= 3)) & (~df_sorted['Is_Reparado'])
    cond_media = (st_upper == 'CRIADA')

    df_sorted['Prioridade'] = np.select(
        [cond_baixa, cond_critica, cond_alta, cond_media],
        ['BAIXA', 'CRÍTICA', 'ALTA', 'MÉDIA'],
        default='NORMAL'
    )

    if 'Estado' in df_sorted.columns:
        df_sorted['Estado'] = df_sorted['Estado'].fillna('').astype(str).str.strip().str.upper()
    else:
        df_sorted['Estado'] = ''

    return df_sorted

def calcular_ofensores(df_base):
    if len(df_base) == 0:
        return pd.DataFrame(columns=['Unidade (Digiteam)', 'Consultora Responsável', 'Total_OS', 'OS_Ate_10d', 'TAT_Medio', 'Pct_10d', 'Status_Meta', 'Gap_OS_Meta'])
    
    df_val = df_base[df_base['Status Atual'] != 'CANCELADA']
    if len(df_val) == 0:
        return pd.DataFrame(columns=['Unidade (Digiteam)', 'Consultora Responsável', 'Total_OS', 'OS_Ate_10d', 'TAT_Medio', 'Pct_10d', 'Status_Meta', 'Gap_OS_Meta'])
    
    ate_10d_series = (df_val['TAT_Dias'] <= 10).astype(int)
    
    resumo = df_val.assign(Ate_10d_Num=ate_10d_series).groupby(
        ['Unidade (Digiteam)', 'Consultora Responsável'], as_index=False
    ).agg(
        Total_OS=('Código OS', 'count'),
        OS_Ate_10d=('Ate_10d_Num', 'sum'),
        TAT_Medio=('TAT_Dias', 'mean')
    )
    resumo['Pct_10d'] = (resumo['OS_Ate_10d'] / resumo['Total_OS']) * 100.0
    resumo['Status_Meta'] = np.where(resumo['Pct_10d'] >= META_10_DIAS, "✅ Na Meta (≥ 62%)", "❌ Ofensor (< 62%)")
    resumo['Gap_OS_Meta'] = np.maximum(0, np.ceil((META_10_DIAS / 100.0) * resumo['Total_OS']) - resumo['OS_Ate_10d']).astype(int)
    
    df_of = resumo[resumo['Pct_10d'] < META_10_DIAS].copy()
    df_of = df_of.sort_values(by=['Gap_OS_Meta', 'Pct_10d', 'Total_OS'], ascending=[False, True, False]).reset_index(drop=True)
    return df_of

def gerar_excel_completo(df_export, df_ofensores_exp, data_corte_str, nome_aba_base="Controle Operacional OS"):
    wb = openpyxl.Workbook(write_only=False)
    wb.remove(wb.active)

    ws_dash = wb.create_sheet(title="Dashboard Gestão")
    ws_ofensores = wb.create_sheet(title="Postos Ofensores 10D")
    ws_base = wb.create_sheet(title=nome_aba_base)
    ws_pecas = wb.create_sheet(title="Controle de Peças")

    COLOR_PANASONIC_EXCEL = "004098"
    WHITE = "FFFFFF"
    BORDER_GRAY = "D1D5DB"
    
    font_title = Font(name="Segoe UI", size=15, bold=True, color=COLOR_PANASONIC_EXCEL)
    font_sub = Font(name="Segoe UI", size=9, italic=True, color="4B5563")
    font_header = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
    font_bold = Font(name="Segoe UI", size=9, bold=True)
    fill_header = PatternFill(start_color=COLOR_PANASONIC_EXCEL, end_color=COLOR_PANASONIC_EXCEL, fill_type="solid")
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    thin_border = Side(border_style="thin", color=BORDER_GRAY)
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    n_rows = len(df_export)
    n_fim = str(max(5, n_rows + 4))
    aba = str(nome_aba_base)

    # 1. ABA DE POSTOS OFENSORES
    ws_ofensores["A1"] = "PANASONIC DO BRASIL - POSTOS OFENSORES DO INDICADOR DE 10 DIAS"
    ws_ofensores["A1"].font = font_title
    ws_ofensores["A2"] = f"Meta Estabelecida: ≥ {META_10_DIAS:.0f}% de atendimentos em até 10 dias corridos | Data de Corte: {data_corte_str}"
    ws_ofensores["A2"].font = font_sub

    headers_of = ["Posto Autorizado (Unidade)", "Consultora Responsável", "Total de OS Válidas", "OS em até 10 Dias", "% Atingido", "Meta Target", "Status Meta", "Déficit de OS p/ Meta"]
    for c_idx, h in enumerate(headers_of, 1):
        c = ws_ofensores.cell(row=4, column=c_idx, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center")
    ws_ofensores.row_dimensions[4].height = 26

    for r_idx, r_data in enumerate(df_ofensores_exp.itertuples(), start=5):
        r_str = str(r_idx)
        posto_nome = str(r_data[1])
        cons_nome = str(r_data[2])

        ws_ofensores.cell(row=r_idx, column=1, value=posto_nome).border = border_cell
        ws_ofensores.cell(row=r_idx, column=2, value=cons_nome).border = border_cell
        
        f_tot_val = '=COUNTIFS(\'' + aba + '\'!$N$5:$N$' + n_fim + ', A' + r_str + ', \'' + aba + '\'!$I$5:$I$' + n_fim + ', "<>CANCELADA")'
        c3 = ws_ofensores.cell(row=r_idx, column=3, value=f_tot_val)
        c3.number_format = "#,##0"; c3.alignment = Alignment(horizontal="right"); c3.border = border_cell
        
        f_ate_10 = '=COUNTIFS(\'' + aba + '\'!$N$5:$N$' + n_fim + ', A' + r_str + ', \'' + aba + '\'!$I$5:$I$' + n_fim + ', "<>CANCELADA", \'' + aba + '\'!$D$5:$D$' + n_fim + ', "<=10")'
        c4 = ws_ofensores.cell(row=r_idx, column=4, value=f_ate_10)
        c4.number_format = "#,##0"; c4.alignment = Alignment(horizontal="right"); c4.border = border_cell
        
        c5 = ws_ofensores.cell(row=r_idx, column=5, value='=IF(C' + r_str + '>0, D' + r_str + '/C' + r_str + ', 0)')
        c5.number_format = "0.0%"; c5.alignment = Alignment(horizontal="right"); c5.fill = fill_red; c5.font = font_bold; c5.border = border_cell
        
        c6 = ws_ofensores.cell(row=r_idx, column=6, value=META_10_DIAS / 100.0)
        c6.number_format = "0.0%"; c6.alignment = Alignment(horizontal="right"); c6.border = border_cell
        
        c7 = ws_ofensores.cell(row=r_idx, column=7, value='=IF(E' + r_str + '>=F' + r_str + ', "NA META (≥ 62%)", "OFENSOR (< 62%)")')
        c7.alignment = Alignment(horizontal="center"); c7.border = border_cell
        
        c8 = ws_ofensores.cell(row=r_idx, column=8, value='=MAX(0, ROUNDUP(C' + r_str + '*F' + r_str + ', 0)-D' + r_str + ')')
        c8.number_format = "#,##0"; c8.alignment = Alignment(horizontal="right"); c8.border = border_cell

    ws_ofensores.freeze_panes = "C5"
    ws_ofensores.auto_filter.ref = f"A4:H{max(5, len(df_ofensores_exp)+4)}"
    for col_l, w in {"A": 42, "B": 24, "C": 18, "D": 18, "E": 14, "F": 14, "G": 18, "H": 22}.items():
        ws_ofensores.column_dimensions[col_l].width = w

    # 2. ABA BASE
    ws_base["A1"] = "PANASONIC DO BRASIL - GESTÃO DE REDE & PÓS-VENDA"
    ws_base["A1"].font = font_title
    ws_base["A2"] = f"Acompanhamento Operacional | Data de Corte: {data_corte_str} | Ordenado do mais antigo para o mais novo"
    ws_base["A2"].font = font_sub

    headers_base = [
        "Consultora Responsável", "Prioridade Operacional", "Faixa SLA / CDC", "TAT Atual (Dias)", "Tem Peça Lançada?",
        "Código da OS", "Ticket", "Tipo de Demanda", "Status Atual", "Data de Criação",
        "Data de Conclusão", "Categoria do Produto", "Modelo do Produto", "Unidade Técnica (Rede)",
        "Região / Cidade", "UF", "Técnico de Campo", "Código Peça Lançada", "Descrição Peça Lançada",
        "Código Peça Trocada", "Descrição Peça Trocada", "Sintoma / Defeito Relatado", "Ação Operacional Recomendada"
    ]
    for c_idx, h in enumerate(headers_base, 1):
        c = ws_base.cell(row=4, column=c_idx, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_base.row_dimensions[4].height = 26

    cod_os = df_export['Código OS'].astype(str).tolist() if 'Código OS' in df_export else ['']*n_rows
    ticket = df_export['Número do Ticket'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).tolist() if 'Número do Ticket' in df_export else ['']*n_rows
    demanda = df_export['Tipo de Demanda'].fillna('').astype(str).tolist() if 'Tipo de Demanda' in df_export else ['']*n_rows
    status = df_export['Status Atual'].fillna('').astype(str).tolist() if 'Status Atual' in df_export else ['']*n_rows
    dt_criac_str = df_export['dt_criacao'].dt.strftime('%Y-%m-%d %H:%M').fillna('').tolist()
    dt_concl_comb = df_export['dt_conclusao'].fillna(df_export['dt_finalizacao'])
    dt_concl_str = dt_concl_comb.dt.strftime('%Y-%m-%d %H:%M').fillna('').tolist()
    cat = df_export['Categoria do Produto'].fillna('').astype(str).tolist() if 'Categoria do Produto' in df_export else ['']*n_rows
    mod = df_export['Modelo do Produto'].fillna('').astype(str).tolist() if 'Modelo do Produto' in df_export else ['']*n_rows
    unid = df_export['Unidade (Digiteam)'].fillna('').astype(str).tolist()
    cid = df_export['Cidade'].fillna('').astype(str).tolist() if 'Cidade' in df_export else ['']*n_rows
    uf = df_export['Estado'].fillna('').astype(str).tolist()
    tec = df_export['Agente de Campo'].fillna('').astype(str).tolist() if 'Agente de Campo' in df_export else ['']*n_rows
    c_pl = df_export['Código Peça Lançada'].fillna('').astype(str).str.strip().tolist() if 'Código Peça Lançada' in df_export else ['']*n_rows
    d_pl = df_export['Peça Lançada'].fillna('').astype(str).str.strip().tolist() if 'Peça Lançada' in df_export else ['']*n_rows
    c_pt = df_export['Codigo Troca de Peças'].fillna('').astype(str).str.strip().tolist() if 'Codigo Troca de Peças' in df_export else ['']*n_rows
    d_pt = df_export['Troca de Peças'].fillna('').astype(str).str.strip().tolist() if 'Troca de Peças' in df_export else ['']*n_rows
    sint = df_export['Sintoma e Defeito'].fillna('').astype(str).str.strip().tolist() if 'Sintoma e Defeito' in df_export else ['']*n_rows
    cons_resp = df_export['Consultora Responsável'].tolist()

    for i in range(n_rows):
        r_str = str(i + 5)
        dc = dt_criac_str[i]
        f_tat = f'=IF(K{r_str}="", INT(DATEVALUE("{data_corte_str}")-DATEVALUE(LEFT(J{r_str},10))), INT(DATEVALUE(LEFT(K{r_str},10))-DATEVALUE(LEFT(J{r_str},10))))' if dc else 0
        f_faixa = f'=IF(D{r_str}<=10, "Até 10 dias (No Target)", IF(D{r_str}<=30, "11 a 30 dias (Alerta)", "> 30 dias (Estouro CDC)"))'
        f_peca = f'=IF(OR(R{r_str}<>"", T{r_str}<>""), "SIM", "NÃO")'
        f_prio = f'=IF(I{r_str}="CANCELADA", "BAIXA", IF(OR(ISNUMBER(SEARCH("REINCIDENCIA", H{r_str})), ISNUMBER(SEARCH("REINCIDÊNCIA", H{r_str})), D{r_str}>30), "CRÍTICA", IF(AND(E{r_str}="SIM", I{r_str}<>"FINALIZADA", I{r_str}<>"CONCLUÍDA"), "ALTA", IF(AND(D{r_str}>=3, I{r_str}<>"FINALIZADA", I{r_str}<>"CONCLUÍDA"), "ALTA", IF(I{r_str}="CRIADA", "MÉDIA", "NORMAL")))))'
        f_acao = f'=IF(OR(I{r_str}="CONCLUÍDA", I{r_str}="FINALIZADA"), "Encerrado com Sucesso", IF(I{r_str}="CANCELADA", "Verificar Motivo Cancelamento", IF(D{r_str}>30, "AÇÃO IMEDIATA: Estouro de Prazo CDC (Risco Jurídico/Troca)", IF(D{r_str}>10, "ATENÇÃO: Fora da Meta de 10 Dias - Priorizar Atendimento", IF(E{r_str}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(I{r_str}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(I{r_str}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento")))))))'

        ws_base.append([
            cons_resp[i], f_prio, f_faixa, f_tat, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
            cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
        ])

    ws_base.freeze_panes = "F5"
    ws_base.auto_filter.ref = f"A4:W{max(5, n_rows+4)}"
    for col_l, w in {"A": 22, "B": 20, "C": 24, "D": 14, "E": 16, "F": 18, "G": 14, "H": 28, "I": 18, "J": 18, "K": 18, "L": 22, "M": 18, "N": 38, "O": 20, "P": 8, "Q": 25, "R": 24, "S": 35, "T": 24, "U": 35, "V": 30, "W": 44}.items():
        ws_base.column_dimensions[col_l].width = w

    # 3. ABA PEÇAS
    ws_pecas["A1"] = "PANASONIC DO BRASIL - CONTROLE ESPECÍFICO DE PEÇAS"
    ws_pecas["A1"].font = font_title
    headers_p = ["Consultora Responsável", "Código da OS", "Ticket", "Status Atual", "Tipo de Demanda", "Unidade Técnica (Rede)", "Cidade / UF", "Código da Peça", "Descrição da Peça", "Tipo de Registro", "Técnico de Campo", "Status Logístico Sugerido"]
    for c_idx, h in enumerate(headers_p, 1):
        c = ws_pecas.cell(row=4, column=c_idx, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center")

    count_p = 0
    for i in range(n_rows):
        if c_pl[i] or d_pl[i] or c_pt[i] or d_pt[i]:
            st_val = status[i]
            cid_uf = f"{cid[i]} / {uf[i]}"
            c_resp = cons_resp[i]
            if c_pl[i] or d_pl[i]:
                st_log = "Instalada / Baixada" if st_val in ["CONCLUÍDA", "FINALIZADA"] else "Pendente Despacho / Chegada na Unidade"
                ws_pecas.append([c_resp, cod_os[i], ticket[i], st_val, demanda[i], unid[i], cid_uf, c_pl[i], d_pl[i], "Peça Lançada (Aguardando / Em Processamento)", tec[i], st_log])
                count_p += 1
            if c_pt[i] or d_pt[i]:
                st_log = "Peça Trocada - Iniciar Logística Reversa" if st_val in ["CONCLUÍDA", "FINALIZADA"] else "Em Processo de Troca em Campo"
                ws_pecas.append([c_resp, cod_os[i], ticket[i], st_val, demanda[i], unid[i], cid_uf, c_pt[i], d_pt[i], "Troca de Peça Efetuada em Campo", tec[i], st_log])
                count_p += 1

    ws_pecas.auto_filter.ref = f"A4:L{max(5, count_p+4)}"
    ws_pecas.freeze_panes = "D5"

    # 4. ABA DASHBOARD
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash["A1"] = "PANASONIC DO BRASIL - PAINEL DE CONTROLE E GESTÃO DA REDE"; ws_dash["A1"].font = font_title
    ws_dash["A2"] = f"Meta de 10 Dias: ≥ {META_10_DIAS:.0f}% | Data de Corte: {data_corte_str}"; ws_dash["A2"].font = font_sub

    def add_kpi(ws, col, title, form, fmt="#,##0"):
        ws.cell(row=4, column=col, value=title).font = Font(name="Segoe UI", size=9, bold=True, color="6B7280")
        c = ws.cell(row=5, column=col, value=form)
        c.font = Font(name="Segoe UI", size=18, bold=True, color=COLOR_PANASONIC_EXCEL); c.number_format = fmt
        for r in range(4, 7):
            for cell_c in range(col, col + 2):
                cell = ws.cell(row=r, column=cell_c)
                cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
                top_s = Side(border_style="medium", color=COLOR_PANASONIC_EXCEL) if r == 4 else Side(border_style="thin", color="CBD5E1")
                cell.border = Border(left=Side(border_style="thin", color="CBD5E1"),
                                     right=Side(border_style="thin", color="CBD5E1"),
                                     top=top_s,
                                     bottom=Side(border_style="thin", color="CBD5E1"))
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
        ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col+1)

    f_valid_denom = 'MAX(1, COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "<>CANCELADA"))'
    f_kpi_10d = '=COUNTIFS(\'' + aba + '\'!D5:D' + n_fim + ', "<=10", \'' + aba + '\'!I5:I' + n_fim + ', "<>CANCELADA") / ' + f_valid_denom
    f_kpi_30d = '=COUNTIFS(\'' + aba + '\'!D5:D' + n_fim + ', "<=30", \'' + aba + '\'!I5:I' + n_fim + ', "<>CANCELADA") / ' + f_valid_denom
    f_kpi_reparo = '=(SUMIFS(\'' + aba + '\'!D5:D' + n_fim + ', \'' + aba + '\'!I5:I' + n_fim + ', "CONCLUÍDA") + SUMIFS(\'' + aba + '\'!D5:D' + n_fim + ', \'' + aba + '\'!I5:I' + n_fim + ', "FINALIZADA")) / MAX(1, COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "CONCLUÍDA") + COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "FINALIZADA"))'
    f_kpi_reinc = '=(COUNTIFS(\'' + aba + '\'!H5:H' + n_fim + ', "*REINCIDENCIA*", \'' + aba + '\'!I5:I' + n_fim + ', "<>CANCELADA") + COUNTIFS(\'' + aba + '\'!H5:H' + n_fim + ', "*REINCIDÊNCIA*", \'' + aba + '\'!I5:I' + n_fim + ', "<>CANCELADA")) / ' + f_valid_denom
    f_kpi_total = '=COUNTA(\'' + aba + '\'!F5:F' + n_fim + ')'
    f_kpi_aberto = '=COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "<>CONCLUÍDA") - COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "FINALIZADA") - COUNTIF(\'' + aba + '\'!I5:I' + n_fim + ', "CANCELADA")'

    add_kpi(ws_dash, 1, "% ATÉ 10 DIAS (META ≥62%)", f_kpi_10d, "0.0%")
    add_kpi(ws_dash, 3, "% ATÉ 30 DIAS (CDC)", f_kpi_30d, "0.0%")
    add_kpi(ws_dash, 5, "TAT MÉDIO REPARO", f_kpi_reparo, "0.0")
    add_kpi(ws_dash, 7, "% REINCIDENTES", f_kpi_reinc, "0.00%")
    add_kpi(ws_dash, 9, "TOTAL DE ORDENS", f_kpi_total, "#,##0")
    add_kpi(ws_dash, 11, "ORDENS EM ABERTO", f_kpi_aberto, "#,##0")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# -------------------------------------------------------------
# CONFIGURAÇÃO DE DATA DE CONSULTA NA BARRA LATERAL
# -------------------------------------------------------------
st.sidebar.markdown(f"<h3 style='color:{PANASONIC_BLUE};'>📅 Referência da Consulta</h3>", unsafe_allow_html=True)
data_referencia = st.sidebar.date_input(
    "Data de Corte / Consulta:",
    value=date.today(),
    help="Para ordens ainda NÃO finalizadas (sem data de conclusão), o tempo decorrido de atendimento é calculado com base nesta data."
)
data_corte_iso = data_referencia.strftime('%Y-%m-%d')

# -------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# -------------------------------------------------------------
if file_servico is not None:
    bytes_s = file_servico.getvalue()
    nome_s = file_servico.name
    bytes_c = file_consultoras.getvalue() if file_consultoras is not None else None
    nome_c = file_consultoras.name if file_consultoras is not None else None

    with st.spinner("Processando base da rede Panasonic..."):
        df_sorted = carregar_e_processar_dados(bytes_s, nome_s, bytes_c, nome_c, data_corte_iso)

    # -------------------------------------------------------------
    # BARRA LATERAL: FILTROS OPERACIONAIS
    # -------------------------------------------------------------
    st.sidebar.markdown(f"<h3 style='color:{PANASONIC_BLUE};'>🎯 Filtros Operacionais</h3>", unsafe_allow_html=True)
    
    lista_consultoras = sorted(df_sorted['Consultora Responsável'].dropna().astype(str).unique().tolist())
    sel_consultoras = st.sidebar.multiselect(
        "👩‍💼 Consultora(s) Responsável(is):",
        options=lista_consultoras,
        default=[],
        help="Deixe em branco para visualizar todas, ou selecione uma ou mais consultoras ao mesmo tempo."
    )

    lista_demandas = sorted([x for x in df_sorted['Tipo de Demanda'].dropna().astype(str).unique().tolist() if x != ''])
    sel_demanda = st.sidebar.multiselect("📋 Tipo de Demanda:", lista_demandas, default=[])

    faixas_opcoes = ["Até 10 dias (No Target)", "11 a 30 dias (Alerta)", "> 30 dias (Estouro CDC)"]
    sel_faixas = st.sidebar.multiselect("⏱️ Faixa de Dias (SLA / CDC):", faixas_opcoes, default=[])

    lista_status = sorted([x for x in df_sorted['Status Atual'].dropna().astype(str).unique().tolist() if x != ''])
    sel_status = st.sidebar.multiselect("🚦 Status Operacional:", lista_status, default=[])

    lista_prio = ["CRÍTICA", "ALTA", "MÉDIA", "NORMAL", "BAIXA"]
    sel_prio = st.sidebar.multiselect("⚡ Nível de Prioridade:", lista_prio, default=[])

    ufs = sorted([x for x in df_sorted['Estado'].dropna().astype(str).unique().tolist() if x != ''])
    sel_uf = st.sidebar.multiselect("📍 Estado (UF):", ufs, default=[])

    # Filtragem em memória ultra rápida usando máscaras booleanas
    mask_final = pd.Series(True, index=df_sorted.index)
    if sel_consultoras:
        mask_final &= df_sorted['Consultora Responsável'].isin(sel_consultoras)
    if sel_demanda:
        mask_final &= df_sorted['Tipo de Demanda'].isin(sel_demanda)
    if sel_faixas:
        mask_final &= df_sorted['Faixa_TAT'].isin(sel_faixas)
    if sel_status:
        mask_final &= df_sorted['Status Atual'].isin(sel_status)
    if sel_prio:
        mask_final &= df_sorted['Prioridade'].isin(sel_prio)
    if sel_uf:
        mask_final &= df_sorted['Estado'].isin(sel_uf)

    df_filtrado = df_sorted[mask_final]

    # -------------------------------------------------------------
    # CÁLCULO DOS INDICADORES SOLICITADOS
    # -------------------------------------------------------------
    total_os = len(df_filtrado)
    
    df_validas = df_filtrado[df_filtrado['Status Atual'] != 'CANCELADA']
    total_validas = len(df_validas)
    
    qtd_ate_10 = int((df_validas['TAT_Dias'] <= 10).sum()) if total_validas > 0 else 0
    pct_ate_10 = (qtd_ate_10 / total_validas * 100.0) if total_validas > 0 else 0.0
    delta_meta_10 = pct_ate_10 - META_10_DIAS
    
    qtd_ate_30 = int((df_validas['TAT_Dias'] <= 30).sum()) if total_validas > 0 else 0
    pct_ate_30 = (qtd_ate_30 / total_validas * 100.0) if total_validas > 0 else 0.0
    
    df_reparadas = df_filtrado[df_filtrado['Is_Reparado']]
    tat_medio_reparo = float(df_reparadas['TAT_Dias'].mean()) if len(df_reparadas) > 0 else 0.0
    
    qtd_reinc = int(df_validas['Is_Reincidente'].sum()) if total_validas > 0 else 0
    pct_reinc = (qtd_reinc / total_validas * 100.0) if total_validas > 0 else 0.0

    if sel_consultoras:
        if len(sel_consultoras) == 1:
            nome_visao = f"Consultora: {sel_consultoras[0]}"
        else:
            nome_visao = f"{len(sel_consultoras)} Consultoras Selecionadas ({', '.join(sel_consultoras[:2])}...)"
    else:
        nome_visao = "Toda a Rede Panasonic"

    st.markdown(f"### 📌 Painel de Indicadores Estratégicos: **{nome_visao}**")
    st.caption(f"Cálculo com data de referência: **{data_referencia.strftime('%d/%m/%Y')}** (Ordens em aberto calculadas até esta data)")

    if total_os == 0:
        st.warning("⚠️ Nenhuma ordem de serviço encontrada para a combinação de filtros selecionada.")
    else:
        if pct_ate_10 >= META_10_DIAS:
            st.markdown(f"""
            <div class="alerta-meta-atingida">
                ✅ <b>META ATINGIDA:</b> O índice atual de atendimentos em até 10 dias é de <b>{pct_ate_10:.1f}%</b>, superando a meta estabelecida de <b>{META_10_DIAS:.1f}%</b> (+{delta_meta_10:.1f}% vs Target).
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="alerta-meta-ofensor">
                ⚠️ <b>ABAIXO DA META:</b> O índice atual de atendimentos em até 10 dias é de <b>{pct_ate_10:.1f}%</b>, ficando abaixo da meta estipulada de <b>{META_10_DIAS:.1f}%</b> ({delta_meta_10:.1f}% vs Target).
            </div>
            """, unsafe_allow_html=True)

        # Linha 1: Os 4 Indicadores Estratégicos
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric(
            "⚡ Atendimentos ≤ 10 Dias",
            f"{pct_ate_10:.1f}%",
            f"{delta_meta_10:+.1f}% vs Meta (62%)",
            help=f"Meta: ≥ {META_10_DIAS:.0f}%. Considera concluídas com TAT ≤ 10d e ordens em aberto com até 10d na data de corte."
        )
        kpi2.metric(
            "⚖️ Atendimentos ≤ 30 Dias (CDC)",
            f"{pct_ate_30:.1f}%",
            f"{qtd_ate_30:,} de {total_validas:,} ordens válidas".replace(",", "."),
            help="Conformidade com o Artigo 18 do CDC (limite de 30 dias corridos)."
        )
        kpi3.metric(
            "🔧 TAT Médio de Reparo",
            f"{tat_medio_reparo:.1f} dias",
            f"{len(df_reparadas):,} ordens concluídas".replace(",", "."),
            help="Tempo médio em dias para reparo efetivo das ordens concluídas/finalizadas."
        )
        kpi4.metric(
            "🔄 Taxa de Reincidência",
            f"{pct_reinc:.2f}%",
            f"{qtd_reinc:,} ordens reincidentes".replace(",", "."),
            delta_color="inverse",
            help="Percentual de atendimentos identificados como reincidência técnica na rede."
        )

        # Linha 2: Métricas de Volume
        c_vol1, c_vol2, c_vol3, c_vol4 = st.columns(4)
        em_aberto = int(df_filtrado['Is_Aberta'].sum())
        com_pecas = int(df_filtrado['Tem_Peca'].sum())
        criticas = int((df_filtrado['Prioridade'] == 'CRÍTICA').sum())

        c_vol1.metric("Total de Ordens na Visão", f"{total_os:,}".replace(",", "."))
        c_vol2.metric("Ordens em Aberto", f"{em_aberto:,}".replace(",", "."))
        c_vol3.metric("Ordens com Peças Lançadas", f"{com_pecas:,}".replace(",", "."))
        c_vol4.metric("Ordens Críticas / Estouro", f"{criticas:,}".replace(",", "."))

        st.divider()

        # -------------------------------------------------------------
        # SEÇÃO: DIAGNÓSTICO DOS POSTOS OFENSORES DA META (10 DIAS < 62%)
        # -------------------------------------------------------------
        st.subheader("🚨 Diagnóstico de Postos Ofensores (Atendimento em até 10 Dias < 62%)")
        st.markdown("""
        Esta análise identifica os postos autorizados que estão **puxando o indicador para baixo**. 
        O cálculo considera o total de ordens válidas de cada posto e quanto falta para atingirem a meta de **62%**.
        """)

        df_ofensores_filtrado = calcular_ofensores(df_filtrado)
        
        c_filt_of1, c_filt_of2, c_filt_of3 = st.columns([2, 2, 1])
        with c_filt_of1:
            max_os_posto = max(2, int(df_ofensores_filtrado['Total_OS'].max())) if len(df_ofensores_filtrado) > 0 else 2
            min_ordens = st.slider(
                "Filtrar postos com no mínimo de ordens:",
                min_value=1,
                max_value=max_os_posto,
                value=1,
                help="Use 1 para listar TODOS os postos ofensores ou aumente para focar nos de maior volume."
            )
        with c_filt_of2:
            busca_posto = st.text_input("🔎 Buscar Posto Ofensor específico (ex: DUSO):", "")

        df_ofensores_view = df_ofensores_filtrado[df_ofensores_filtrado['Total_OS'] >= min_ordens].copy()
        if busca_posto:
            df_ofensores_view = df_ofensores_view[
                df_ofensores_view['Unidade (Digiteam)'].str.contains(busca_posto, case=False, na=False) |
                df_ofensores_view['Consultora Responsável'].str.contains(busca_posto, case=False, na=False)
            ]

        with c_filt_of3:
            st.metric(
                "Postos Ofensores",
                f"{len(df_ofensores_view)} postos",
                f"{len(df_ofensores_filtrado)} na seleção",
                delta_color="inverse"
            )

        if len(df_ofensores_view) > 0:
            col_g_of1, col_g_of2 = st.columns([3, 2])
            with col_g_of1:
                top_15_ofensores = df_ofensores_view.sort_values(by=['Gap_OS_Meta', 'Pct_10d'], ascending=[False, True]).head(15).copy()
                fig_of = px.bar(
                    top_15_ofensores,
                    x='Pct_10d',
                    y='Unidade (Digiteam)',
                    orientation='h',
                    text='Pct_10d',
                    title="Top 15 Maiores Ofensores por Impacto (Déficit de OS p/ Meta)",
                    color='Pct_10d',
                    color_continuous_scale='Reds_r'
                )
                fig_of.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                fig_of.add_vline(
                    x=META_10_DIAS,
                    line_dash="dash",
                    line_color="#004098",
                    annotation_text="Meta 62%",
                    annotation_position="top right"
                )
                fig_of.update_layout(
                    yaxis={'categoryorder':'total ascending'},
                    height=450,
                    xaxis_title="% de Atendimentos em até 10 Dias",
                    yaxis_title="Posto Autorizado",
                    margin=dict(l=20, r=20, t=40, b=20)
                )
                st.plotly_chart(fig_of, use_container_width=True)

            with col_g_of2:
                st.markdown("#### 📋 Tabela Detalhada de Postos Fora da Meta")
                tabela_ofensores = df_ofensores_view[[
                    'Unidade (Digiteam)', 'Consultora Responsável', 'Total_OS', 'OS_Ate_10d', 'Pct_10d', 'Gap_OS_Meta'
                ]].copy()
                tabela_ofensores.columns = [
                    'Posto Autorizado', 'Consultora', 'Total OS', '≤ 10 Dias', '% Atingido', 'Déficit (OS p/ Meta)'
                ]
                tabela_ofensores['% Atingido'] = tabela_ofensores['% Atingido'].map(lambda x: f"{x:.1f}%")
                st.dataframe(tabela_ofensores, use_container_width=True, height=410)
        else:
            st.success("🎉 Parabéns! Nenhum posto está abaixo da meta de 62% para os filtros e critérios selecionados.")

        st.divider()

        # -------------------------------------------------------------
        # DEMAIS GRÁFICOS VISUAIS
        # -------------------------------------------------------------
        g_col1, g_col2 = st.columns(2)

        with g_col1:
            st.subheader("Faixas de Atendimento (SLA 10d vs CDC 30d)")
            df_faixa_cnt = df_filtrado['Faixa_TAT'].value_counts().reset_index()
            df_faixa_cnt.columns = ['Faixa de Atendimento', 'Quantidade']
            fig_faixa = px.pie(
                df_faixa_cnt,
                names='Faixa de Atendimento',
                values='Quantidade',
                hole=0.45,
                color='Faixa de Atendimento',
                color_discrete_map={
                    'Até 10 dias (No Target)': '#10B981',
                    '11 a 30 dias (Alerta)': '#00A3E0',
                    '> 30 dias (Estouro CDC)': '#EF4444'
                }
            )
            fig_faixa.update_layout(height=340, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_faixa, use_container_width=True)

        with g_col2:
            st.subheader("Distribuição por Tipo de Demanda")
            df_dem_cnt = df_filtrado['Tipo de Demanda'].value_counts().head(7).reset_index()
            df_dem_cnt.columns = ['Tipo de Demanda', 'Quantidade']
            fig_dem = px.bar(
                df_dem_cnt,
                x='Tipo de Demanda',
                y='Quantidade',
                text='Quantidade',
                color_discrete_sequence=[PANASONIC_BLUE]
            )
            fig_dem.update_layout(showlegend=False, height=340, margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig_dem, use_container_width=True)

        # -------------------------------------------------------------
        # TABELA DINÂMICA DE CONSULTA DE ORDENS
        # -------------------------------------------------------------
        st.divider()
        st.subheader("🔍 Consulta Rápida de Atendimentos")
        busca = st.text_input("Pesquise por OS, Ticket, Modelo, Tipo de Demanda, Posto, Cidade ou Peça:")
        
        colunas_exibir = [
            'Prioridade', 'Faixa_TAT', 'TAT_Dias', 'Tem_Peca_Txt', 'Código OS', 'Número do Ticket',
            'Tipo de Demanda', 'Consultora Responsável', 'Unidade (Digiteam)', 'Status Atual', 'Data de Criação',
            'Categoria do Produto', 'Modelo do Produto', 'Cidade', 'Estado'
        ]
        colunas_validas = [c for c in colunas_exibir if c in df_filtrado.columns]
        
        df_tabela = df_filtrado[colunas_validas].copy()
        if busca:
            mask = df_tabela.astype(str).apply(lambda row: row.str.contains(busca, case=False).any(), axis=1)
            df_tabela = df_tabela[mask]

        st.dataframe(df_tabela.head(100), use_container_width=True, height=350)
        st.caption(f"Mostrando até 100 de {len(df_tabela)} ordens encontradas no filtro atual.")

        # -------------------------------------------------------------
        # GERADOR DE EXCEL SOB DEMANDA (EVITA ESTOURO DE MEMÓRIA AO FILTRAR)
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📥 Exportação de Planilhas Formatadas (Padrão Panasonic)")
        st.caption("Clique no botão abaixo para preparar o arquivo Excel da visão desejada sem pesar a navegação.")

        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.write("📊 **Planilha Geral Panasonic (Toda a Rede):**")
            if st.button("⚙️ Gerar Planilha Geral Completa", key="btn_gerar_geral"):
                with st.spinner("Montando Excel consolidado de toda a rede..."):
                    df_ofensores_geral = calcular_ofensores(df_sorted)
                    st.session_state["excel_geral_bytes"] = gerar_excel_completo(
                        df_sorted, df_ofensores_geral, data_corte_iso, "Controle Operacional Geral"
                    )
                    st.session_state["excel_geral_nome"] = f"Panasonic_Gestao_Rede_Consolidado_{data_referencia.strftime('%Y%m%d')}.xlsx"

            if "excel_geral_bytes" in st.session_state:
                st.download_button(
                    label="📥 Baixar Excel Geral Pronto (.xlsx)",
                    data=st.session_state["excel_geral_bytes"],
                    file_name=st.session_state["excel_geral_nome"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_geral"
                )

        with d_col2:
            tem_filtro = bool(sel_consultoras or sel_demanda or sel_faixas or sel_status or sel_prio or sel_uf)
            if tem_filtro:
                st.write("👩‍💼 **Carteira / Recorte Filtrado Atual:**")
                if st.button("⚙️ Gerar Planilha do Filtro Atual", key="btn_gerar_filtro"):
                    with st.spinner("Montando Excel da seleção filtrada..."):
                        st.session_state["excel_filtro_bytes"] = gerar_excel_completo(
                            df_filtrado, df_ofensores_filtrado, data_corte_iso, "Base Operacional Filtrada"
                        )
                        st.session_state["excel_filtro_nome"] = f"Panasonic_Gestao_Rede_Filtrada_{data_referencia.strftime('%Y%m%d')}.xlsx"

                if "excel_filtro_bytes" in st.session_state:
                    st.download_button(
                        label="📥 Baixar Excel do Filtro Atual (.xlsx)",
                        data=st.session_state["excel_filtro_bytes"],
                        file_name=st.session_state["excel_filtro_nome"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_filtro"
                    )
            else:
                st.info("💡 Aplique qualquer filtro na barra lateral (ex: Consultora, UF, Status) para habilitar o gerador exclusivo desse recorte.")
else:
    st.info("👆 Por favor, envie o relatório de ordens de serviço acima para iniciar o painel.")
