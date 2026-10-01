import os
import io
import time
import pandas as pd
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

st.markdown(f"""
<style>
    /* Estilos Gerais */
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
    /* Botões personalizados */
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
    /* Métricas */
    [data-testid="stMetricValue"] {{
        color: {PANASONIC_BLUE} !important;
        font-weight: 700 !important;
    }}
</style>
""", unsafe_allow_html=True)

# Topo com Identidade Panasonic
st.markdown(f"""
<div class="panasonic-header-container">
    <div>
        <div class="panasonic-title">Panasonic do Brasil | Gestão de Rede & Pós-Venda</div>
        <div class="panasonic-subtitle">Acompanhamento Operacional de Ordens de Serviço, Controle de TAT, Conformidade CDC (30 Dias), Rastreio de Peças e Gestão de Consultoras</div>
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

# Função de Leitura Resiliente
def ler_arquivo(file):
    if file.name.endswith('.csv'):
        try:
            return pd.read_csv(file, low_memory=False)
        except Exception:
            file.seek(0)
            return pd.read_csv(file, sep=';', low_memory=False)
    else:
        return pd.read_excel(file)

# -------------------------------------------------------------
# PROCESSAMENTO DE DADOS
# -------------------------------------------------------------
if file_servico is not None:
    with st.spinner("Processando ordens de serviço da rede Panasonic..."):
        df_raw = ler_arquivo(file_servico)
        df_raw.columns = [c.replace('\ufeff', '').replace('"', '').strip() for c in df_raw.columns]

        # Tratamento de datas
        df_raw['dt_criacao'] = pd.to_datetime(df_raw['Data de Criação'], errors='coerce')
        df_raw['dt_conclusao'] = pd.to_datetime(df_raw.get('Data de Conclusão', None), errors='coerce')
        df_raw['dt_finalizacao'] = pd.to_datetime(df_raw.get('Data de Finalização', None), errors='coerce')

        # Ordenar rigorosamente do mais antigo para o mais novo
        df_sorted = df_raw.sort_values(by='dt_criacao', ascending=True, na_position='last').reset_index(drop=True)

        # Mapeamento de Consultoras
        if file_consultoras is not None:
            df_cons = ler_arquivo(file_consultoras)
            df_cons.columns = [c.replace('\ufeff', '').replace('"', '').strip() for c in df_cons.columns]

            col_unid_cons = None
            for c in df_cons.columns:
                if any(k in c.lower() for k in ['unidade', 'autorizada', 'posto', 'nome', 'razão', 'credenciada']):
                    col_unid_cons = c
                    break
            if not col_unid_cons:
                col_unid_cons = df_cons.columns[0]

            col_nome_cons = None
            for c in df_cons.columns:
                if any(k in c.lower() for k in ['consultora', 'responsável', 'responsavel', 'gestora', 'atendente', 'consultor']):
                    col_nome_cons = c
                    break
            if not col_nome_cons:
                col_nome_cons = df_cons.columns[1] if len(df_cons.columns) > 1 else df_cons.columns[0]

            map_dict = dict(zip(
                df_cons[col_unid_cons].astype(str).str.strip().str.upper(),
                df_cons[col_nome_cons].astype(str).str.strip()
            ))
            
            df_sorted['Consultora Responsável'] = df_sorted['Unidade (Digiteam)'].astype(str).str.strip().str.upper().map(map_dict)
            df_sorted['Consultora Responsável'] = df_sorted['Consultora Responsável'].fillna('Não Atribuída')
        else:
            df_sorted['Consultora Responsável'] = 'Não Atribuída (Insira a planilha 2)'

        # Cálculo de TAT em dias
        hoje = pd.Timestamp.now()
        dt_fim = df_sorted['dt_conclusao'].fillna(df_sorted['dt_finalizacao'])
        
        tat_dias = []
        for i, r in df_sorted.iterrows():
            if pd.notna(r['dt_criacao']):
                if pd.notna(dt_fim[i]):
                    tat_dias.append(max(0, int((dt_fim[i] - r['dt_criacao']).total_seconds() // 86400)))
                else:
                    tat_dias.append(max(0, int((hoje - r['dt_criacao']).total_seconds() // 86400)))
            else:
                tat_dias.append(0)
        df_sorted['TAT_Dias'] = tat_dias

        # Faixas de TAT
        def get_faixa(d):
            if d <= 1: return "0-1 dia"
            elif d <= 2: return "2 dias"
            elif d <= 3: return "3 dias"
            else: return "> 3 dias"
        df_sorted['Faixa_TAT'] = df_sorted['TAT_Dias'].apply(get_faixa)

        # Indicador de Conformidade CDC (Prazo legal de até 30 dias)
        df_sorted['Status_CDC'] = df_sorted['TAT_Dias'].apply(lambda d: "Dentro do Prazo (≤ 30d)" if d <= 30 else "Estouro CDC (> 30d)")

        # Flag de Peça
        c_pl = df_sorted.get('Código Peça Lançada', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        c_pt = df_sorted.get('Codigo Troca de Peças', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        df_sorted['Tem_Peca'] = (c_pl != '') | (c_pt != '')
        df_sorted['Tem_Peca_Txt'] = df_sorted['Tem_Peca'].map({True: 'SIM', False: 'NÃO'})

        # Matriz de Priorização Operacional
        def get_prioridade(row):
            st_val = str(row.get('Status Atual', '')).strip()
            dem_val = str(row.get('Tipo de Demanda', '')).strip().upper()
            if st_val == 'CANCELADA': return 'BAIXA'
            if 'REINCIDENCIA' in dem_val or 'REINCIDÊNCIA' in dem_val or row['TAT_Dias'] > 30: return 'CRÍTICA'
            if (row['Tem_Peca'] or row['TAT_Dias'] >= 3) and st_val not in ['CONCLUÍDA', 'FINALIZADA']: return 'ALTA'
            if st_val == 'CRIADA': return 'MÉDIA'
            return 'NORMAL'
        df_sorted['Prioridade'] = df_sorted.apply(get_prioridade, axis=1)

    # -------------------------------------------------------------
    # BARRA LATERAL: FILTROS OPERACIONAIS
    # -------------------------------------------------------------
    st.sidebar.markdown(f"<h3 style='color:{PANASONIC_BLUE};'>🎯 Filtros Operacionais</h3>", unsafe_allow_html=True)
    
    # 1. Filtro de Consultora
    lista_consultoras = sorted(df_sorted['Consultora Responsável'].unique())
    sel_consultora = st.sidebar.selectbox("👩‍💼 Consultora Responsável:", ["TODAS"] + lista_consultoras)

    # 2. Filtro de Tipo de Demanda
    lista_demandas = sorted(df_sorted['Tipo de Demanda'].dropna().astype(str).unique())
    sel_demanda = st.sidebar.multiselect("📋 Tipo de Demanda:", lista_demandas, default=[])

    # 3. NOVO: Filtro de Conformidade CDC
    lista_cdc = ["Dentro do Prazo (≤ 30d)", "Estouro CDC (> 30d)"]
    sel_cdc = st.sidebar.multiselect("⚖️ Conformidade CDC (30 Dias):", lista_cdc, default=[])

    # 4. Filtro de Status
    lista_status = sorted(df_sorted['Status Atual'].dropna().unique())
    sel_status = st.sidebar.multiselect("🚦 Status Operacional:", lista_status, default=[])

    # 5. Filtro de Prioridade
    lista_prio = ["CRÍTICA", "ALTA", "MÉDIA", "NORMAL", "BAIXA"]
    sel_prio = st.sidebar.multiselect("⚡ Nível de Prioridade:", lista_prio, default=[])

    # 6. Filtro de UF
    ufs = sorted(df_sorted['Estado'].dropna().astype(str).str.upper().unique())
    sel_uf = st.sidebar.multiselect("📍 Estado (UF):", ufs, default=[])

    # Aplicação Dinâmica dos Filtros
    df_filtrado = df_sorted.copy()
    if sel_consultora != "TODAS":
        df_filtrado = df_filtrado[df_filtrado['Consultora Responsável'] == sel_consultora]
    if sel_demanda:
        df_filtrado = df_filtrado[df_filtrado['Tipo de Demanda'].isin(sel_demanda)]
    if sel_cdc:
        df_filtrado = df_filtrado[df_filtrado['Status_CDC'].isin(sel_cdc)]
    if sel_status:
        df_filtrado = df_filtrado[df_filtrado['Status Atual'].isin(sel_status)]
    if sel_prio:
        df_filtrado = df_filtrado[df_filtrado['Prioridade'].isin(sel_prio)]
    if sel_uf:
        df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.upper().isin(sel_uf)]

    # -------------------------------------------------------------
    # CARDS DE INDICADORES (KPIs) COM PRAZO CDC
    # -------------------------------------------------------------
    total_os = len(df_filtrado)
    tat_medio = df_filtrado['TAT_Dias'].mean() if total_os > 0 else 0
    em_aberto = df_filtrado[~df_filtrado['Status Atual'].isin(['CONCLUÍDA', 'FINALIZADA', 'CANCELADA'])].shape[0]
    com_pecas = df_filtrado[df_filtrado['Tem_Peca']].shape[0]
    criticas = df_filtrado[df_filtrado['Prioridade'] == 'CRÍTICA'].shape[0]

    # Cálculo da taxa CDC (considerando ordens válidas, excluindo canceladas)
    df_validas_cdc = df_filtrado[df_filtrado['Status Atual'] != 'CANCELADA']
    total_validas_cdc = len(df_validas_cdc)
    dentro_cdc_cnt = (df_validas_cdc['TAT_Dias'] <= 30).sum()
    pct_cdc = (dentro_cdc_cnt / total_validas_cdc * 100) if total_validas_cdc > 0 else 0

    st.markdown(f"### 📌 Visão Consolidada: **{sel_consultora if sel_consultora != 'TODAS' else 'Toda a Rede Panasonic'}**")
    
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Total de Ordens", f"{total_os:,}".replace(",", "."))
    kpi2.metric("TAT Médio Geral", f"{tat_medio:.1f} dias")
    kpi3.metric("🎯 % Dentro CDC (≤30d)", f"{pct_cdc:.1f}%", help="Percentual de ordens com resolução ou tempo em aberto de até 30 dias corridos.")
    kpi4.metric("OS em Aberto", f"{em_aberto:,}".replace(",", "."))
    kpi5.metric("OS com Peças", f"{com_pecas:,}".replace(",", "."))
    kpi6.metric("OS Críticas / Reincidência", f"{criticas:,}".replace(",", "."))

    st.divider()

    # -------------------------------------------------------------
    # GRÁFICOS VISUAIS INTERATIVOS
    # -------------------------------------------------------------
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.subheader("Conformidade com o Prazo CDC (30 Dias)")
        df_cdc_cnt = df_filtrado['Status_CDC'].value_counts().reset_index()
        df_cdc_cnt.columns = ['Status CDC', 'Quantidade']
        fig_cdc = px.pie(
            df_cdc_cnt,
            names='Status CDC',
            values='Quantidade',
            hole=0.45,
            color='Status CDC',
            color_discrete_map={
                'Dentro do Prazo (≤ 30d)': '#004098',
                'Estouro CDC (> 30d)': '#EF4444'
            }
        )
        fig_cdc.update_layout(height=340, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_cdc, use_container_width=True)

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

    # Postos com Maior Volume
    if sel_consultora != "TODAS":
        st.subheader(f"🏢 Postos com Mais Demandas - {sel_consultora}")
    else:
        st.subheader("🏢 Top 10 Autorizadas com Mais Demandas da Rede")
    
    top_unidades = df_filtrado['Unidade (Digiteam)'].value_counts().head(10).reset_index()
    top_unidades.columns = ['Autorizada / Unidade', 'Ordens']
    fig_unid = px.bar(
        top_unidades,
        x='Ordens',
        y='Autorizada / Unidade',
        orientation='h',
        text='Ordens',
        color_discrete_sequence=[PANASONIC_DARK]
    )
    fig_unid.update_layout(yaxis={'categoryorder':'total ascending'}, height=330, margin=dict(l=20, r=20, t=20, b=20))
    st.plotly_chart(fig_unid, use_container_width=True)

    # -------------------------------------------------------------
    # TABELA DINÂMICA DE CONSULTA
    # -------------------------------------------------------------
    st.divider()
    st.subheader("🔍 Consulta Rápida de Atendimentos")
    busca = st.text_input("Pesquise por OS, Ticket, Modelo, Tipo de Demanda, Posto, Cidade ou Peça:")
    
    colunas_exibir = [
        'Prioridade', 'Status_CDC', 'TAT_Dias', 'Faixa_TAT', 'Tem_Peca_Txt', 'Código OS', 'Número do Ticket',
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
    # GERADOR DE EXCEL OFICIAL (.XLSX) COM INDICADOR CDC
    # -------------------------------------------------------------
    st.divider()
    st.subheader("📥 Exportação de Planilhas Formatadas (Padrão Panasonic)")

    def gerar_excel_completo(df_export, nome_aba_base="Controle Operacional OS"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        ws_dash = wb.create_sheet(title="Dashboard Gestão")
        ws_base = wb.create_sheet(title=nome_aba_base)
        ws_pecas = wb.create_sheet(title="Controle de Peças")

        COLOR_PANASONIC_EXCEL = "004098"
        WHITE = "FFFFFF"
        BORDER_GRAY = "D1D5DB"
        
        font_title = Font(name="Segoe UI", size=15, bold=True, color=COLOR_PANASONIC_EXCEL)
        font_sub = Font(name="Segoe UI", size=9, italic=True, color="4B5563")
        font_header = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        fill_header = PatternFill(start_color=COLOR_PANASONIC_EXCEL, end_color=COLOR_PANASONIC_EXCEL, fill_type="solid")

        # Cabeçalho Base
        ws_base["A1"] = f"PANASONIC DO BRASIL - GESTÃO DE REDE & PÓS-VENDA"
        ws_base["A1"].font = font_title
        ws_base["A2"] = "Acompanhamento Operacional | Ordenado do mais antigo para o mais novo | Métricas de TAT, CDC e Peças"
        ws_base["A2"].font = font_sub

        headers_base = [
            "Consultora Responsável", "Prioridade Operacional", "Status Prazo CDC (30d)", "TAT Atual (Dias)", "Faixa de TAT", "Tem Peça Lançada?",
            "Código da OS", "Ticket", "Tipo de Demanda", "Status Atual", "Data de Criação",
            "Data de Conclusão", "Categoria do Produto", "Modelo do Produto", "Unidade Técnica (Rede)",
            "Região / Cidade", "UF", "Técnico de Campo", "Código Peça Lançada", "Descrição Peça Lançada",
            "Código Peça Trocada", "Descrição Peça Trocada", "Sintoma / Defeito Relatado", "Ação Operacional Recomendada"
        ]
        for c_idx, h in enumerate(headers_base, 1):
            c = ws_base.cell(row=4, column=c_idx, value=h)
            c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_base.row_dimensions[4].height = 26

        n_rows = len(df_export)
        cod_os = df_export['Código OS'].astype(str).tolist()
        ticket = df_export.get('Número do Ticket', pd.Series(['']*n_rows)).fillna('').astype(str).str.replace(r'\.0$', '', regex=True).tolist()
        demanda = df_export.get('Tipo de Demanda', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        status = df_export.get('Status Atual', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        dt_criac_str = [d.strftime('%Y-%m-%d %H:%M') if pd.notna(d) else '' for d in df_export['dt_criacao']]
        dt_concl_comb = df_export['dt_conclusao'].fillna(df_export.get('dt_finalizacao', pd.Series([None]*n_rows)))
        dt_concl_str = [d.strftime('%Y-%m-%d %H:%M') if pd.notna(d) else '' for d in dt_concl_comb]
        cat = df_export.get('Categoria do Produto', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        mod = df_export.get('Modelo do Produto', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        unid = df_export.get('Unidade (Digiteam)', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        cid = df_export.get('Cidade', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        uf = df_export.get('Estado', pd.Series(['']*n_rows)).fillna('').astype(str).str.upper().tolist()
        tec = df_export.get('Agente de Campo', pd.Series(['']*n_rows)).fillna('').astype(str).tolist()
        c_pl = df_export.get('Código Peça Lançada', pd.Series(['']*n_rows)).fillna('').astype(str).str.strip().tolist()
        d_pl = df_export.get('Peça Lançada', pd.Series(['']*n_rows)).fillna('').astype(str).str.strip().tolist()
        c_pt = df_export.get('Codigo Troca de Peças', pd.Series(['']*n_rows)).fillna('').astype(str).str.strip().tolist()
        d_pt = df_export.get('Troca de Peças', pd.Series(['']*n_rows)).fillna('').astype(str).str.strip().tolist()
        sint = df_export.get('Sintoma e Defeito', pd.Series(['']*n_rows)).fillna('').astype(str).str.strip().tolist()
        cons_resp = df_export['Consultora Responsável'].tolist()

        for i in range(n_rows):
            r = i + 5
            dc = dt_criac_str[i]
            # Mapeamento de colunas da linha:
            # A: Consultora
            # B: Prioridade
            # C: Status Prazo CDC (30d) -> =IF(D{r}<=30, "Dentro do CDC", "ESTOURO CDC")
            # D: TAT (Dias) -> Conclusão L vs Criação K
            # E: Faixa TAT
            # F: Tem Peça -> S ou U
            # G: Código OS, H: Ticket, I: Demanda, J: Status, K: Criação, L: Conclusão
            f_tat = f'=IF(L{r}="", INT(TODAY()-DATEVALUE(LEFT(K{r},10))), INT(DATEVALUE(LEFT(L{r},10))-DATEVALUE(LEFT(K{r},10))))' if dc else 0
            f_cdc = f'=IF(D{r}<=30, "Dentro do CDC", "ESTOURO CDC")'
            f_peca = f'=IF(OR(S{r}<>"", U{r}<>""), "SIM", "NÃO")'
            f_prio = f'=IF(J{r}="CANCELADA", "BAIXA", IF(OR(ISNUMBER(SEARCH("REINCIDENCIA", I{r})), ISNUMBER(SEARCH("REINCIDÊNCIA", I{r})), D{r}>30), "CRÍTICA", IF(AND(F{r}="SIM", J{r}<>"FINALIZADA", J{r}<>"CONCLUÍDA"), "ALTA", IF(AND(D{r}>=3, J{r}<>"FINALIZADA", J{r}<>"CONCLUÍDA"), "ALTA", IF(J{r}="CRIADA", "MÉDIA", "NORMAL")))))'
            f_faixa = f'=IF(D{r}<=1, "0-1 dia", IF(D{r}<=2, "2 dias", IF(D{r}<=3, "3 dias", "> 3 dias")))'
            f_acao = f'=IF(OR(J{r}="CONCLUÍDA", J{r}="FINALIZADA"), "Encerrado com Sucesso", IF(J{r}="CANCELADA", "Verificar Motivo Cancelamento", IF(D{r}>30, "AÇÃO IMEDIATA: Estouro de Prazo CDC (Risco Jurídico/Troca)", IF(F{r}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(J{r}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(J{r}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento"))))))'

            ws_base.append([
                cons_resp[i], f_prio, f_cdc, f_tat, f_faixa, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
                cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
            ])

        ws_base.freeze_panes = "G5"
        ws_base.auto_filter.ref = f"A4:X{n_rows+4}"
        
        col_widths_base = {
            "A": 22, "B": 20, "C": 20, "D": 14, "E": 13, "F": 16, "G": 18, "H": 14, "I": 28, "J": 18,
            "K": 18, "L": 18, "M": 22, "N": 18, "O": 35, "P": 20, "Q": 8, "R": 25,
            "S": 24, "T": 35, "U": 24, "V": 35, "W": 30, "X": 40
        }
        for col_letter, width in col_widths_base.items():
            ws_base.column_dimensions[col_letter].width = width

        # Aba Peças
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

        ws_pecas.auto_filter.ref = f"A4:L{count_p+4}"
        ws_pecas.freeze_panes = "D5"

        # Aba Dashboard (Com Indicador do CDC)
        ws_dash.views.sheetView[0].showGridLines = True
        ws_dash["A1"] = "PANASONIC DO BRASIL - PAINEL DE CONTROLE E GESTÃO DA REDE"; ws_dash["A1"].font = font_title

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

        # KPIs no Excel (6 Cards)
        add_kpi(ws_dash, 1, "TOTAL DE ORDENS", f'=COUNTA(\'{nome_aba_base}\'!G5:G{n_rows+4})')
        add_kpi(ws_dash, 3, "TAT MÉDIO GERAL (DIAS)", f'=AVERAGE(\'{nome_aba_base}\'!D5:D{n_rows+4})', "0.0")
        add_kpi(ws_dash, 5, "🎯 % DENTRO CDC (≤ 30D)", f'=COUNTIF(\'{nome_aba_base}\'!C5:C{n_rows+4}, "Dentro do CDC") / MAX(1, COUNTA(\'{nome_aba_base}\'!C5:C{n_rows+4}))', "0.0%")
        add_kpi(ws_dash, 7, "ORDENS EM ABERTO", f'=COUNTIF(\'{nome_aba_base}\'!J5:J{n_rows+4}, "<>CONCLUÍDA") - COUNTIF(\'{nome_aba_base}\'!J5:J{n_rows+4}, "FINALIZADA") - COUNTIF(\'{nome_aba_base}\'!J5:J{n_rows+4}, "CANCELADA")')
        add_kpi(ws_dash, 9, "OS COM PEÇAS", f'=COUNTIF(\'{nome_aba_base}\'!F5:F{n_rows+4}, "SIM")')
        add_kpi(ws_dash, 11, "REINCIDÊNCIAS / CRÍTICAS", f'=COUNTIF(\'{nome_aba_base}\'!B5:B{n_rows+4}, "CRÍTICA")')

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.write("📊 **Planilha Geral Panasonic (Toda a Rede):**")
        buf_geral = gerar_excel_completo(df_sorted, "Controle Operacional Geral")
        st.download_button(
            label="📥 Baixar Excel Consolidado (Toda a Rede)",
            data=buf_geral,
            file_name="Panasonic_Gestao_de_Rede_Consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with d_col2:
        if sel_consultora != "TODAS":
            st.write(f"👩‍💼 **Carteira da Consultora ({sel_consultora}):**")
            buf_cons = gerar_excel_completo(df_filtrado, f"Base {sel_consultora[:20]}")
            st.download_button(
                label=f"📥 Baixar Carteira de {sel_consultora}",
                data=buf_cons,
                file_name=f"Panasonic_Gestao_Rede_{sel_consultora.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Selecione uma consultora específica na barra lateral para habilitar o download individual da carteira dela.")
else:
    st.info("👆 Por favor, envie o relatório de ordens de serviço acima para iniciar o painel.")
