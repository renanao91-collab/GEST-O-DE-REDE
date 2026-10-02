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
        <div class="panasonic-subtitle">Acompanhamento Operacional de Ordens de Serviço, Controle de TAT de Reparo, Metas CDC e Gestão por Consultoras</div>
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
        is_reparado = []
        for i, r in df_sorted.iterrows():
            st_val = str(r.get('Status Atual', '')).strip().upper()
            reparado = st_val in ['CONCLUÍDA', 'FINALIZADA']
            is_reparado.append(reparado)
            
            if pd.notna(r['dt_criacao']):
                if pd.notna(dt_fim[i]):
                    tat_dias.append(max(0, int((dt_fim[i] - r['dt_criacao']).total_seconds() // 86400)))
                else:
                    tat_dias.append(max(0, int((hoje - r['dt_criacao']).total_seconds() // 86400)))
            else:
                tat_dias.append(0)
                
        df_sorted['TAT_Dias'] = tat_dias
        df_sorted['Is_Reparado'] = is_reparado

        # Faixas de TAT
        def get_faixa(d):
            if d <= 10: return "Até 10 dias"
            elif d <= 30: return "11 a 30 dias"
            else: return "> 30 dias (Estouro CDC)"
        df_sorted['Faixa_TAT'] = df_sorted['TAT_Dias'].apply(get_faixa)

        # Flag de Reincidência
        dem_col = df_sorted['Tipo de Demanda'].fillna('').astype(str).str.upper()
        df_sorted['Is_Reincidente'] = dem_col.str.contains('REINCIDENCIA|REINCIDÊNCIA', regex=True)

        # Flag de Peça
        c_pl = df_sorted.get('Código Peça Lançada', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        c_pt = df_sorted.get('Codigo Troca de Peças', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        df_sorted['Tem_Peca'] = (c_pl != '') | (c_pt != '')
        df_sorted['Tem_Peca_Txt'] = df_sorted['Tem_Peca'].map({True: 'SIM', False: 'NÃO'})

        # Matriz de Priorização Operacional
        def get_prioridade(row):
            st_val = str(row.get('Status Atual', '')).strip()
            if st_val == 'CANCELADA': return 'BAIXA'
            if row['Is_Reincidente'] or row['TAT_Dias'] > 30: return 'CRÍTICA'
            if (row['Tem_Peca'] or row['TAT_Dias'] >= 3) and not row['Is_Reparado']: return 'ALTA'
            if st_val == 'CRIADA': return 'MÉDIA'
            return 'NORMAL'
        df_sorted['Prioridade'] = df_sorted.apply(get_prioridade, axis=1)

    # -------------------------------------------------------------
    # BARRA LATERAL: FILTROS OPERACIONAIS (SELEÇÃO MÚLTIPLA)
    # -------------------------------------------------------------
    st.sidebar.markdown(f"<h3 style='color:{PANASONIC_BLUE};'>🎯 Filtros Operacionais</h3>", unsafe_allow_html=True)
    
    # 1. Filtro com SELEÇÃO MÚLTIPLA de Consultoras
    lista_consultoras = sorted(df_sorted['Consultora Responsável'].unique())
    sel_consultoras = st.sidebar.multiselect(
        "👩‍💼 Consultora(s) Responsável(is):",
        options=lista_consultoras,
        default=[],
        help="Deixe em branco para visualizar todas, ou selecione uma ou mais consultoras ao mesmo tempo."
    )

    # 2. Filtro de Tipo de Demanda
    lista_demandas = sorted(df_sorted['Tipo de Demanda'].dropna().astype(str).unique())
    sel_demanda = st.sidebar.multiselect("📋 Tipo de Demanda:", lista_demandas, default=[])

    # 3. Filtro de Faixa de Dias
    faixas_opcoes = ["Até 10 dias", "11 a 30 dias", "> 30 dias (Estouro CDC)"]
    sel_faixas = st.sidebar.multiselect("⏱️ Faixa de Dias (SLA / CDC):", faixas_opcoes, default=[])

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
    if sel_consultoras:
        df_filtrado = df_filtrado[df_filtrado['Consultora Responsável'].isin(sel_consultoras)]
    if sel_demanda:
        df_filtrado = df_filtrado[df_filtrado['Tipo de Demanda'].isin(sel_demanda)]
    if sel_faixas:
        df_filtrado = df_filtrado[df_filtrado['Faixa_TAT'].isin(sel_faixas)]
    if sel_status:
        df_filtrado = df_filtrado[df_filtrado['Status Atual'].isin(sel_status)]
    if sel_prio:
        df_filtrado = df_filtrado[df_filtrado['Prioridade'].isin(sel_prio)]
    if sel_uf:
        df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.upper().isin(sel_uf)]

    # -------------------------------------------------------------
    # CÁLCULO DOS INDICADORES SOLICITADOS
    # -------------------------------------------------------------
    total_os = len(df_filtrado)
    
    # Base válida (desconsiderando ordens canceladas para apuração de qualidade da assistência)
    df_validas = df_filtrado[df_filtrado['Status Atual'] != 'CANCELADA']
    total_validas = len(df_validas)
    
    # 1. Percentual com até 10 dias
    qtd_ate_10 = (df_validas['TAT_Dias'] <= 10).sum()
    pct_ate_10 = (qtd_ate_10 / total_validas * 100) if total_validas > 0 else 0
    
    # 2. Percentual com até 30 dias (CDC)
    qtd_ate_30 = (df_validas['TAT_Dias'] <= 30).sum()
    pct_ate_30 = (qtd_ate_30 / total_validas * 100) if total_validas > 0 else 0
    
    # 3. TAT Médio de Reparo (Apenas ordens concluídas/finalizadas com reparo efetivo)
    df_reparadas = df_filtrado[df_filtrado['Is_Reparado']]
    tat_medio_reparo = df_reparadas['TAT_Dias'].mean() if len(df_reparadas) > 0 else 0
    
    # 4. Percentual de atendimentos reincidentes
    qtd_reinc = df_validas['Is_Reincidente'].sum()
    pct_reinc = (qtd_reinc / total_validas * 100) if total_validas > 0 else 0

    # Nome da visão selecionada
    if sel_consultoras:
        if len(sel_consultoras) == 1:
            nome_visao = f"Consultora: {sel_consultoras[0]}"
        else:
            nome_visao = f"{len(sel_consultoras)} Consultoras Selecionadas ({', '.join(sel_consultoras[:2])}...)"
    else:
        nome_visao = "Toda a Rede Panasonic"

    st.markdown(f"### 📌 Painel de Indicadores Estratégicos: **{nome_visao}**")
    
    # Linha 1: Os 4 Indicadores Solicitados
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric(
        "⚡ Atendimentos ≤ 10 Dias",
        f"{pct_ate_10:.1f}%",
        f"{qtd_ate_10:,} de {total_validas:,} ordens válidas".replace(",", "."),
        help="Percentual de ordens válidas atendidas em até 10 dias corridos."
    )
    kpi2.metric(
        "⚖️ Atendimentos ≤ 30 Dias (CDC)",
        f"{pct_ate_30:.1f}%",
        f"{qtd_ate_30:,} de {total_validas:,} ordens válidas".replace(",", "."),
        help="Percentual de ordens dentro do limite legal de 30 dias do Artigo 18 do CDC."
    )
    kpi3.metric(
        "🔧 TAT Médio de Reparo",
        f"{tat_medio_reparo:.1f} dias",
        f"{len(df_reparadas):,} ordens concluídas".replace(",", "."),
        help="Tempo médio em dias para reparo e encerramento das ordens de serviço concluídas/finalizadas."
    )
    kpi4.metric(
        "🔄 Taxa de Reincidência",
        f"{pct_reinc:.2f}%",
        f"{qtd_reinc:,} ordens reincidentes".replace(",", "."),
        delta_color="inverse",
        help="Percentual de atendimentos identificados como reincidência na rede."
    )

    # Linha 2: Métricas de Volume Complementares
    c_vol1, c_vol2, c_vol3, c_vol4 = st.columns(4)
    em_aberto = df_filtrado[~df_filtrado['Status Atual'].isin(['CONCLUÍDA', 'FINALIZADA', 'CANCELADA'])].shape[0]
    com_pecas = df_filtrado[df_filtrado['Tem_Peca']].shape[0]
    criticas = df_filtrado[df_filtrado['Prioridade'] == 'CRÍTICA'].shape[0]

    c_vol1.metric("Total de Ordens na Visão", f"{total_os:,}".replace(",", "."))
    c_vol2.metric("Ordens em Aberto", f"{em_aberto:,}".replace(",", "."))
    c_vol3.metric("Ordens com Peças Lançadas", f"{com_pecas:,}".replace(",", "."))
    c_vol4.metric("Ordens Críticas / Reincidência", f"{criticas:,}".replace(",", "."))

    st.divider()

    # -------------------------------------------------------------
    # GRÁFICOS VISUAIS INTERATIVOS
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
                'Até 10 dias': '#10B981',
                '11 a 30 dias': '#00A3E0',
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

    # Postos com Maior Volume da Seleção
    st.subheader("🏢 Postos / Autorizadas com Mais Demandas")
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
    # GERADOR DE EXCEL OFICIAL (.XLSX) COM AS MÉTRICAS SOLICITADAS
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
            # Mapeamento de colunas:
            # D: TAT Dias (Conclusão K vs Criação J)
            # C: Faixa SLA
            # E: Tem Peça
            # B: Prioridade
            f_tat = f'=IF(K{r}="", INT(TODAY()-DATEVALUE(LEFT(J{r},10))), INT(DATEVALUE(LEFT(K{r},10))-DATEVALUE(LEFT(J{r},10))))' if dc else 0
            f_faixa = f'=IF(D{r}<=10, "Até 10 dias", IF(D{r}<=30, "11 a 30 dias", "> 30 dias (Estouro CDC)"))'
            f_peca = f'=IF(OR(R{r}<>"", T{r}<>""), "SIM", "NÃO")'
            f_prio = f'=IF(I{r}="CANCELADA", "BAIXA", IF(OR(ISNUMBER(SEARCH("REINCIDENCIA", H{r})), ISNUMBER(SEARCH("REINCIDÊNCIA", H{r})), D{r}>30), "CRÍTICA", IF(AND(E{r}="SIM", I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(AND(D{r}>=3, I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(I{r}="CRIADA", "MÉDIA", "NORMAL")))))'
            f_acao = f'=IF(OR(I{r}="CONCLUÍDA", I{r}="FINALIZADA"), "Encerrado com Sucesso", IF(I{r}="CANCELADA", "Verificar Motivo Cancelamento", IF(D{r}>30, "AÇÃO IMEDIATA: Estouro de Prazo CDC (Risco Jurídico/Troca)", IF(E{r}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(I{r}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(I{r}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento"))))))'

            ws_base.append([
                cons_resp[i], f_prio, f_faixa, f_tat, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
                cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
            ])

        ws_base.freeze_panes = "F5"
        ws_base.auto_filter.ref = f"A4:W{n_rows+4}"
        
        col_widths_base = {
            "A": 22, "B": 20, "C": 22, "D": 14, "E": 16, "F": 18, "G": 14, "H": 28, "I": 18,
            "J": 18, "K": 18, "L": 22, "M": 18, "N": 35, "O": 20, "P": 8, "Q": 25,
            "R": 24, "S": 35, "T": 24, "U": 35, "V": 30, "W": 40
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

        # Aba Dashboard
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

        # KPIs Excel com as fórmulas requisitadas:
        add_kpi(ws_dash, 1, "% ATÉ 10 DIAS", f'=COUNTIF(\'{nome_aba_base}\'!D5:D{n_rows+4}, "<=10") / MAX(1, COUNTA(\'{nome_aba_base}\'!D5:D{n_rows+4}))', "0.0%")
        add_kpi(ws_dash, 3, "% ATÉ 30 DIAS (CDC)", f'=COUNTIF(\'{nome_aba_base}\'!D5:D{n_rows+4}, "<=30") / MAX(1, COUNTA(\'{nome_aba_base}\'!D5:D{n_rows+4}))', "0.0%")
        add_kpi(ws_dash, 5, "TAT MÉDIO REPARO", f'=AVERAGEIFS(\'{nome_aba_base}\'!D5:D{n_rows+4}, \'{nome_aba_base}\'!I5:I{n_rows+4}, "CONCLUÍDA")', "0.0")
        add_kpi(ws_dash, 7, "% REINCIDENTES", f'=(COUNTIF(\'{nome_aba_base}\'!H5:H{n_rows+4}, "*REINCIDENCIA*") + COUNTIF(\'{nome_aba_base}\'!H5:H{n_rows+4}, "*REINCIDÊNCIA*")) / MAX(1, COUNTA(\'{nome_aba_base}\'!H5:H{n_rows+4}))', "0.00%")
        add_kpi(ws_dash, 9, "TOTAL DE ORDENS", f'=COUNTA(\'{nome_aba_base}\'!F5:F{n_rows+4})', "#,##0")
        add_kpi(ws_dash, 11, "ORDENS EM ABERTO", f'=COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "<>CONCLUÍDA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "FINALIZADA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "CANCELADA")', "#,##0")

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
        if sel_consultoras:
            lbl_cons = ", ".join(sel_consultoras[:2]) + ("..." if len(sel_consultoras) > 2 else "")
            st.write(f"👩‍💼 **Carteira Selecionada ({lbl_cons}):**")
            buf_cons = gerar_excel_completo(df_filtrado, "Base Consultoras Filtradas")
            st.download_button(
                label=f"📥 Baixar Excel da(s) Consultora(s) Selecionada(s)",
                data=buf_cons,
                file_name="Panasonic_Gestao_Rede_Consultoras_Filtradas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Selecione uma ou mais consultoras na barra lateral para habilitar o download filtrado da carteira delas.")
else:
    st.info("👆 Por favor, envie o relatório de ordens de serviço acima para iniciar o painel.")
