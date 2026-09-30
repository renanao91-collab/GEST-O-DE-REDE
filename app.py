import os
import io
import time
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Central de Gestão de Rede & Consultoras",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# ESTILOS CSS CUSTOMIZADOS
# -------------------------------------------------------------
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 5px;
    }
    .sub-header {
        font-size: 14px;
        color: #4B5563;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F8FAFC;
        border-left: 4px solid #1B365D;
        padding: 12px;
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">⚡ Central de Gestão de Rede & Acompanhamento de Atendimentos</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Painel inteligente com mapeamento de Consultoras Responsáveis, monitoramento de TAT, rastreio de peças e priorização operacional.</div>', unsafe_allow_html=True)

# -------------------------------------------------------------
# ÁREA DE UPLOAD (2 ARQUIVOS)
# -------------------------------------------------------------
col_up1, col_up2 = st.columns(2)

with col_up1:
    st.subheader("1. Relatório de Serviços")
    file_servico = st.file_uploader(
        "Selecione o relatório de chamados (CSV ou XLSX):",
        type=["csv", "xlsx"],
        key="servico"
    )

with col_up2:
    st.subheader("2. Base de Consultoras (De-Para)")
    file_consultoras = st.file_uploader(
        "Selecione a planilha com Unidade x Consultora (XLSX ou CSV):",
        type=["xlsx", "csv"],
        key="consultoras"
    )
    with st.expander("ℹ️ Como deve ser a planilha de Consultoras?"):
        st.write("""
        Deve conter pelo menos duas colunas:
        1. **Nome da Unidade / Autorizada** (ex: *Unidade*, *Autorizada*, *Posto*, ou *Unidade (Digiteam)*)
        2. **Consultora Responsável** (ex: *Consultora*, *Responsável*, *Gestora*)
        """)

# Função utilitária para leitura resiliente
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
    with st.spinner("Lendo e estruturando dados do relatório..."):
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

            # Detectar coluna de Unidade / Posto
            col_unid_cons = None
            for c in df_cons.columns:
                if any(k in c.lower() for k in ['unidade', 'autorizada', 'posto', 'nome', 'razão', 'credenciada']):
                    col_unid_cons = c
                    break
            if not col_unid_cons:
                col_unid_cons = df_cons.columns[0]

            # Detectar coluna de Consultora
            col_nome_cons = None
            for c in df_cons.columns:
                if any(k in c.lower() for k in ['consultora', 'responsável', 'responsavel', 'gestora', 'atendente', 'consultor']):
                    col_nome_cons = c
                    break
            if not col_nome_cons:
                col_nome_cons = df_cons.columns[1] if len(df_cons.columns) > 1 else df_cons.columns[0]

            # Normalizar para merge
            map_dict = dict(zip(df_cons[col_unid_cons].astype(str).str.strip().str.upper(), df_cons[col_nome_cons].astype(str).str.strip()))
            
            # Aplicar no dataframe principal
            df_sorted['Consultora Responsável'] = df_sorted['Unidade (Digiteam)'].astype(str).str.strip().str.upper().map(map_dict)
            df_sorted['Consultora Responsável'] = df_sorted['Consultora Responsável'].fillna('Não Atribuída')
        else:
            df_sorted['Consultora Responsável'] = 'Não Atribuída (Insira a planilha 2)'

        # Cálculo de TAT em dias
        hoje = pd.Timestamp.now()
        dt_fim = df_sorted['dt_conclusao'].fillna(df_sorted['dt_finalizacao'])
        
        # Para concluídas: fim - criação; Para abertas: hoje - criação
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

        # Faixa de TAT
        def get_faixa(d):
            if d <= 1: return "0-1 dia"
            elif d <= 2: return "2 dias"
            elif d <= 3: return "3 dias"
            else: return "> 3 dias"
        df_sorted['Faixa_TAT'] = df_sorted['TAT_Dias'].apply(get_faixa)

        # Flag de Peça
        c_pl = df_sorted.get('Código Peça Lançada', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        c_pt = df_sorted.get('Codigo Troca de Peças', pd.Series(['']*len(df_sorted))).fillna('').astype(str).str.strip()
        df_sorted['Tem_Peca'] = (c_pl != '') | (c_pt != '')
        df_sorted['Tem_Peca_Txt'] = df_sorted['Tem_Peca'].map({True: 'SIM', False: 'NÃO'})

        # Priorização
        def get_prioridade(row):
            st_val = str(row.get('Status Atual', '')).strip()
            dem_val = str(row.get('Tipo de Demanda', '')).strip().upper()
            if st_val == 'CANCELADA': return 'BAIXA'
            if 'REINCIDENCIA' in dem_val: return 'CRÍTICA'
            if (row['Tem_Peca'] or row['TAT_Dias'] >= 3) and st_val not in ['CONCLUÍDA', 'FINALIZADA']: return 'ALTA'
            if st_val == 'CRIADA': return 'MÉDIA'
            return 'NORMAL'
        df_sorted['Prioridade'] = df_sorted.apply(get_prioridade, axis=1)

    # -------------------------------------------------------------
    # BARRA LATERAL: FILTROS DINÂMICOS
    # -------------------------------------------------------------
    st.sidebar.header("🎯 Filtros Operacionais")
    
    # Filtro de Consultora
    lista_consultoras = sorted(df_sorted['Consultora Responsável'].unique())
    sel_consultora = st.sidebar.selectbox("Consultora Responsável:", ["TODAS"] + lista_consultoras)

    # Filtro de Status
    lista_status = sorted(df_sorted['Status Atual'].dropna().unique())
    sel_status = st.sidebar.multiselect("Status Operacional:", lista_status, default=[])

    # Filtro de Prioridade
    lista_prio = ["CRÍTICA", "ALTA", "MÉDIA", "NORMAL", "BAIXA"]
    sel_prio = st.sidebar.multiselect("Prioridade:", lista_prio, default=[])

    # Filtro de UF
    ufs = sorted(df_sorted['Estado'].dropna().astype(str).str.upper().unique())
    sel_uf = st.sidebar.multiselect("Estado (UF):", ufs, default=[])

    # Aplicação dos Filtros
    df_filtrado = df_sorted.copy()
    if sel_consultora != "TODAS":
        df_filtrado = df_filtrado[df_filtrado['Consultora Responsável'] == sel_consultora]
    if sel_status:
        df_filtrado = df_filtrado[df_filtrado['Status Atual'].isin(sel_status)]
    if sel_prio:
        df_filtrado = df_filtrado[df_filtrado['Prioridade'].isin(sel_prio)]
    if sel_uf:
        df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.upper().isin(sel_uf)]

    # -------------------------------------------------------------
    # CARDS DE INDICADORES (KPIs)
    # -------------------------------------------------------------
    total_os = len(df_filtrado)
    tat_medio = df_filtrado['TAT_Dias'].mean() if total_os > 0 else 0
    em_aberto = df_filtrado[~df_filtrado['Status Atual'].isin(['CONCLUÍDA', 'FINALIZADA', 'CANCELADA'])].shape[0]
    com_pecas = df_filtrado[df_filtrado['Tem_Peca']].shape[0]
    criticas = df_filtrado[df_filtrado['Prioridade'] == 'CRÍTICA'].shape[0]

    st.markdown(f"### 📌 Visão Consolidada: **{sel_consultora if sel_consultora != 'TODAS' else 'Toda a Rede'}**")
    
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Total de Ordens", f"{total_os:,}".replace(",", "."))
    kpi2.metric("TAT Médio Geral", f"{tat_medio:.1f} dias")
    kpi3.metric("OS em Aberto", f"{em_aberto:,}".replace(",", "."))
    kpi4.metric("OS com Peças", f"{com_pecas:,}".replace(",", "."))
    kpi5.metric("OS Críticas (Reincidência)", f"{criticas:,}".replace(",", "."))

    st.divider()

    # -------------------------------------------------------------
    # GRÁFICOS VISUAIS INTERATIVOS
    # -------------------------------------------------------------
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.subheader("Distribuição por Status Atual")
        df_status_cnt = df_filtrado['Status Atual'].value_counts().reset_index()
        df_status_cnt.columns = ['Status', 'Quantidade']
        fig_status = px.bar(
            df_status_cnt.head(8),
            x='Status',
            y='Quantidade',
            text='Quantidade',
            color='Status',
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_status.update_layout(showlegend=False, height=350, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_status, use_container_width=True)

    with g_col2:
        st.subheader("Distribuição por Faixa de TAT")
        df_tat_cnt = df_filtrado['Faixa_TAT'].value_counts().reset_index()
        df_tat_cnt.columns = ['Faixa', 'Quantidade']
        fig_tat = px.pie(
            df_tat_cnt,
            names='Faixa',
            values='Quantidade',
            hole=0.45,
            color='Faixa',
            color_discrete_map={
                '0-1 dia': '#10B981',
                '2 dias': '#F59E0B',
                '3 dias': '#F97316',
                '> 3 dias': '#EF4444'
            }
        )
        fig_tat.update_layout(height=350, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_tat, use_container_width=True)

    # Top Postos da Consultora
    if sel_consultora != "TODAS":
        st.subheader(f"🏢 Postos com Mais Demandas - {sel_consultora}")
    else:
        st.subheader("🏢 Top 10 Autorizadas com Mais Demandas")
    
    top_unidades = df_filtrado['Unidade (Digiteam)'].value_counts().head(10).reset_index()
    top_unidades.columns = ['Autorizada / Unidade', 'Ordens']
    fig_unid = px.bar(top_unidades, x='Ordens', y='Autorizada / Unidade', orientation='h', text='Ordens', color_discrete_sequence=['#1B365D'])
    fig_unid.update_layout(yaxis={'categoryorder':'total ascending'}, height=320, margin=dict(l=20, r=20, t=20, b=20))
    st.plotly_chart(fig_unid, use_container_width=True)

    # -------------------------------------------------------------
    # TABELA DINÂMICA DE CONSULTA
    # -------------------------------------------------------------
    st.divider()
    st.subheader("🔍 Consulta Rápida de Ordens de Serviço")
    busca = st.text_input("Pesquise por OS, Ticket, Modelo, Cidade ou Peça:")
    
    colunas_exibir = [
        'Prioridade', 'TAT_Dias', 'Faixa_TAT', 'Tem_Peca_Txt', 'Código OS', 'Número do Ticket',
        'Consultora Responsável', 'Unidade (Digiteam)', 'Status Atual', 'Data de Criação',
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
    # GERADOR DE EXCEL OFICIAL (.XLSX) COM ABAS E FÓRMULAS
    # -------------------------------------------------------------
    st.divider()
    st.subheader("📥 Exportação de Planilhas Formatadas")

    def gerar_excel_completo(df_export, nome_aba_base="Controle Operacional OS"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        ws_dash = wb.create_sheet(title="Dashboard Gestão")
        ws_base = wb.create_sheet(title=nome_aba_base)
        ws_pecas = wb.create_sheet(title="Controle de Peças")

        NAVY = "1B365D"
        WHITE = "FFFFFF"
        BORDER_GRAY = "D1D5DB"
        font_title = Font(name="Segoe UI", size=15, bold=True, color=NAVY)
        font_sub = Font(name="Segoe UI", size=9, italic=True, color="4B5563")
        font_header = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
        fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

        # Cabeçalho Base
        ws_base["A1"] = f"GESTÃO DE REDE - ACOMPANHAMENTO OPERACIONAL"
        ws_base["A1"].font = font_title
        ws_base["A2"] = "Ordenado do chamado mais antigo para o mais novo | Métricas de TAT e Rastreio de Peças"
        ws_base["A2"].font = font_sub

        headers_base = [
            "Consultora Responsável", "Prioridade Operacional", "TAT Atual (Dias)", "Faixa de TAT", "Tem Peça Lançada?",
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
            # Formulado no Excel:
            # K: Conclusão, J: Criação, C: TAT, E: Peça, B: Prioridade
            f_tat = f'=IF(K{r}="", INT(TODAY()-DATEVALUE(LEFT(J{r},10))), INT(DATEVALUE(LEFT(K{r},10))-DATEVALUE(LEFT(J{r},10))))' if dc else 0
            f_peca = f'=IF(OR(R{r}<>"", T{r}<>""), "SIM", "NÃO")'
            f_prio = f'=IF(I{r}="CANCELADA", "BAIXA", IF(ISNUMBER(SEARCH("REINCIDENCIA", H{r})), "CRÍTICA", IF(AND(E{r}="SIM", I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(AND(C{r}>=3, I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(I{r}="CRIADA", "MÉDIA", "NORMAL")))))'
            f_faixa = f'=IF(C{r}<=1, "0-1 dia", IF(C{r}<=2, "2 dias", IF(C{r}<=3, "3 dias", "> 3 dias")))'
            f_acao = f'=IF(OR(I{r}="CONCLUÍDA", I{r}="FINALIZADA"), "Encerrado com Sucesso", IF(I{r}="CANCELADA", "Verificar Motivo Cancelamento", IF(E{r}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(I{r}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(I{r}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento")))))'

            ws_base.append([
                cons_resp[i], f_prio, f_tat, f_faixa, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
                cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
            ])

        ws_base.freeze_panes = "F5"
        ws_base.auto_filter.ref = f"A4:W{n_rows+4}"
        
        col_widths_base = {
            "A": 22, "B": 20, "C": 14, "D": 13, "E": 16, "F": 18, "G": 14, "H": 28, "I": 18,
            "J": 18, "K": 18, "L": 22, "M": 18, "N": 35, "O": 20, "P": 8, "Q": 25,
            "R": 24, "S": 35, "T": 24, "U": 35, "V": 30, "W": 36
        }
        for col_letter, width in col_widths_base.items():
            ws_base.column_dimensions[col_letter].width = width

        # Aba Peças
        ws_pecas["A1"] = "GESTÃO DE REDE - CONTROLE ESPECÍFICO DE PEÇAS"
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
        ws_dash["A1"] = "PAINEL DE CONTROLE E GESTÃO DA REDE DE ATENDIMENTO"; ws_dash["A1"].font = font_title

        def add_kpi(ws, col, title, form, fmt="#,##0"):
            ws.cell(row=4, column=col, value=title).font = Font(name="Segoe UI", size=9, bold=True, color="6B7280")
            c = ws.cell(row=5, column=col, value=form)
            c.font = Font(name="Segoe UI", size=18, bold=True, color=NAVY); c.number_format = fmt
            for r in range(4, 7):
                for cell_c in range(col, col + 2):
                    cell = ws.cell(row=r, column=cell_c)
                    cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
                    top_s = Side(border_style="medium", color=NAVY) if r == 4 else Side(border_style="thin", color="CBD5E1")
                    cell.border = Border(left=Side(border_style="thin", color="CBD5E1"),
                                         right=Side(border_style="thin", color="CBD5E1"),
                                         top=top_s,
                                         bottom=Side(border_style="thin", color="CBD5E1"))
            ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
            ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col+1)

        add_kpi(ws_dash, 1, "TOTAL DE ORDENS", f'=COUNTA(\'{nome_aba_base}\'!F5:F{n_rows+4})')
        add_kpi(ws_dash, 3, "TAT MÉDIO GERAL (DIAS)", f'=AVERAGE(\'{nome_aba_base}\'!C5:C{n_rows+4})', "0.0")
        add_kpi(ws_dash, 5, "ORDENS EM ABERTO", f'=COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "<>CONCLUÍDA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "FINALIZADA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "CANCELADA")')
        add_kpi(ws_dash, 7, "OS COM PEÇAS", f'=COUNTIF(\'{nome_aba_base}\'!E5:E{n_rows+4}, "SIM")')
        add_kpi(ws_dash, 9, "REINCIDÊNCIAS / CRÍTICAS", f'=COUNTIF(\'{nome_aba_base}\'!B5:B{n_rows+4}, "CRÍTICA")')

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.write("📊 **Planilha Consolidada da Rede (Geral):**")
        buf_geral = gerar_excel_completo(df_sorted, "Controle Operacional Geral")
        st.download_button(
            label="📥 Baixar Excel Completo (Todas as Consultoras)",
            data=buf_geral,
            file_name="Gestao_de_Rede_Consolidado_Geral.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with d_col2:
        if sel_consultora != "TODAS":
            st.write(f"👩‍💼 **Planilha Específica ({sel_consultora}):**")
            buf_cons = gerar_excel_completo(df_filtrado, f"Base {sel_consultora[:20]}")
            st.download_button(
                label=f"📥 Baixar Apenas OS de {sel_consultora}",
                data=buf_cons,
                file_name=f"Gestao_de_Rede_{sel_consultora.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Selecione uma consultora específica na barra lateral para liberar o download individual da carteira dela.")
else:
    st.info("👆 Por favor, envie o relatório de serviços acima para iniciar o painel.")
