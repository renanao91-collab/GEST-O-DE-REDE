import os
import io
import time
from datetime import datetime, date
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

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
# CONFIGURAÇÃO DE DATA DE CONSULTA NA BARRA LATERAL
# -------------------------------------------------------------
st.sidebar.markdown(f"<h3 style='color:{PANASONIC_BLUE};'>📅 Referência da Consulta</h3>", unsafe_allow_html=True)
data_referencia = st.sidebar.date_input(
    "Data de Corte / Consulta:",
    value=date.today(),
    help="Para ordens ainda NÃO finalizadas (sem data de conclusão), o tempo decorrido de atendimento é calculado com base nesta data."
)
ts_referencia = pd.Timestamp(data_referencia).normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)

# -------------------------------------------------------------
# PROCESSAMENTO DE DADOS
# -------------------------------------------------------------
if file_servico is not None:
    with st.spinner("Processando ordens de serviço da rede Panasonic com cálculo dinâmico de TAT..."):
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

        # CÁLCULO PRECISO DO TAT:
        # Se tem Conclusão/Finalização: usa Data Conclusão - Data Criação
        # Se NÃO tem Conclusão (Aberta/Pendente): usa Data de Referência (Consulta/Hoje) - Data Criação
        dt_fim = df_sorted['dt_conclusao'].fillna(df_sorted['dt_finalizacao'])
        
        tat_dias = []
        is_reparado = []
        is_aberta = []
        
        for i, r in df_sorted.iterrows():
            st_val = str(r.get('Status Atual', '')).strip().upper()
            reparado = st_val in ['CONCLUÍDA', 'FINALIZADA']
            aberta = st_val not in ['CONCLUÍDA', 'FINALIZADA', 'CANCELADA']
            
            is_reparado.append(reparado)
            is_aberta.append(aberta)
            
            if pd.notna(r['dt_criacao']):
                if pd.notna(dt_fim[i]):
                    diff_seconds = (dt_fim[i] - r['dt_criacao']).total_seconds()
                    tat_dias.append(max(0, int(diff_seconds // 86400)))
                else:
                    # Ordem em aberto: calcula com a data da consulta
                    diff_seconds = (ts_referencia - r['dt_criacao']).total_seconds()
                    tat_dias.append(max(0, int(diff_seconds // 86400)))
            else:
                tat_dias.append(0)
                
        df_sorted['TAT_Dias'] = tat_dias
        df_sorted['Is_Reparado'] = is_reparado
        df_sorted['Is_Aberta'] = is_aberta

        # Faixas de Atendimento
        def get_faixa(d):
            if d <= 10: return "Até 10 dias (No Target)"
            elif d <= 30: return "11 a 30 dias (Alerta)"
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
    # BARRA LATERAL: FILTROS OPERACIONAIS
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
    faixas_opcoes = ["Até 10 dias (No Target)", "11 a 30 dias (Alerta)", "> 30 dias (Estouro CDC)"]
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
    
    # Base válida (desconsiderando canceladas)
    df_validas = df_filtrado[df_filtrado['Status Atual'] != 'CANCELADA']
    total_validas = len(df_validas)
    
    # 1. Percentual com até 10 dias
    qtd_ate_10 = (df_validas['TAT_Dias'] <= 10).sum()
    pct_ate_10 = (qtd_ate_10 / total_validas * 100) if total_validas > 0 else 0
    delta_meta_10 = pct_ate_10 - META_10_DIAS
    
    # 2. Percentual com até 30 dias (CDC)
    qtd_ate_30 = (df_validas['TAT_Dias'] <= 30).sum()
    pct_ate_30 = (qtd_ate_30 / total_validas * 100) if total_validas > 0 else 0
    
    # 3. TAT Médio de Reparo (Concluídas/Finalizadas)
    df_reparadas = df_filtrado[df_filtrado['Is_Reparado']]
    tat_medio_reparo = df_reparadas['TAT_Dias'].mean() if len(df_reparadas) > 0 else 0
    
    # 4. Percentual de Reincidentes
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
    st.caption(f"Cálculo com data de referência: **{data_referencia.strftime('%d/%m/%Y')}** (Ordens em aberto calculadas até esta data)")

    # Banner informativo da Meta
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
    em_aberto = df_filtrado[df_filtrado['Is_Aberta']].shape[0]
    com_pecas = df_filtrado[df_filtrado['Tem_Peca']].shape[0]
    criticas = df_filtrado[df_filtrado['Prioridade'] == 'CRÍTICA'].shape[0]

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

    # Agregação por Posto
    df_validas_postos = df_validas.copy()
    df_validas_postos['Ate_10d_Num'] = (df_validas_postos['TAT_Dias'] <= 10).astype(int)
    
    resumo_postos = df_validas_postos.groupby(['Unidade (Digiteam)', 'Consultora Responsável']).agg(
        Total_OS=('Código OS', 'count'),
        OS_Ate_10d=('Ate_10d_Num', 'sum'),
        TAT_Medio=('TAT_Dias', 'mean')
    ).reset_index()

    resumo_postos['Pct_10d'] = (resumo_postos['OS_Ate_10d'] / resumo_postos['Total_OS']) * 100
    resumo_postos['Status_Meta'] = resumo_postos['Pct_10d'].apply(
        lambda x: "✅ Na Meta (≥ 62%)" if x >= META_10_DIAS else "❌ Ofensor (< 62%)"
    )
    # Gap: quantas ordens adicionais em 10 dias seriam necessárias para atingir 62%
    resumo_postos['Gap_OS_Meta'] = resumo_postos.apply(
        lambda r: max(0, int(round((META_10_DIAS / 100.0 * r['Total_OS']) - r['OS_Ate_10d']))),
        axis=1
    )

    # Filtrar apenas os ofensores
    df_ofensores = resumo_postos[resumo_postos['Pct_10d'] < META_10_DIAS].copy()
    
    # Opção de filtro de volume mínimo
    c_filt_of1, c_filt_of2 = st.columns([2, 2])
    with c_filt_of1:
        min_ordens = st.slider(
            "Filtrar postos com no mínimo de ordens:",
            min_value=1,
            max_value=max(10, int(resumo_postos['Total_OS'].max())) if len(resumo_postos) > 0 else 10,
            value=5,
            help="Postos com poucas ordens podem distorcer o percentual. Ajuste para focar nos maiores volumes."
        )
    
    df_ofensores_view = df_ofensores[df_ofensores['Total_OS'] >= min_ordens].sort_values(
        by=['Pct_10d', 'Total_OS'],
        ascending=[True, False]
    )

    with c_filt_of2:
        st.metric(
            "Total de Postos Ofensores Identificados",
            f"{len(df_ofensores_view)} postos",
            f"{len(df_ofensores)} ofensores no total geral",
            delta_color="inverse"
        )

    # Gráfico dos Maiores Ofensores
    if len(df_ofensores_view) > 0:
        col_g_of1, col_g_of2 = st.columns([3, 2])
        with col_g_of1:
            top_15_ofensores = df_ofensores_view.head(15).copy()
            fig_of = px.bar(
                top_15_ofensores,
                x='Pct_10d',
                y='Unidade (Digiteam)',
                orientation='h',
                text='Pct_10d',
                title=f"Top 15 Postos Mais Críticos (% Atendimento ≤ 10 Dias)",
                color='Pct_10d',
                color_continuous_scale='Reds_r'
            )
            fig_of.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            # Linha da Meta de 62%
            fig_of.add_vline(
                x=META_10_DIAS,
                line_dash="dash",
                line_color="#004098",
                annotation_text="Meta 62%",
                annotation_position="top right"
            )
            fig_of.update_layout(
                yaxis={'categoryorder':'total descending'},
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
    # GERADOR DE EXCEL OFICIAL (.XLSX) COM ABA DE OFENSORES
    # -------------------------------------------------------------
    st.divider()
    st.subheader("📥 Exportação de Planilhas Formatadas (Padrão Panasonic)")

    def gerar_excel_completo(df_export, df_ofensores_exp, nome_aba_base="Controle Operacional OS"):
        wb = openpyxl.Workbook()
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

        # 1. ABA DE POSTOS OFENSORES
        ws_ofensores["A1"] = "PANASONIC DO BRASIL - POSTOS OFENSORES DO INDICADOR DE 10 DIAS"
        ws_ofensores["A1"].font = font_title
        ws_ofensores["A2"] = f"Meta Estabelecida: ≥ {META_10_DIAS:.0f}% de atendimentos em até 10 dias corridos | Data de Corte: {data_referencia.strftime('%d/%m/%Y')}"
        ws_ofensores["A2"].font = font_sub

        headers_of = ["Posto Autorizado (Unidade)", "Consultora Responsável", "Total de OS Válidas", "OS em até 10 Dias", "% Atingido", "Meta Target", "Status Meta", "Déficit de OS p/ Meta"]
        for c_idx, h in enumerate(headers_of, 1):
            c = ws_ofensores.cell(row=4, column=c_idx, value=h)
            c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center")
        ws_ofensores.row_dimensions[4].height = 26

        for r_idx, r_data in enumerate(df_ofensores_exp.itertuples(), start=5):
            ws_ofensores.cell(row=r_idx, column=1, value=r_data._1).border = border_cell
            ws_ofensores.cell(row=r_idx, column=2, value=r_data._2).border = border_cell
            
            c3 = ws_ofensores.cell(row=r_idx, column=3, value=r_data.Total_OS)
            c3.number_format = "#,##0"; c3.alignment = Alignment(horizontal="right"); c3.border = border_cell
            
            c4 = ws_ofensores.cell(row=r_idx, column=4, value=r_data.OS_Ate_10d)
            c4.number_format = "#,##0"; c4.alignment = Alignment(horizontal="right"); c4.border = border_cell
            
            c5 = ws_ofensores.cell(row=r_idx, column=5, value=r_data.Pct_10d / 100.0)
            c5.number_format = "0.0%"; c5.alignment = Alignment(horizontal="right"); c5.fill = fill_red; c5.font = font_bold; c5.border = border_cell
            
            c6 = ws_ofensores.cell(row=r_idx, column=6, value=META_10_DIAS / 100.0)
            c6.number_format = "0.0%"; c6.alignment = Alignment(horizontal="right"); c6.border = border_cell
            
            c7 = ws_ofensores.cell(row=r_idx, column=7, value="OFENSOR (< 62%)")
            c7.alignment = Alignment(horizontal="center"); c7.border = border_cell
            
            c8 = ws_ofensores.cell(row=r_idx, column=8, value=r_data.Gap_OS_Meta)
            c8.number_format = "#,##0"; c8.alignment = Alignment(horizontal="right"); c8.border = border_cell
            ws_ofensores.row_dimensions[r_idx].height = 20

        ws_ofensores.freeze_panes = "C5"
        ws_ofensores.auto_filter.ref = f"A4:H{len(df_ofensores_exp)+4}"
        for col_l, w in {"A": 38, "B": 22, "C": 18, "D": 18, "E": 14, "F": 14, "G": 18, "H": 22}.items():
            ws_ofensores.column_dimensions[col_l].width = w

        # 2. ABA BASE
        ws_base["A1"] = f"PANASONIC DO BRASIL - GESTÃO DE REDE & PÓS-VENDA"
        ws_base["A1"].font = font_title
        ws_base["A2"] = f"Acompanhamento Operacional | Data de Corte: {data_referencia.strftime('%d/%m/%Y')} | Ordenado do mais antigo para o mais novo"
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

        data_corte_str = data_referencia.strftime('%Y-%m-%d')
        for i in range(n_rows):
            r = i + 5
            dc = dt_criac_str[i]
            # Fórmula do TAT: Se K (Conclusão) estiver vazia, calcula até a data de corte DATEVALUE("{data_corte_str}")
            f_tat = f'=IF(K{r}="", INT(DATEVALUE("{data_corte_str}")-DATEVALUE(LEFT(J{r},10))), INT(DATEVALUE(LEFT(K{r},10))-DATEVALUE(LEFT(J{r},10))))' if dc else 0
            f_faixa = f'=IF(D{r}<=10, "Até 10 dias (No Target)", IF(D{r}<=30, "11 a 30 dias (Alerta)", "> 30 dias (Estouro CDC)"))'
            f_peca = f'=IF(OR(R{r}<>"", T{r}<>""), "SIM", "NÃO")'
            f_prio = f'=IF(I{r}="CANCELADA", "BAIXA", IF(OR(ISNUMBER(SEARCH("REINCIDENCIA", H{r})), ISNUMBER(SEARCH("REINCIDÊNCIA", H{r})), D{r}>30), "CRÍTICA", IF(AND(E{r}="SIM", I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(AND(D{r}>=3, I{r}<>"FINALIZADA", I{r}<>"CONCLUÍDA"), "ALTA", IF(I{r}="CRIADA", "MÉDIA", "NORMAL")))))'
            f_acao = f'=IF(OR(I{r}="CONCLUÍDA", I{r}="FINALIZADA"), "Encerrado com Sucesso", IF(I{r}="CANCELADA", "Verificar Motivo Cancelamento", IF(D{r}>30, "AÇÃO IMEDIATA: Estouro de Prazo CDC (Risco Jurídico/Troca)", IF(D{r}>10, "ATENÇÃO: Fora da Meta de 10 Dias - Priorizar Atendimento", IF(E{r}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(I{r}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(I{r}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento")))))))'

            ws_base.append([
                cons_resp[i], f_prio, f_faixa, f_tat, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
                cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
            ])

        ws_base.freeze_panes = "F5"
        ws_base.auto_filter.ref = f"A4:W{n_rows+4}"
        for col_l, w in {"A": 22, "B": 20, "C": 24, "D": 14, "E": 16, "F": 18, "G": 14, "H": 28, "I": 18, "J": 18, "K": 18, "L": 22, "M": 18, "N": 35, "O": 20, "P": 8, "Q": 25, "R": 24, "S": 35, "T": 24, "U": 35, "V": 30, "W": 44}.items():
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

        ws_pecas.auto_filter.ref = f"A4:L{count_p+4}"
        ws_pecas.freeze_panes = "D5"

        # 4. ABA DASHBOARD
        ws_dash.views.sheetView[0].showGridLines = True
        ws_dash["A1"] = "PANASONIC DO BRASIL - PAINEL DE CONTROLE E GESTÃO DA REDE"; ws_dash["A1"].font = font_title
        ws_dash["A2"] = f"Meta de 10 Dias: ≥ {META_10_DIAS:.0f}% | Data de Corte: {data_referencia.strftime('%d/%m/%Y')}"; ws_dash["A2"].font = font_sub

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

        add_kpi(ws_dash, 1, "% ATÉ 10 DIAS (META ≥62%)", f'=COUNTIF(\'{nome_aba_base}\'!D5:D{n_rows+4}, "<=10") / MAX(1, COUNTA(\'{nome_aba_base}\'!D5:D{n_rows+4}))', "0.0%")
        add_kpi(ws_dash, 3, "% ATÉ 30 DIAS (CDC)", f'=COUNTIF(\'{nome_aba_base}\'!D5:D{n_rows+4}, "<=30") / MAX(1, COUNTA(\'{nome_aba_base}\'!D5:D{n_rows+4}))', "0.0%")
        add_kpi(ws_dash, 5, "TAT MÉDIO REPARO", f'=AVERAGEIFS(\'{nome_aba_base}\'!D5:D{n_rows+4}, \'{nome_aba_base}\'!I5:I{n_rows+4}, "CONCLUÍDA")', "0.0")
        add_kpi(ws_dash, 7, "% REINCIDENTES", f'=(COUNTIF(\'{nome_aba_base}\'!H5:H{n_rows+4}, "*REINCIDENCIA*") + COUNTIF(\'{nome_aba_base}\'!H5:H{n_rows+4}, "*REINCIDÊNCIA*")) / MAX(1, COUNTA(\'{nome_aba_base}\'!H5:H{n_rows+4}))', "0.00%")
        add_kpi(ws_dash, 9, "TOTAL DE ORDENS", f'=COUNTA(\'{nome_aba_base}\'!F5:F{n_rows+4})\', "#,##0")
        add_kpi(ws_dash, 11, "ORDENS EM ABERTO", f'=COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "<>CONCLUÍDA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "FINALIZADA") - COUNTIF(\'{nome_aba_base}\'!I5:I{n_rows+4}, "CANCELADA")', "#,##0")

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.write("📊 **Planilha Geral Panasonic (Toda a Rede):**")
        buf_geral = gerar_excel_completo(df_sorted, df_ofensores, "Controle Operacional Geral")
        st.download_button(
            label="📥 Baixar Excel Geral (Com Aba de Ofensores)",
            data=buf_geral,
            file_name=f"Panasonic_Gestao_Rede_Consolidado_{data_referencia.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with d_col2:
        if sel_consultoras:
            lbl_cons = ", ".join(sel_consultoras[:2]) + ("..." if len(sel_consultoras) > 2 else "")
            st.write(f"👩‍💼 **Carteira Selecionada ({lbl_cons}):**")
            df_ofensores_sel = df_ofensores_view if len(df_ofensores_view) > 0 else df_ofensores
            buf_cons = gerar_excel_completo(df_filtrado, df_ofensores_sel, "Base Consultoras Filtradas")
            st.download_button(
                label=f"📥 Baixar Excel da(s) Consultora(s) Selecionada(s)",
                data=buf_cons,
                file_name=f"Panasonic_Gestao_Rede_Filtradas_{data_referencia.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("💡 Selecione uma ou mais consultoras na barra lateral para habilitar o download filtrado da carteira delas.")
else:
    st.info("👆 Por favor, envie o relatório de ordens de serviço acima para iniciar o painel.")
