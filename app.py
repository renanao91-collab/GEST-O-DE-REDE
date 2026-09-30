import os
import io
import time
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
import streamlit as st

st.set_page_config(
    page_title="Gerador de Acompanhamento - Gestão de Rede",
    page_icon="📊",
    layout="wide"
)

# Estilos e Cabeçalho do App
st.title("📊 Gestão de Rede - Processador de Relatório de Serviço")
st.markdown("""
Faça o upload do relatório extraído do sistema em formato **CSV** ou **Excel (.xlsx)**. 
O aplicativo processará os dados e gerará automaticamente a planilha completa com:
- **Ordenação Temporal:** Do chamado mais antigo para o mais novo
- **Cálculo de TAT:** Tempo médio/decorrido dinâmico
- **Rastreio de Peças:** Identificação de peças lançadas e trocadas
- **Matriz de Priorização:** Crítica, Alta, Média, Normal e Baixa
- **Dashboard Executivo:** Indicadores consolidados (KPIs) e gráficos
""")

uploaded_file = st.file_uploader("Arraste e solte o arquivo aqui ou clique para selecionar", type=["csv", "xlsx"])

def processar_dados(file):
    # 1. Leitura do arquivo
    nome_arquivo = file.name
    if nome_arquivo.endswith('.csv'):
        try:
            df = pd.read_csv(file)
        except Exception:
            file.seek(0)
            df = pd.read_csv(file, sep=';')
    else:
        df = pd.read_excel(file)

    # Limpeza de cabeçalhos
    df.columns = [c.replace('\ufeff', '').replace('"', '').strip() for c in df.columns]

    # Tratamento de datas
    df['dt_criacao'] = pd.to_datetime(df['Data de Criação'], errors='coerce')
    df['dt_conclusao'] = pd.to_datetime(df.get('Data de Conclusão', None), errors='coerce')
    df['dt_finalizacao'] = pd.to_datetime(df.get('Data de Finalização', None), errors='coerce')

    # Ordenar da mais antiga para a mais nova
    df_sorted = df.sort_values(by='dt_criacao', ascending=True, na_position='last').reset_index(drop=True)
    n_rows = len(df_sorted)

    # Criar Workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_dash = wb.create_sheet(title="Dashboard Gestão")
    ws_base = wb.create_sheet(title="Controle Operacional OS")
    ws_pecas = wb.create_sheet(title="Controle de Peças")

    # Estilos
    NAVY = "1B365D"
    WHITE = "FFFFFF"
    BORDER_GRAY = "D1D5DB"
    font_title = Font(name="Segoe UI", size=15, bold=True, color=NAVY)
    font_sub = Font(name="Segoe UI", size=9, italic=True, color="4B5563")
    font_header = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
    fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

    # --- ABA 2: BASE OPERACIONAL ---
    ws_base["A1"] = "GESTÃO DE REDE - ACOMPANHAMENTO OPERACIONAL DE ORDENS DE SERVIÇO"
    ws_base["A1"].font = font_title
    ws_base["A2"] = "Ordenado do chamado mais antigo para o mais novo | Métricas de TAT e Rastreio de Peças"
    ws_base["A2"].font = font_sub

    headers_base = [
        "Prioridade Operacional", "TAT Atual (Dias)", "Faixa de TAT", "Tem Peça Lançada?",
        "Código da OS", "Ticket", "Tipo de Demanda", "Status Atual", "Data de Criação",
        "Data de Conclusão", "Categoria do Produto", "Modelo do Produto", "Unidade Técnica (Rede)",
        "Região / Cidade", "UF", "Técnico de Campo", "Código Peça Lançada", "Descrição Peça Lançada",
        "Código Peça Trocada", "Descrição Peça Trocada", "Sintoma / Defeito Relatado", "Ação Operacional Recomendada"
    ]
    for c_idx, h in enumerate(headers_base, 1):
        c = ws_base.cell(row=4, column=c_idx, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_base.row_dimensions[4].height = 26

    # Vetores de dados
    cod_os = df_sorted['Código OS'].astype(str).tolist()
    ticket = df_sorted['Número do Ticket'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).tolist() if 'Número do Ticket' in df_sorted else ['']*n_rows
    demanda = df_sorted['Tipo de Demanda'].fillna('').astype(str).tolist() if 'Tipo de Demanda' in df_sorted else ['']*n_rows
    status = df_sorted['Status Atual'].fillna('').astype(str).tolist() if 'Status Atual' in df_sorted else ['']*n_rows
    dt_criac_str = [d.strftime('%Y-%m-%d %H:%M') if pd.notna(d) else '' for d in df_sorted['dt_criacao']]
    dt_concl_comb = df_sorted['dt_conclusao'].fillna(df_sorted['dt_finalizacao'])
    dt_concl_str = [d.strftime('%Y-%m-%d %H:%M') if pd.notna(d) else '' for d in dt_concl_comb]
    cat = df_sorted['Categoria do Produto'].fillna('').astype(str).tolist() if 'Categoria do Produto' in df_sorted else ['']*n_rows
    mod = df_sorted['Modelo do Produto'].fillna('').astype(str).tolist() if 'Modelo do Produto' in df_sorted else ['']*n_rows
    unid = df_sorted['Unidade (Digiteam)'].fillna('').astype(str).tolist() if 'Unidade (Digiteam)' in df_sorted else ['']*n_rows
    cid = df_sorted['Cidade'].fillna('').astype(str).tolist() if 'Cidade' in df_sorted else ['']*n_rows
    uf = df_sorted['Estado'].fillna('').astype(str).str.upper().tolist() if 'Estado' in df_sorted else ['']*n_rows
    tec = df_sorted['Agente de Campo'].fillna('').astype(str).tolist() if 'Agente de Campo' in df_sorted else ['']*n_rows
    c_pl = df_sorted['Código Peça Lançada'].fillna('').astype(str).str.strip().tolist() if 'Código Peça Lançada' in df_sorted else ['']*n_rows
    d_pl = df_sorted['Peça Lançada'].fillna('').astype(str).str.strip().tolist() if 'Peça Lançada' in df_sorted else ['']*n_rows
    c_pt = df_sorted['Codigo Troca de Peças'].fillna('').astype(str).str.strip().tolist() if 'Codigo Troca de Peças' in df_sorted else ['']*n_rows
    d_pt = df_sorted['Troca de Peças'].fillna('').astype(str).str.strip().tolist() if 'Troca de Peças' in df_sorted else ['']*n_rows
    sint = df_sorted['Sintoma e Defeito'].fillna('').astype(str).str.strip().tolist() if 'Sintoma e Defeito' in df_sorted else ['']*n_rows

    for i in range(n_rows):
        r = i + 5
        dc = dt_criac_str[i]
        f_tat = f'=IF(J{r}="", INT(TODAY()-DATEVALUE(LEFT(I{r},10))), INT(DATEVALUE(LEFT(J{r},10))-DATEVALUE(LEFT(I{r},10))))' if dc else 0
        f_peca = f'=IF(OR(Q{r}<>"", S{r}<>""), "SIM", "NÃO")'
        f_prio = f'=IF(H{r}="CANCELADA", "BAIXA", IF(ISNUMBER(SEARCH("REINCIDENCIA", G{r})), "CRÍTICA", IF(AND(D{r}="SIM", H{r}<>"FINALIZADA", H{r}<>"CONCLUÍDA"), "ALTA", IF(AND(B{r}>=3, H{r}<>"FINALIZADA", H{r}<>"CONCLUÍDA"), "ALTA", IF(H{r}="CRIADA", "MÉDIA", "NORMAL")))))'
        f_faixa = f'=IF(B{r}<=1, "0-1 dia", IF(B{r}<=2, "2 dias", IF(B{r}<=3, "3 dias", "> 3 dias")))'
        f_acao = f'=IF(OR(H{r}="CONCLUÍDA", H{r}="FINALIZADA"), "Encerrado com Sucesso", IF(H{r}="CANCELADA", "Verificar Motivo Cancelamento", IF(D{r}="SIM", "Cobrar Envio/Chegada de Peça na Unidade", IF(H{r}="CRIADA", "Atribuir Técnico e Agendar Atendimento", IF(H{r}="AGENDADA", "Acompanhar Deslocamento Técnico", "Monitorar Atendimento")))))'

        ws_base.append([
            f_prio, f_tat, f_faixa, f_peca, cod_os[i], ticket[i], demanda[i], status[i], dc, dt_concl_str[i],
            cat[i], mod[i], unid[i], cid[i], uf[i], tec[i], c_pl[i], d_pl[i], c_pt[i], d_pt[i], sint[i], f_acao
        ])

    ws_base.freeze_panes = "E5"
    ws_base.auto_filter.ref = f"A4:V{n_rows+4}"
    
    col_widths_base = {
        "A": 22, "B": 14, "C": 13, "D": 16, "E": 18, "F": 14, "G": 28, "H": 18,
        "I": 18, "J": 18, "K": 22, "L": 18, "M": 35, "N": 20, "O": 8, "P": 25,
        "Q": 24, "R": 35, "S": 24, "T": 35, "U": 30, "V": 36
    }
    for col_letter, width in col_widths_base.items():
        ws_base.column_dimensions[col_letter].width = width

    # --- ABA 3: CONTROLE DE PEÇAS ---
    ws_pecas["A1"] = "GESTÃO DE REDE - CONTROLE ESPECÍFICO DE PEÇAS LANÇADAS E APLICADAS"
    ws_pecas["A1"].font = font_title
    headers_p = ["Código da OS", "Ticket", "Status Atual", "Tipo de Demanda", "Unidade Técnica (Rede)", "Cidade / UF", "Código da Peça", "Descrição da Peça", "Tipo de Registro", "Técnico de Campo", "Status Logístico Sugerido"]
    for c_idx, h in enumerate(headers_p, 1):
        c = ws_pecas.cell(row=4, column=c_idx, value=h)
        c.font = font_header; c.fill = fill_header; c.alignment = Alignment(horizontal="center", vertical="center")

    count_p = 0
    for i in range(n_rows):
        if c_pl[i] or d_pl[i] or c_pt[i] or d_pt[i]:
            st_val = status[i]
            cid_uf = f"{cid[i]} / {uf[i]}"
            if c_pl[i] or d_pl[i]:
                st_log = "Instalada / Baixada" if st_val in ["CONCLUÍDA", "FINALIZADA"] else "Pendente Despacho / Chegada na Unidade"
                ws_pecas.append([cod_os[i], ticket[i], st_val, demanda[i], unid[i], cid_uf, c_pl[i], d_pl[i], "Peça Lançada (Aguardando / Em Processamento)", tec[i], st_log])
                count_p += 1
            if c_pt[i] or d_pt[i]:
                st_log = "Peça Trocada - Iniciar Logística Reversa" if st_val in ["CONCLUÍDA", "FINALIZADA"] else "Em Processo de Troca em Campo"
                ws_pecas.append([cod_os[i], ticket[i], st_val, demanda[i], unid[i], cid_uf, c_pt[i], d_pt[i], "Troca de Peça Efetuada em Campo", tec[i], st_log])
                count_p += 1

    ws_pecas.auto_filter.ref = f"A4:K{count_p+4}"
    ws_pecas.freeze_panes = "C5"
    col_widths_pecas = {
        "A": 18, "B": 14, "C": 18, "D": 28, "E": 36, "F": 25,
        "G": 28, "H": 45, "I": 35, "J": 25, "K": 35
    }
    for col_letter, width in col_widths_pecas.items():
        ws_pecas.column_dimensions[col_letter].width = width

    # --- ABA 1: DASHBOARD ---
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash["A1"] = "PAINEL DE CONTROLE E GESTÃO DA REDE DE ATENDIMENTO"
    ws_dash["A1"].font = font_title

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

    add_kpi(ws_dash, 1, "TOTAL DE ORDENS", f'=COUNTA(\'Controle Operacional OS\'!E5:E{n_rows+4})')
    add_kpi(ws_dash, 3, "TAT MÉDIO GERAL (DIAS)", f'=AVERAGE(\'Controle Operacional OS\'!B5:B{n_rows+4})', "0.0")
    add_kpi(ws_dash, 5, "ORDENS EM ABERTO", f'=COUNTIF(\'Controle Operacional OS\'!H5:H{n_rows+4}, "<>CONCLUÍDA") - COUNTIF(\'Controle Operacional OS\'!H5:H{n_rows+4}, "FINALIZADA") - COUNTIF(\'Controle Operacional OS\'!H5:H{n_rows+4}, "CANCELADA")')
    add_kpi(ws_dash, 7, "OS COM PEÇAS", f'=COUNTIF(\'Controle Operacional OS\'!D5:D{n_rows+4}, "SIM")')
    add_kpi(ws_dash, 9, "REINCIDÊNCIAS / CRÍTICAS", f'=COUNTIF(\'Controle Operacional OS\'!A5:A{n_rows+4}, "CRÍTICA")')

    # Salvar em buffer de memória para download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, n_rows, count_p

if uploaded_file is not None:
    st.info("Arquivo recebido! Processando dados operacionais...")
    with st.spinner("Construindo painel executivo, ordenando chamados e calculando métricas de TAT..."):
        t0 = time.time()
        excel_buffer, total_os, total_pecas = processar_dados(uploaded_file)
        tempo_proc = time.time() - t0

    st.success(f"Concluído com sucesso em {tempo_proc:.1f} segundos!")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Ordens Processadas", f"{total_os:,}".replace(",", "."))
    col2.metric("Registros de Peças", f"{total_pecas:,}".replace(",", "."))
    col3.metric("Tempo de Processamento", f"{tempo_proc:.1f}s")

    st.download_button(
        label="📥 Baixar Planilha Pronta (.xlsx)",
        data=excel_buffer,
        file_name="Acompanhamento_Gestao_de_Rede_Consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
