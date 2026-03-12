import re
import streamlit as st
import pandas as pd
import math
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta, time

# ==========================================
# 1. MOTOR DE CÁLCULO
# ==========================================
def arredondar_meio_fte(n):
    if n <= 0:
        return 0
    # Arredonda para o próximo múltiplo de 0.5 (ex: 1.1 -> 1.5)
    return math.ceil(n * 2) / 2


def extrair_dados_inteligente(df):
    colunas = df.columns.tolist()
    chaves_nome = ['atividade', 'operação', 'operacao', 'nome', 'task']
    chaves_tempo = ['tempo', 'ciclo', 'minutos', 'tc']
    chaves_precedencia = ['precedência', 'precedencia', 'antecessora']

    col_n, col_t, col_p = None, None, None

    for c in colunas:
        c_low = str(c).lower()
        if not col_n and any(k in c_low for k in chaves_nome):
            col_n = c
        if not col_t and any(k in c_low for k in chaves_tempo):
            col_t = c
        if not col_p and any(k in c_low for k in chaves_precedencia):
            col_p = c

    return col_n or colunas[0], col_t or colunas[1], col_p


def ordenar_topologicamente(df):
    grafo = {row['Atividade']: set(row['Precedência']) for _, row in df.iterrows()}
    ordem = []

    while grafo:
        livres = [n for n, deps in grafo.items() if not deps]

        if not livres:
            raise ValueError("Existe um ciclo nas precedências.")

        for n in livres:
            ordem.append(n)
            grafo.pop(n)

        for deps in grafo.values():
            deps.difference_update(livres)

    return df.set_index("Atividade").loc[ordem].reset_index()


def ajustar_com_pausas(dt_inicio, duracao_min, turno_inicio, turno_fim, pausas):
    atual = dt_inicio
    minutos_restantes = duracao_min

    while minutos_restantes > 0:
        fim_turno_hoje = datetime.combine(atual.date(), turno_fim)

        if atual >= fim_turno_hoje:
            atual = datetime.combine(atual.date() + timedelta(days=1), turno_inicio)
            continue

        em_pausa = False
        for p_ini, p_fim in pausas:
            pausa_inicio = datetime.combine(atual.date(), p_ini)
            pausa_fim = datetime.combine(atual.date(), p_fim)

            if pausa_inicio <= atual < pausa_fim:
                atual = pausa_fim
                em_pausa = True
                break

        if em_pausa:
            continue

        proximo_evento = fim_turno_hoje

        for p_ini, _ in pausas:
            pausa_inicio = datetime.combine(atual.date(), p_ini)
            if atual < pausa_inicio < proximo_evento:
                proximo_evento = pausa_inicio

        trabalho = min(
            minutos_restantes,
            (proximo_evento - atual).total_seconds() / 60
        )

        atual += timedelta(minutes=trabalho)
        minutos_restantes -= trabalho

    return atual


def validar_precedencias(df):
    atividades = set(df['Atividade'])
    for _, row in df.iterrows():
        for p in row['Precedência']:
            if p not in atividades:
                raise ValueError(f"Precedência inválida: {p}")


def calcular_cronograma_temporal(df, n_dias, start_dt, t_ini, t_fim, pausas, vol_alvo):

    validar_precedencias(df)
    df = ordenar_topologicamente(df)

    data_limite = start_dt + timedelta(days=n_dias)

    cronograma = []

    # criar FTE reais por operação (derivados do balanceamento)
    recursos = {}

    for _, row in df.iterrows():

        op = row["Atividade"]

        # número de FTE vindo do balanceamento
        fte_balanceamento = row["FTE_Necessario"]

        # converter para recursos físicos
        fte_reais = max(1, math.ceil(fte_balanceamento))

        recursos[op] = []

        for i in range(fte_reais):
            recursos[op].append({
                "id": f"FTE_{op}_{i + 1}",
                "operacao": op,
                "livre_em": start_dt
            })

    obra_id = 1

    while True:

        tempos_fim = {}
        tarefas_obra = []
        obra_viavel = True

        for _, row in df.iterrows():

            op = row["Atividade"]

            if row["Precedência"]:
                liberacao = max(tempos_fim[p] for p in row["Precedência"])
            else:
                liberacao = start_dt

            # escolher FTE mais cedo disponível
            recurso = min(recursos[op], key=lambda r: r["livre_em"])

            inicio = max(liberacao, recurso["livre_em"])

            fim = ajustar_com_pausas(
                inicio,
                row["Tempo Ciclo (min)"],
                t_ini,
                t_fim,
                pausas
            )

            # parar exatamente no horizonte da simulação
            if inicio >= data_limite:
                obra_viavel = False
                break

            tempos_fim[op] = fim
            recurso["livre_em"] = fim

            tarefas_obra.append({
                "Obra": f"Obra {obra_id}",
                "Task": recurso["id"],
                "Start": inicio,
                "Finish": fim,
                "Operacao": op
            })

        if not obra_viavel:
            break

        cronograma.extend(tarefas_obra)
        obra_id += 1

    return cronograma

# ==========================================
# 2. UI
# ==========================================
AZUL_KPMG = '#00338D'
VERMELHO_KPMG = '#E30513'
PRETO = '#000000'
CYAN_WIP = '#75D1E0'

st.set_page_config(page_title="KPMG VSM Tool", layout="wide")
st.markdown(f'<h1 style="color:{AZUL_KPMG};">KPMG VSM Tool</h1>', unsafe_allow_html=True)

# SIDEBAR
st.sidebar.header("⚙️ Definições")

data_ini = st.sidebar.date_input("Início", datetime.now())
h_ini = st.sidebar.time_input("Hora Entrada", time(8, 0))
h_fim = st.sidebar.time_input("Hora Saída", time(17, 0))

n_dias = st.sidebar.number_input("# Dias a simular", 1, 30, 1)
vol_alvo = st.sidebar.number_input("Objetivo Output / Dia", min_value=1, value=50)

if 'pausas' not in st.session_state:
    st.session_state.pausas = []

st.sidebar.subheader("☕ Gestão de Pausas")

c_p1, c_p2 = st.sidebar.columns(2)
p_ini_i = c_p1.time_input("Início Pausa", time(12, 0))
p_fim_i = c_p2.time_input("Fim Pausa", time(13, 0))

if st.sidebar.button("➕ Adicionar Pausa"):
    st.session_state.pausas.append((p_ini_i, p_fim_i))
    st.rerun()

if st.session_state.pausas:
    for i, (p_s, p_e) in enumerate(st.session_state.pausas):
        st.sidebar.caption(f"{i+1}. {p_s.strftime('%H:%M')} - {p_e.strftime('%H:%M')}")

    if st.sidebar.button("🗑️ Limpar Pausas"):
        st.session_state.pausas = []
        st.rerun()

# CARGA DADOS
if 'df_tarefas' not in st.session_state:
    st.session_state.df_tarefas = pd.DataFrame({
        "Atividade": pd.Series(dtype=str),
        "Tempo Ciclo (min)": pd.Series(dtype=float),
        "Precedência": pd.Series(dtype=object)
    })

metodo = st.radio("Carga de Dados:", ["Manual", "Excel"], horizontal=True)

if metodo == "Excel":
    f = st.file_uploader("Upload", type=["xlsx"])
    if f:
        df_b = pd.read_excel(f)
        col_n, col_t, col_p = extrair_dados_inteligente(df_b)

        st.session_state.df_tarefas = pd.DataFrame({
            "Atividade": df_b[col_n].astype(str),
            "Tempo Ciclo (min)": pd.to_numeric(df_b[col_t], errors='coerce').fillna(0.0),
            "Precedência": df_b[col_p].apply(
                lambda x: [i.strip() for i in str(x).split(',')]
                if pd.notnull(x) and str(x).strip() != "" else []
            ) if col_p else [[]] * len(df_b)
        })

# ==========================================
# EDITOR DE OPERAÇÕES (estável + precedências dinâmicas)
# ==========================================

with st.form("editor_operacoes_form"):

    # lista dinâmica de operações existentes
    atividades_disponiveis = (
        st.session_state.df_tarefas["Atividade"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    df_editado = st.data_editor(
        st.session_state.df_tarefas,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Precedência": st.column_config.MultiselectColumn(
                "Precedência",
                options=atividades_disponiveis
            )
        },
        key="editor_operacoes"
    )

    guardar_operacoes = st.form_submit_button("💾 Gerar Precedências")

if guardar_operacoes:
    df_editado["Precedência"] = df_editado["Precedência"].apply(
        lambda x: x if isinstance(x, list) else []
    )

    st.session_state.df_tarefas = df_editado.copy()
    st.success("Operações atualizadas.")

# ==========================================
# 3. ANALYTICS
# ==========================================
# Criar uma "memória" para o cálculo não desaparecer
if 'calculo_feito' not in st.session_state:
    st.session_state.calculo_feito = False

if st.button("🚀 RUN", use_container_width=True):
    st.session_state.calculo_feito = True

# Só entra aqui se o botão tiver sido clicado alguma vez
if st.session_state.calculo_feito:

    df_v = df_editado[df_editado["Atividade"].str.strip() != ""].copy()

    # garantir que precedências são sempre listas
    df_v["Precedência"] = df_v["Precedência"].apply(
        lambda x: x if isinstance(x, list) else []
    )

    if not df_v.empty:

        t_util = (
            (datetime.combine(datetime.today(), h_fim) -
             datetime.combine(datetime.today(), h_ini)).total_seconds() / 60
        ) - sum([
            (datetime.combine(datetime.today(), p[1]) -
             datetime.combine(datetime.today(), p[0])).total_seconds() / 60
            for p in st.session_state.pausas
        ])

        takt = t_util / vol_alvo if vol_alvo > 0 else 0

        df_v['FTE_Necessario'] = (
            (df_v['Tempo Ciclo (min)'] / takt).apply(arredondar_meio_fte)  # Nome atualizado aqui
            if takt > 0 else 0
        )

        # --- CÁLCULOS PARA OS CARDS ---
        # 1. Somatório dos WIPs blindados (Obras físicas)
        wip_total_obras = 0
        for _, row in df_v.iterrows():
            tc_atual = row['Tempo Ciclo (min)']
            fte_necessario = row['FTE_Necessario']
            wip_base = math.ceil(fte_necessario)
            wip_protecao = 0
            sucessores = df_v[df_v['Precedência'].apply(
                lambda p: row['Atividade'] in p if isinstance(p, list) else False
            )]
            for _, suc in sucessores.iterrows():
                tc_suc = suc['Tempo Ciclo (min)']
                fte_suc = suc['FTE_Necessario']
                if tc_suc > 0 and fte_suc > 0:
                    cadencia_sucessor = tc_suc / fte_suc
                    necessario = math.ceil(tc_atual / cadencia_sucessor)
                    wip_protecao = max(wip_protecao, necessario)
            wip_total_obras += max(wip_base, wip_protecao)

        # 2. Somatório de FTEs (Recursos humanos do balanceamento)
        total_fte_necessario = df_v['FTE_Necessario'].sum()

        st.markdown("---")

        # Ajuste para 3 colunas
        c1, c2, c3 = st.columns(3)

        c1.markdown(
            f'<div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;">'
            f'<p style="color:white;margin:0;font-size:14px;">⏱️ TAKT TIME</p>'
            f'<h2 style="color:white;margin:0;">{takt:.2f} min/peça</h2></div>',
            unsafe_allow_html=True
        )

        c2.markdown(
            f'<div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;">'
            f'<p style="color:white;margin:0;font-size:14px;">⚙️ WIP TOTAL (ARRANQUE)</p>'
            f'<h2 style="color:white;margin:0;">{wip_total_obras:.0f} Obras</h2></div>',
            unsafe_allow_html=True
        )

        # Novo Card: Total de FTE
        c3.markdown(
            f'<div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;">'
            f'<p style="color:white;margin:0;font-size:14px;">👥 TOTAL FTE NECESSÁRIO</p>'
            f'<h2 style="color:white;margin:0;">{total_fte_necessario:.1f} FTE</h2></div>',
            unsafe_allow_html=True
        )

        # 1. REORDENAÇÃO DAS ABAS (Mapa agora em primeiro)
        tab_mapa, tab_bal, tab_gantt = st.tabs(["🌲 Mapa de Precedências & WIP", "📊 Balanceamento", "📅 Gantt"])

        with tab_mapa:
            st.subheader("Fluxo Operacional, TC & WIP Inicial")

            # --- 1. CÁLCULO DE POSIÇÕES (Dinamizado) ---
            niveis = {}
            while len(niveis) < len(df_v):
                for _, row in df_v.iterrows():
                    atv = row["Atividade"]
                    preds = row["Precedência"]
                    if atv in niveis: continue
                    if not preds:
                        niveis[atv] = 0
                    elif all(p in niveis for p in preds):
                        niveis[atv] = max(niveis[p] for p in preds) + 1

            niveis_dict = {}
            for atv, nivel in niveis.items():
                niveis_dict.setdefault(nivel, []).append(atv)

            pos = {}
            # Espaçamento aumentado para dar "ar" em ecrãs pequenos
            spacing_x, spacing_y = 6, 3.5

            for nivel, atividades in niveis_dict.items():
                atividades = sorted(atividades)
                n = len(atividades)
                for i, atv in enumerate(atividades):
                    pos[atv] = (nivel * spacing_x, (i - (n - 1) / 2) * spacing_y)

            fig_net = go.Figure()

            # --- 2. SETAS DE FLUXO ---
            for _, row in df_v.iterrows():
                destino = row["Atividade"]
                for origem in row["Precedência"]:
                    if origem in pos:
                        x0, y0 = pos[origem]
                        x1, y1 = pos[destino]
                        fig_net.add_annotation(
                            x=x1 - 1.2, y=y1, ax=x0 + 1.2, ay=y0,
                            xref="x", yref="y", axref="x", ayref="y",
                            showarrow=True, arrowhead=3, arrowsize=1.2, arrowwidth=2, arrowcolor=PRETO
                        )

                        # --- 3. DESENHO DAS CAIXAS (FIX: Sem negrito no WIP) ---
                        num_atividades = len(df_v)
                        marker_size = 80 if num_atividades < 10 else 60

                        for atv, (x, y) in pos.items():
                            row_data = df_v[df_v['Atividade'] == atv]
                            tc_atual = row_data['Tempo Ciclo (min)'].values[0]
                            fte_atual = row_data['FTE_Necessario'].values[0]

                            # Lógica de Proteção para WIP
                            wip_base = math.ceil(fte_atual)
                            wip_protecao = 0
                            sucessores = df_v[df_v['Precedência'].apply(lambda p: atv in p)]
                            for _, suc in sucessores.iterrows():
                                if suc['Tempo Ciclo (min)'] > 0 and suc['FTE_Necessario'] > 0:
                                    cadencia_suc = suc['Tempo Ciclo (min)'] / suc['FTE_Necessario']
                                    wip_protecao = max(wip_protecao, math.ceil(tc_atual / cadencia_suc))

                            wip_final = max(wip_base, wip_protecao)

                            # Caixa Azul KPMG (Mantém o nome em negrito para destaque)
                            fig_net.add_trace(go.Scatter(
                                x=[x], y=[y], mode="markers+text",
                                marker=dict(size=marker_size, symbol="square", color=AZUL_KPMG,
                                            line=dict(width=1, color=PRETO)),
                                text=f"<b>{atv}</b>",
                                textfont=dict(color="white", size=12),
                                textposition="middle center",
                                hoverinfo="skip"
                            ))

                            # Info fora da caixa: TC e WIP (Sem a tag <b> no WIP)
                            fig_net.add_trace(go.Scatter(
                                x=[x], y=[y - (marker_size / 45)],
                                mode="text",
                                text=f"⏱️ {tc_atual} min<br>WIP Inicial: {wip_final}",
                                textfont=dict(size=11, color="black"),
                                textposition="bottom center",
                                hoverinfo="skip"
                            ))

            # Formatação final do layout
            fig_net.update_layout(
                showlegend=False, plot_bgcolor="white", height=500,
                margin=dict(l=20, r=20, t=20, b=20), autosize=True,
                xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
            )

            # Forçamos a responsividade
            st.plotly_chart(fig_net, use_container_width=True, config={'responsive': True})

        with tab_bal:
            # 1. CRIAR O OBJETO DO GRÁFICO
            fig_bal = make_subplots(specs=[[{"secondary_y": True}]])

            # 2. ADICIONAR BARRAS (Capacidade)
            fig_bal.add_trace(go.Bar(
                x=df_v['Atividade'],
                y=(t_util * df_v['FTE_Necessario'] / df_v['Tempo Ciclo (min)']),
                name="Capacidade/Dia",
                marker_color=AZUL_KPMG,
                opacity=0.6
            ), secondary_y=False)

            # 3. ADICIONAR LINHA (FTE com Negrito e Destaque)
            fig_bal.add_trace(go.Scatter(
                x=df_v['Atividade'],
                y=df_v['FTE_Necessario'],
                name="FTE Necessário",
                mode='lines+markers+text',
                text=df_v['FTE_Necessario'],
                textposition="top center",
                texttemplate='<b>%{text}</b>',
                textfont=dict(family="Arial", size=14, color=VERMELHO_KPMG),
                marker=dict(size=12, color=VERMELHO_KPMG),
                line=dict(color=VERMELHO_KPMG, width=3),
                cliponaxis=False
            ), secondary_y=True)

            # 4. CONFIGURAR LAYOUT (Sem título e legenda ao meio)
            max_fte = df_v['FTE_Necessario'].max() if not df_v.empty else 5

            fig_bal.update_layout(
                title_text="",  # Subtítulo eliminado conforme pedido
                hovermode="x unified",
                margin=dict(t=50, b=100),  # Espaço ajustado para a legenda em baixo

                # Configuração da Legenda ao Centro
                legend=dict(
                    orientation="h",  # Horizontal
                    yanchor="top",
                    y=-0.2,  # Posiciona abaixo do eixo X
                    xanchor="center",
                    x=0.5  # Centraliza horizontalmente
                ),

                yaxis2=dict(
                    range=[0, max_fte * 1.4],
                    overlaying='y',
                    side='right',
                    title="<b>FTE Necessário</b>",
                    showgrid=False
                ),
                yaxis=dict(
                    title="<b>Capacidade (Obras/Dia)</b>",
                    showgrid=True,
                    gridcolor='lightgrey'
                ),
                xaxis=dict(title="<b>Operações</b>")
            )

            # Linha de Objectivo
            fig_bal.add_hline(y=vol_alvo, line_dash="dash", line_color=PRETO, annotation_text="Target")

            st.plotly_chart(fig_bal, use_container_width=True)

        with tab_gantt:
            # 1. CÁLCULO DO CRONOGRAMA
            cronograma_base = calcular_cronograma_temporal(
                df_v, n_dias, datetime.combine(data_ini, h_ini),
                h_ini, h_fim, st.session_state.pausas, vol_alvo
            )

            if cronograma_base:
                df_res = pd.DataFrame(cronograma_base)
                ops_unicas = sorted(df_res["Operacao"].unique())

                with st.expander("🛒 Filtrar Operações no Gantt", expanded=True):
                    sel_all = st.toggle("Selecionar Tudo", value=True, key="master_filter_final")
                    cols = st.columns(3)
                    vistos = {op: cols[i % 3].checkbox(op, value=sel_all, key=f"f_{op}_{sel_all}")
                              for i, op in enumerate(ops_unicas)}

                op_selecionadas = [op for op, marcado in vistos.items() if marcado]
                df_filtrado = df_res[df_res["Operacao"].isin(op_selecionadas)].copy()

                if not df_filtrado.empty:
                    # Fragmentação para garantir visibilidade correta com rangebreaks
                    fragmentos = []
                    for _, row in df_filtrado.iterrows():
                        temp_start = row['Start']
                        while temp_start < row['Finish']:
                            work_end = datetime.combine(temp_start.date(), h_fim)
                            seg_end = min(row['Finish'], work_end)

                            found_pause = False
                            for p_ini, p_fim in st.session_state.pausas:
                                p_dt_i, p_dt_f = datetime.combine(temp_start.date(), p_ini), datetime.combine(
                                    temp_start.date(), p_fim)
                                if temp_start < p_dt_i < seg_end:
                                    fragmentos.append({**row, 'Start': temp_start, 'Finish': p_dt_i})
                                    temp_start = p_dt_f
                                    found_pause = True;
                                    break

                            if not found_pause:
                                if temp_start < seg_end:
                                    fragmentos.append({**row, 'Start': temp_start, 'Finish': seg_end})
                                if seg_end == work_end and seg_end < row['Finish']:
                                    temp_start = datetime.combine(temp_start.date() + timedelta(days=1), h_ini)
                                else:
                                    temp_start = row['Finish']

                    df_plot = pd.DataFrame(fragmentos)
                    df_plot['Recurso'] = df_plot['Task'].str.replace('FTE_', '').str.replace('_', ' - Posto ')


                    # Ordenação Inteligente (Posto 1, Posto 2, etc.)
                    def natural_sort_key(s):
                        return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


                    recursos_ordenados = sorted(df_plot['Recurso'].unique(), key=natural_sort_key)

                    # Criação do Gantt
                    fig_g = px.timeline(
                        df_plot, x_start="Start", x_end="Finish", y="Recurso", color="Obra",
                        category_orders={"Recurso": recursos_ordenados}
                    )

                    # Ajustes de Visualização (Inversão do Y e Ocultar Noites)
                    fig_g.update_yaxes(autorange="reversed")
                    fig_g.update_xaxes(rangebreaks=[dict(bounds=[h_fim.hour, h_ini.hour], pattern="hour")])

                    # Adição visual das Pausas
                    if st.session_state.pausas:
                        for d in range((df_plot['Finish'].max().date() - df_plot['Start'].min().date()).days + 1):
                            dia = df_plot['Start'].min().date() + timedelta(days=d)
                            for p_ini, p_fim in st.session_state.pausas:
                                t_i, t_f = datetime.combine(dia, p_ini), datetime.combine(dia, p_fim)
                                fig_g.add_vrect(x0=t_i, x1=t_f, fillcolor="black", opacity=1, layer="above",
                                                line_width=0)
                                fig_g.add_annotation(x=t_i + (t_f - t_i) / 2, y=0.5, yref="paper", text="PAUSA",
                                                     font=dict(color="white", size=10), showarrow=False, textangle=-90)

                    fig_g.update_layout(height=max(400, len(recursos_ordenados) * 35))
                    st.plotly_chart(fig_g, use_container_width=True)