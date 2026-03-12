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


def calcular_cronograma_temporal(df, n_dias, start_dt, t_ini, t_fim, pausas):

    validar_precedencias(df)
    df = ordenar_topologicamente(df)

    data_limite = start_dt + timedelta(days=n_dias)

    cronograma = []
    recursos = {}

    # criar postos físicos
    for _, row in df.iterrows():

        op = row["Atividade"]
        fte = max(1, math.ceil(row["FTE_Necessario"]))

        recursos[op] = []

        for i in range(fte):
            recursos[op].append({
                "id": f"{op}_Posto_{i+1}",
                "livre_em": start_dt
            })

    # ---------------------------------
    # CALCULAR WIP INICIAL
    # ---------------------------------

    wip_por_operacao = {}

    for _, row in df.iterrows():

        tc = row["Tempo Ciclo (min)"]
        fte = row["FTE_Necessario"]

        wip_base = math.ceil(fte)

        wip_protecao = 0

        sucessores = df[df['Precedência'].apply(
            lambda p: row['Atividade'] in p if isinstance(p, list) else False
        )]

        for _, suc in sucessores.iterrows():

            tc_s = suc["Tempo Ciclo (min)"]
            fte_s = suc["FTE_Necessario"]

            if tc_s > 0 and fte_s > 0:

                cad = tc_s / fte_s
                wip_protecao = max(wip_protecao, math.ceil(tc / cad))

        wip_por_operacao[row["Atividade"]] = max(wip_base, wip_protecao)

    # ---------------------------------
    # GERAR WIP INICIAL
    # ---------------------------------

    obra_id = 1
    fila_operacao = {op: [] for op in df["Atividade"]}

    for op in fila_operacao:

        for _ in range(wip_por_operacao[op]):
            fila_operacao[op].append(f"Obra {obra_id}")
            obra_id += 1

    # ---------------------------------
    # SIMULAÇÃO
    # ---------------------------------

    while True:

        houve_trabalho = False

        for _, row in df.iterrows():

            op = row["Atividade"]

            if not fila_operacao[op]:
                continue

            obra = fila_operacao[op].pop(0)

            recurso = min(recursos[op], key=lambda r: r["livre_em"])

            inicio = recurso["livre_em"]

            if inicio >= data_limite:
                return pd.DataFrame(cronograma)

            fim = ajustar_com_pausas(
                inicio,
                row["Tempo Ciclo (min)"],
                t_ini,
                t_fim,
                pausas
            )

            recurso["livre_em"] = fim

            cronograma.append({
                "Obra": obra,
                "Operacao": op,
                "Recurso": recurso["id"],
                "Start": inicio,
                "Finish": fim
            })

            houve_trabalho = True

            # enviar obra para sucessores
            sucessores = df[df['Precedência'].apply(
                lambda p: op in p if isinstance(p, list) else False
            )]

            for _, suc in sucessores.iterrows():
                fila_operacao[suc["Atividade"]].append(obra)

        if not houve_trabalho:
            break

    return pd.DataFrame(cronograma)
# ==========================================
# 2. UI
# ==========================================
AZUL_KPMG = '#00338D'
VERMELHO_KPMG = '#E30513'
PRETO = '#000000'
CYAN_WIP = '#75D1E0'

st.set_page_config(page_title="KPMG VSM Tool", layout="wide")
st.markdown(f'<h1 style="color:{AZUL_KPMG};">KPMG VSM Tool</h1>', unsafe_allow_html=True)

# ==========================================
# 2. UI - SIDEBAR (ORGANIZADA)
# ==========================================
st.sidebar.header("⚙️ Definições Gerais")

# 1. INICIALIZAÇÃO DE ESTADOS (Evita o AttributeError)
if 'turnos' not in st.session_state:
    st.session_state.turnos = []
if 'pausas' not in st.session_state:
    st.session_state.pausas = []

# 2. INPUTS BÁSICOS
data_ini = st.sidebar.date_input("Data de Início", datetime.now(), key="data_principal")
n_dias = st.sidebar.number_input("# Dias a simular", 1, 30, 1, key="sim_dias")
vol_alvo = st.sidebar.number_input("Objetivo Output / Dia", min_value=1, value=50, key="obj_vol")

# 3. GESTÃO DE TURNOS
st.sidebar.markdown("---")
st.sidebar.subheader("🕒 Gestão de Turnos")
c_t1, c_t2 = st.sidebar.columns(2)
h_ini_new = c_t1.time_input("Entrada", time(8, 0), key="new_t_ini")
h_fim_new = c_t2.time_input("Saída", time(16, 0), key="new_t_fim")

if st.sidebar.button("➕ Adicionar Turno", use_container_width=True):
    st.session_state.turnos.append((h_ini_new, h_fim_new))
    st.rerun()

if st.session_state.turnos:
    for i, (t_s, t_e) in enumerate(st.session_state.turnos):
        st.sidebar.caption(f"Turno {i + 1}: {t_s.strftime('%H:%M')} - {t_e.strftime('%H:%M')}")

    if st.sidebar.button("🗑️ Limpar Turnos", key="clear_turnos"):
        st.session_state.turnos = []
        st.rerun()
else:
    st.sidebar.warning("⚠️ Adicione pelo menos um turno.")

# 4. GESTÃO DE PAUSAS
st.sidebar.markdown("---")
st.sidebar.subheader("☕ Gestão de Pausas")
c_p1, c_p2 = st.sidebar.columns(2)
p_ini_i = c_p1.time_input("Início Pausa", time(12, 0), key="new_p_ini")
p_fim_i = c_p2.time_input("Fim Pausa", time(13, 0), key="new_p_fim")

if st.sidebar.button("➕ Adicionar Pausa", use_container_width=True):
    st.session_state.pausas.append((p_ini_i, p_fim_i))
    st.rerun()

if st.session_state.pausas:
    for i, (p_s, p_e) in enumerate(st.session_state.pausas):
        st.sidebar.caption(f"{i+1}. {p_s.strftime('%H:%M')} - {p_e.strftime('%H:%M')}")

    if st.sidebar.button("🗑️ Limpar Pausas", key="clear_pausas"):
        st.session_state.pausas = []
        st.rerun()

# 5. CARGA DADOS
st.sidebar.markdown("---")

# Escondemos a escolha manual, forçando apenas Excel
# metodo = st.radio("Carga de Dados:", ["Manual", "Excel"], horizontal=True) # Comentado para esconder
metodo = "Excel"

if 'df_tarefas' not in st.session_state:
    st.session_state.df_tarefas = pd.DataFrame({
        "Atividade": pd.Series(dtype=str),
        "Tempo Ciclo (min)": pd.Series(dtype=float),
        "Precedência": pd.Series(dtype=object)
    })

if metodo == "Excel":
    st.sidebar.subheader("📁 Importar Processo") # Adicionado para dar contexto já que o rádio sumiu
    f = st.file_uploader("Selecione o ficheiro Excel", type=["xlsx"])
    if f:
        df_b = pd.read_excel(f)
        col_n, col_t, col_p = extrair_dados_inteligente(df_b)
        st.session_state.df_tarefas = pd.DataFrame({
            "Atividade": df_b[col_n].astype(str),
            "Tempo Ciclo (min)": pd.to_numeric(df_b[col_t], errors='coerce').fillna(0.0),
            "Precedência": df_b[col_p].apply(lambda x: [i.strip() for i in str(x).split(',')] if pd.notnull(x) and str(x).strip() != "" else []) if col_p else [[]] * len(df_b)
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
# 3. ANALYTICS (VERSÃO MULTI-TURNO SEM TURNOS PRÉVIOS)
# ==========================================
if 'calculo_feito' not in st.session_state:
    st.session_state.calculo_feito = False

if st.button("🚀 RUN", use_container_width=True):
    st.session_state.calculo_feito = True

if st.session_state.calculo_feito:
    df_v = df_editado[df_editado["Atividade"].str.strip() != ""].copy()

    # TRATAMENTO E CRIAÇÃO PREVENTIVA DA COLUNA
    df_v["Tempo Ciclo (min)"] = pd.to_numeric(df_v["Tempo Ciclo (min)"].astype(str).str.replace(",", "."),
                                              errors="coerce").fillna(0)
    df_v["Precedência"] = df_v["Precedência"].apply(lambda x: x if isinstance(x, list) else [])
    df_v['FTE_Necessario'] = 0.0  # Garante que a coluna existe antes de ser lida

    if not df_v.empty:
        # CÁLCULO DO TEMPO ÚTIL
        tempo_bruto_turnos = 0
        if not st.session_state.turnos:
            st.error("⚠️ ERRO: Por favor, adicione pelo menos um turno na barra lateral para efetuar os cálculos.")
            st.stop()

        for t_ini, t_fim in st.session_state.turnos:
            duracao = (datetime.combine(datetime.today(), t_fim) - datetime.combine(datetime.today(),
                                                                                    t_ini)).total_seconds() / 60
            if duracao < 0: duracao += 1440
            tempo_bruto_turnos += duracao

        total_pausas = sum(
            [(datetime.combine(datetime.today(), p[1]) - datetime.combine(datetime.today(), p[0])).total_seconds() / 60
             for p in st.session_state.pausas])

        t_util = max(0, tempo_bruto_turnos - total_pausas)
        takt = t_util / vol_alvo if vol_alvo > 0 else 0

        # CÁLCULO DO FTE
        if takt > 0:
            df_v['FTE_Necessario'] = df_v['Tempo Ciclo (min)'].apply(lambda x: arredondar_meio_fte(x / takt))

        # CÁLCULOS DE WIP E TPT
        wip_total_obras = 0
        for _, row in df_v.iterrows():
            tc_atual = row['Tempo Ciclo (min)']
            fte_nec = row['FTE_Necessario']  # Agora garantimos que esta coluna existe
            wip_base = math.ceil(fte_nec)
            wip_protecao = 0

            sucessores = df_v[df_v['Precedência'].apply(lambda p: row['Atividade'] in p)]
            for _, suc in sucessores.iterrows():
                if suc['Tempo Ciclo (min)'] > 0 and suc['FTE_Necessario'] > 0:
                    cadencia_suc = suc['Tempo Ciclo (min)'] / suc['FTE_Necessario']
                    wip_protecao = max(wip_protecao, math.ceil(tc_atual / cadencia_suc))

            wip_total_obras += max(wip_base, wip_protecao)

        tpt_total_min = wip_total_obras * takt
        total_fte_necessario = df_v['FTE_Necessario'].sum()

        st.markdown("---")
        # CARDS DE KPI (FTE -> WIP -> TPT -> TAKT)
        st.markdown(f"""
            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 25px;">
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform: uppercase; font-weight: bold; opacity: 0.8;">👤 fte</small><br>
                    <span style="font-size: 24px; font-weight: bold;">{total_fte_necessario:.1f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform: uppercase; font-weight: bold; opacity: 0.8;">⚙️ wip (un.)</small><br>
                    <span style="font-size: 24px; font-weight: bold;">{wip_total_obras:.0f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform: uppercase; font-weight: bold; opacity: 0.8;">🚀 throughput time (min)</small><br>
                    <span style="font-size: 24px; font-weight: bold;">{tpt_total_min:.2f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform: uppercase; font-weight: bold; opacity: 0.8;">⏱️ takt time (min)</small><br>
                    <span style="font-size: 24px; font-weight: bold;">{takt:.2f}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # ABAS DE VISUALIZAÇÃO
        tab_mapa, tab_bal = st.tabs(["🌲 Mapa de Precedências & WIP", "📊 Balanceamento"])

        with tab_mapa:
            st.subheader("Fluxo Operacional, TC & WIP Inicial")

            # --- 1. CÁLCULO DE POSIÇÕES ---
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
            spacing_x, spacing_y = 6, 3.5
            for nivel, atividades in niveis_dict.items():
                atividades = sorted(atividades)
                n = len(atividades)
                for i, atv in enumerate(atividades):
                    pos[atv] = (nivel * spacing_x, (i - (n - 1) / 2) * spacing_y)

            fig_net = go.Figure()

            # --- 2. SETAS ---
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

            # --- 3. CAIXAS E TEXTO NÍTIDO ---
            num_atividades = len(df_v)
            marker_size = 80 if num_atividades < 10 else 60

            for atv, (x, y) in pos.items():
                row_data = df_v[df_v['Atividade'] == atv]
                tc_atual = row_data['Tempo Ciclo (min)'].values[0]
                fte_atual = row_data['FTE_Necessario'].values[0]

                # WIP individual para a caixa
                wip_base = math.ceil(fte_atual)
                wip_protecao = 0
                sucessores = df_v[df_v['Precedência'].apply(lambda p: atv in p)]
                for _, suc in sucessores.iterrows():
                    if suc['Tempo Ciclo (min)'] > 0 and suc['FTE_Necessario'] > 0:
                        cadencia_suc = suc['Tempo Ciclo (min)'] / suc['FTE_Necessario']
                        wip_protecao = max(wip_protecao, math.ceil(tc_atual / cadencia_suc))
                wip_final = max(wip_base, wip_protecao)

                # Caixa Azul
                fig_net.add_trace(go.Scatter(
                    x=[x], y=[y], mode="markers+text",
                    marker=dict(size=marker_size, symbol="square", color=AZUL_KPMG, line=dict(width=1, color=PRETO)),
                    text=f"<b>{atv}</b>",
                    textfont=dict(color="white", size=12),
                    textposition="middle center", hoverinfo="skip"
                ))

                # Texto fora da caixa
                fig_net.add_trace(go.Scatter(
                    x=[x], y=[y - (marker_size / 45)], mode="text",
                    text=f"⏱️ {tc_atual} min<br>WIP Inicial: {wip_final}",
                    textfont=dict(size=11, color="black"),
                    textposition="bottom center", hoverinfo="skip"
                ))

            fig_net.update_layout(
                showlegend=False, plot_bgcolor="white", height=600,
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
            )
            st.plotly_chart(fig_net, use_container_width=True)

        with tab_bal:
            fig_bal = make_subplots(specs=[[{"secondary_y": True}]])

            fig_bal.add_trace(go.Bar(
                x=df_v['Atividade'],
                y=(t_util * df_v['FTE_Necessario'] / df_v['Tempo Ciclo (min)']),
                name="Capacidade/Dia",
                marker_color=AZUL_KPMG, opacity=0.6
            ), secondary_y=False)

            fig_bal.add_trace(go.Scatter(
                x=df_v['Atividade'], y=df_v['FTE_Necessario'],
                name="FTE Necessário", mode='lines+markers+text',
                text=df_v['FTE_Necessario'], textposition="top center",
                texttemplate='<b>%{text}</b>',
                textfont=dict(size=14, color=VERMELHO_KPMG),
                marker=dict(size=12, color=VERMELHO_KPMG),
                line=dict(color=VERMELHO_KPMG, width=3)
            ), secondary_y=True)

            fig_bal.update_layout(
                legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
                yaxis=dict(title="Capacidade (Obras/Dia)"),
                yaxis2=dict(title="FTE Necessário", showgrid=False, overlaying='y', side='right'),
                xaxis=dict(title="Operações")
            )
            fig_bal.add_hline(y=vol_alvo, line_dash="dash", line_color=PRETO, annotation_text="Target")
            st.plotly_chart(fig_bal, use_container_width=True)