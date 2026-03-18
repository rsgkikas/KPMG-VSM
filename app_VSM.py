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
    return math.ceil(n * 2) / 2


def extrair_dados_inteligente(df):
    colunas = df.columns.tolist()

    # Palavras-chave por coluna — basta que UMA delas apareça em qualquer parte do nome
    chaves_nome = [
        # PT
        'atividade', 'actividade', 'operação', 'operacao', 'nome', 'etapa', 'passo', 'tarefa', 'processo', 'step',
        # EN
        'activity', 'operation', 'task', 'process', 'name', 'stage', 'step',
    ]
    chaves_tempo = [
        # PT
        'tempo', 'ciclo', 'duracao', 'duração', 'minuto', 'segundo', 'tc', 'takt',
        # EN
        'time', 'cycle', 'duration', 'minute', 'second', 'processing',
    ]
    chaves_precedencia = [
        # PT
        'precedência', 'precedencia', 'antecessora', 'antecessor', 'anterior', 'dependência', 'dependencia', 'prereq',
        # EN
        'predecessor', 'precedence', 'dependency', 'depends', 'requires', 'after', 'previous',
    ]
    chaves_fte_atual = [
        # PT
        'fte atual', 'fte_atual', 'fte actual', 'fte_actual', 'atual', 'actual', 'existente',
        'colaboradores atual', 'headcount atual', 'hc atual',
        # EN
        'current fte', 'current_fte', 'fte current', 'existing fte', 'fte existing',
        'headcount current', 'current headcount', 'hc current',
        # Só "fte" (caso a coluna se chame apenas "FTE")
        'fte',
    ]

    col_n, col_t, col_p, col_fte = None, None, None, None

    for c in colunas:
        c_low = str(c).lower().strip()
        if not col_n and any(k in c_low for k in chaves_nome):
            col_n = c
        if not col_t and any(k in c_low for k in chaves_tempo):
            col_t = c
        if not col_p and any(k in c_low for k in chaves_precedencia):
            col_p = c
        if not col_fte and any(k in c_low for k in chaves_fte_atual):
            col_fte = c

    return col_n or colunas[0], col_t or colunas[1], col_p, col_fte


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

    for _, row in df.iterrows():
        op = row["Atividade"]
        fte = max(1, math.ceil(row["FTE_Necessario"]))
        recursos[op] = []
        for i in range(fte):
            recursos[op].append({
                "id": f"{op}_Posto_{i + 1}",
                "livre_em": start_dt
            })

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

    obra_id = 1
    fila_operacao = {op: [] for op in df["Atividade"]}
    for op in fila_operacao:
        for _ in range(wip_por_operacao[op]):
            fila_operacao[op].append(f"Obra {obra_id}")
            obra_id += 1

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
            fim = ajustar_com_pausas(inicio, row["Tempo Ciclo (min)"], t_ini, t_fim, pausas)
            recurso["livre_em"] = fim
            cronograma.append({
                "Obra": obra,
                "Operacao": op,
                "Recurso": recurso["id"],
                "Start": inicio,
                "Finish": fim
            })
            houve_trabalho = True
            sucessores = df[df['Precedência'].apply(
                lambda p: op in p if isinstance(p, list) else False
            )]
            for _, suc in sucessores.iterrows():
                fila_operacao[suc["Atividade"]].append(obra)

        if not houve_trabalho:
            break

    return pd.DataFrame(cronograma)


# ==========================================
# NOVO: MOTOR DE MAXIMIZAÇÃO DE OUTPUT
# ==========================================
def calcular_fte_por_operacao_maximizar(df_v, fte_total_disponivel, t_util):
    """
    Distribui FTE pelas operações para maximizar output.

    Estratégia:
    1. A operação gargalo (maior TC) determina o output máximo.
    2. Distribuímos FTE proporcionalmente ao TC de cada operação,
       garantindo pelo menos 0.5 FTE por operação.
    3. Iteramos arredondamentos para garantir que o total bate certo.
    4. O takt implícito é derivado da operação mais carregada (TC/FTE).
    """
    df_res = df_v.copy()
    n_ops = len(df_res)

    if n_ops == 0 or fte_total_disponivel <= 0:
        df_res['FTE_Necessario'] = 0.0
        return df_res, 0, 0

    tc_total = df_res['Tempo Ciclo (min)'].sum()

    if tc_total == 0:
        df_res['FTE_Necessario'] = 0.0
        return df_res, 0, 0

    # Distribuição proporcional contínua
    fte_continuo = df_res['Tempo Ciclo (min)'] / tc_total * fte_total_disponivel

    # Arredondar para múltiplos de 0.5, mínimo 0.5 por operação
    fte_arredondado = fte_continuo.apply(lambda x: max(0.5, math.floor(x * 2) / 2))

    # Ajustar diferença para bater com o total disponível
    diferenca = fte_total_disponivel - fte_arredondado.sum()

    # Adicionar 0.5 às operações com maior "resto" até fechar o total
    restos = (fte_continuo - fte_arredondado).sort_values(ascending=False)
    passos = round(diferenca / 0.5)
    for i in range(max(0, int(passos))):
        idx = restos.index[i % len(restos)]
        fte_arredondado[idx] += 0.5

    df_res['FTE_Necessario'] = fte_arredondado

    # Takt implícito: o pior rácio TC/FTE (operação mais carregada define o ritmo)
    df_res['_takt_local'] = df_res['Tempo Ciclo (min)'] / df_res['FTE_Necessario']
    takt_implicito = df_res['_takt_local'].max()
    df_res.drop(columns=['_takt_local'], inplace=True)

    # Output máximo possível com este takt
    output_maximo = t_util / takt_implicito if takt_implicito > 0 else 0

    return df_res, takt_implicito, output_maximo


# ==========================================
# 2. UI
# ==========================================
AZUL_KPMG = '#00338D'
VERMELHO_KPMG = '#E30513'
LARANJA_BOTTLENECK = '#FF8C00'  # laranja visível mas não conflitua com o texto vermelho do FTE
PRETO = '#000000'
CYAN_WIP = '#75D1E0'

st.set_page_config(page_title="KPMG VSM Tool", layout="wide")
st.markdown(f'<h1 style="color:{AZUL_KPMG};">KPMG VSM Tool</h1>', unsafe_allow_html=True)

# ==========================================
# SIDEBAR
# ==========================================

if 'turnos' not in st.session_state:
    st.session_state.turnos = []
if 'pausas' not in st.session_state:
    st.session_state.pausas = []

# ==========================================
# SELETOR DE MODO DE OTIMIZAÇÃO
# ==========================================
st.sidebar.subheader("🎯 Objetivo de Otimização")

modo_otimizacao = st.sidebar.radio(
    "Modo:",
    ["🎯 Output Target", "💪 Maximizar Output (FTE Fixo)"],
    key="modo_otimizacao",
    help=(
        "**Output Target**: define o nº de obras/dia pretendido e a ferramenta calcula o FTE necessário.\n\n"
        "**Maximizar Output (FTE Fixo)**: define o total de FTE disponível e a ferramenta maximiza o output possível."
    )
)

if modo_otimizacao == "🎯 Output Target":
    vol_alvo = st.sidebar.number_input(
        "Objetivo Output / Dia",
        min_value=1,
        value=50,
        key="obj_vol",
        help="Número de obras/unidades a produzir por dia."
    )
    fte_disponivel = None
else:
    fte_disponivel = st.sidebar.number_input(
        "FTE Total Disponível",
        min_value=0.5,
        value=10.0,
        step=0.5,
        key="fte_total",
        help="Total de colaboradores (FTE) disponíveis para distribuir pelas operações."
    )
    vol_alvo = None

# ==========================================
# TEMPO DISPONÍVEL — MODO SIMPLIFICADO vs DETALHADO
# ==========================================
st.sidebar.markdown("---")
st.sidebar.subheader("⏱️ Tempo Disponível / Dia")

modo_tempo = st.sidebar.radio(
    "Como preferes introduzir?",
    ["⚡ Minutos", "🕒 Turnos + Pausas"],
    key="modo_tempo",
    help=(
        "**Minutos**: introduz directamente os minutos disponíveis para produção.\n\n"
        "**Turnos + Pausas**: define entrada/saída e pausas — o tempo útil é calculado automaticamente."
    )
)

if modo_tempo == "⚡ Minutos":
    t_util_direto = st.sidebar.number_input(
        "Minutos disponíveis / dia",
        min_value=1,
        value=480,
        step=5,
        key="t_util_direto",
        help="Tempo líquido de produção por dia, já descontadas pausas e tempos não produtivos."
    )
else:
    # --- TURNOS ---
    st.sidebar.markdown("**🕒 Turnos**")
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

    # --- PAUSAS ---
    st.sidebar.markdown("**☕ Pausas**")
    c_p1, c_p2 = st.sidebar.columns(2)
    p_ini_i = c_p1.time_input("Início Pausa", time(12, 0), key="new_p_ini")
    p_fim_i = c_p2.time_input("Fim Pausa", time(13, 0), key="new_p_fim")

    if st.sidebar.button("➕ Adicionar Pausa", use_container_width=True):
        st.session_state.pausas.append((p_ini_i, p_fim_i))
        st.rerun()

    if st.session_state.pausas:
        for i, (p_s, p_e) in enumerate(st.session_state.pausas):
            st.sidebar.caption(f"{i + 1}. {p_s.strftime('%H:%M')} - {p_e.strftime('%H:%M')}")
        if st.sidebar.button("🗑️ Limpar Pausas", key="clear_pausas"):
            st.session_state.pausas = []
            st.rerun()

# ==========================================
# IMPORTAR OPERAÇÕES — APENAS VIA EXCEL
# ==========================================

if 'df_tarefas' not in st.session_state:
    st.session_state.df_tarefas = pd.DataFrame({
        "Atividade": pd.Series(dtype=str),
        "Tempo Ciclo (min)": pd.Series(dtype=float),
        "Precedência": pd.Series(dtype=object),
        "FTE Atual": pd.Series(dtype=float),
    })

st.markdown("### 📋 Operações do Processo")

# --- Template para download ---
import io

df_template = pd.DataFrame({
    "Atividade": ["Operação A", "Operação B", "Operação C", "Operação D"],
    "Tempo Ciclo (min)": [10, 15, 8, 12],
    "Precedência": ["", "Operação A", "Operação A", "Operação B, Operação C"],
    "FTE Atual": [2, 3, 2, 2],
})
buffer_template = io.BytesIO()
with pd.ExcelWriter(buffer_template, engine="openpyxl") as writer:
    df_template.to_excel(writer, index=False, sheet_name="Processo")
    instrucoes = pd.DataFrame({
        "Campo": ["Atividade", "Tempo Ciclo (min)", "Precedência", "FTE Atual"],
        "Descrição": [
            "Nome único da operação",
            "Tempo de ciclo em minutos (ex: 12.5)",
            "Nome da(s) operação(ões) antecessora(s). Se houver mais de uma, separar por vírgula (ex: Operação A, Operação B). Deixar em branco se for a primeira.",
            "Número de FTE actualmente alocado a esta operação (opcional — preencher para ver comparação Actual vs Óptimo)"
        ]
    })
    instrucoes.to_excel(writer, index=False, sheet_name="Instruções")
buffer_template.seek(0)

# Caixa compacta azul claro para o template — aparece antes do uploader
st.markdown(
    f"""
    <div style="background-color:#E8F0FB; border:1px solid #B0C4DE; border-radius:8px;
                padding:10px 16px; margin-bottom:12px; display:flex;
                align-items:center; gap:12px; width:fit-content;">
        <span style="font-size:15px;">📄 Não tens o ficheiro preparado?</span>
    </div>
    """,
    unsafe_allow_html=True
)
st.download_button(
    label="⬇️ Descarregar template Excel",
    data=buffer_template,
    file_name="template_vsm.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Descarrega um ficheiro Excel com a estrutura correcta e exemplos de preenchimento.",
)

st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

f = st.file_uploader(
    "Carrega o ficheiro Excel com as operações",
    type=["xlsx"],
    key="excel_upload",
    help="O ficheiro deve ter as colunas: Atividade, Tempo Ciclo (min), Precedência"
)

if f:
    df_b = pd.read_excel(f)
    col_n, col_t, col_p, col_fte_atual = extrair_dados_inteligente(df_b)

    st.session_state.df_tarefas = pd.DataFrame({
        "Atividade": df_b[col_n].astype(str),
        "Tempo Ciclo (min)": pd.to_numeric(df_b[col_t], errors='coerce').fillna(0.0),
        "Precedência": df_b[col_p].apply(
            lambda x: [i.strip() for i in str(x).split(',')] if pd.notnull(x) and str(x).strip() != "" else []
        ) if col_p else [[]] * len(df_b),
        "FTE Atual": pd.to_numeric(df_b[col_fte_atual], errors='coerce') if col_fte_atual else float('nan'),
    })
    tem_fte_atual = col_fte_atual is not None and st.session_state.df_tarefas["FTE Atual"].notna().any()
    st.success(
        f"✅ {len(st.session_state.df_tarefas)} operações importadas."
        + (" FTE Atual detectado — será mostrada comparação Atual vs Ótimo." if tem_fte_atual else "")
    )

# Mostrar tabela lida (read-only, para confirmação visual)
if not st.session_state.df_tarefas.empty:
    df_display = st.session_state.df_tarefas.copy()
    df_display["Precedência"] = df_display["Precedência"].apply(
        lambda x: ", ".join(x) if isinstance(x, list) else str(x)
    )
    st.dataframe(df_display, use_container_width=True, hide_index=True)

df_editado = st.session_state.df_tarefas.copy()

# ==========================================
# 3. ANALYTICS
# ==========================================
if 'calculo_feito' not in st.session_state:
    st.session_state.calculo_feito = False

if st.button("🚀 RUN", use_container_width=True):
    st.session_state.calculo_feito = True

if st.session_state.calculo_feito:
    df_v = df_editado[df_editado["Atividade"].str.strip() != ""].copy()

    df_v["Tempo Ciclo (min)"] = pd.to_numeric(
        df_v["Tempo Ciclo (min)"].astype(str).str.replace(",", "."),
        errors="coerce"
    ).fillna(0)
    df_v["Precedência"] = df_v["Precedência"].apply(lambda x: x if isinstance(x, list) else [])
    df_v['FTE_Necessario'] = 0.0

    # FTE Atual (opcional — vem do Excel)
    if "FTE Atual" not in df_v.columns:
        df_v["FTE Atual"] = float('nan')
    else:
        df_v["FTE Atual"] = pd.to_numeric(df_v["FTE Atual"], errors='coerce')
    tem_fte_atual = df_v["FTE Atual"].notna().any()

    if not df_v.empty:
        # TEMPO ÚTIL — depende do modo escolhido
        if modo_tempo == "⚡ Minutos":
            t_util = float(t_util_direto)
        else:
            if not st.session_state.turnos:
                st.error("⚠️ ERRO: Por favor, adicione pelo menos um turno na barra lateral.")
                st.stop()

            tempo_bruto_turnos = 0
            for t_ini, t_fim in st.session_state.turnos:
                duracao = (
                                  datetime.combine(datetime.today(), t_fim) -
                                  datetime.combine(datetime.today(), t_ini)
                          ).total_seconds() / 60
                if duracao < 0:
                    duracao += 1440
                tempo_bruto_turnos += duracao

            total_pausas = sum([
                (datetime.combine(datetime.today(), p[1]) -
                 datetime.combine(datetime.today(), p[0])).total_seconds() / 60
                for p in st.session_state.pausas
            ])

            t_util = max(0, tempo_bruto_turnos - total_pausas)

        # ==========================================
        # RAMIFICAÇÃO: MODO OUTPUT TARGET vs FTE FIXO
        # ==========================================
        if modo_otimizacao == "🎯 Output Target":
            # --- MODO 1: ORIGINAL ---
            takt = t_util / vol_alvo if vol_alvo > 0 else 0

            if takt > 0:
                df_v['FTE_Necessario'] = df_v['Tempo Ciclo (min)'].apply(
                    lambda x: arredondar_meio_fte(x / takt)
                )

            output_resultado = vol_alvo
            fte_resultado = df_v['FTE_Necessario'].sum()

            # Banner informativo do modo (gargalo calculado abaixo, em scope partilhado)
            _banner_modo1 = True  # flag para imprimir banner após cálculo do gargalo

        else:
            # --- MODO 2: MAXIMIZAR OUTPUT COM FTE FIXO ---
            df_v, takt, output_resultado = calcular_fte_por_operacao_maximizar(
                df_v, fte_disponivel, t_util
            )
            fte_resultado = df_v['FTE_Necessario'].sum()
            _banner_modo1 = False

        # ==========================================
        # ORDENAR TOPOLOGICAMENTE (garante ordem correcta em todos os gráficos)
        # ==========================================
        try:
            df_v = ordenar_topologicamente(df_v)
        except Exception:
            pass  # se houver ciclo, mantém a ordem original

        # ==========================================
        # IDENTIFICAR GARGALO (COMUM A AMBOS OS MODOS)
        # TC / FTE mais alto = operação que limita o ritmo
        # ==========================================
        df_v['_takt_local'] = df_v.apply(
            lambda r: r['Tempo Ciclo (min)'] / r['FTE_Necessario']
            if r['FTE_Necessario'] > 0 else float('inf'),
            axis=1
        )
        operacao_gargalo = df_v.loc[df_v['_takt_local'].idxmax(), 'Atividade']
        df_v.drop(columns=['_takt_local'], inplace=True)

        # (banners removidos — informação disponível nos KPI cards e captions)

        # ==========================================
        # WIP E TPT (COMUM A AMBOS OS MODOS)
        # ==========================================
        wip_total_obras = 0
        for _, row in df_v.iterrows():
            tc_atual = row['Tempo Ciclo (min)']
            fte_nec = row['FTE_Necessario']
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

        # ==========================================
        # TAXA DE OCUPAÇÃO (calculada antes dos cards)
        # ==========================================
        # Ocupação Ótima: TC / (FTE_Necessario * takt) — quanto do tempo útil cada FTE está a trabalhar
        df_v['_ocup_otimo'] = df_v.apply(
            lambda r: min(r['Tempo Ciclo (min)'] / (r['FTE_Necessario'] * takt) * 100, 100)
            if r['FTE_Necessario'] > 0 and takt > 0 else 0,
            axis=1
        )
        ocup_otimo_media = df_v['_ocup_otimo'].mean()

        if tem_fte_atual:
            fte_atual_total = df_v['FTE Atual'].fillna(0)
            takt_atual_por_op = df_v.apply(
                lambda r: r['Tempo Ciclo (min)'] / r['FTE Atual']
                if pd.notna(r['FTE Atual']) and r['FTE Atual'] > 0 else float('inf'),
                axis=1
            )
            takt_atual = takt_atual_por_op.replace(float('inf'), 0).max() if (
                        takt_atual_por_op != float('inf')).any() else 0
            df_v['_ocup_atual'] = df_v.apply(
                lambda r: min(r['Tempo Ciclo (min)'] / (r['FTE Atual'] * takt) * 100, 100)
                if pd.notna(r['FTE Atual']) and r['FTE Atual'] > 0 and takt > 0 else 0,
                axis=1
            )
            ocup_atual_media = df_v['_ocup_atual'].mean()
        else:
            ocup_atual_media = None

        # ==========================================
        # KPI CARDS (adaptados ao modo)
        # ==========================================
        st.markdown("---")

        if modo_otimizacao == "🎯 Output Target":
            label_output = f"{int(output_resultado)}"
            label_output_title = "🎯 Output Target"
        else:
            label_output = f"{output_resultado:.1f}"
            label_output_title = "📈 Output Máx./Dia"

        # Card de ocupação — valores com 0 casas decimais
        if ocup_atual_media is not None:
            ocup_label = "📊 Ocupação Média"
            ocup_valor_html = f"Atual: {ocup_atual_media:.0f}% &rarr; Ótimo: {ocup_otimo_media:.0f}%"
            ocup_font = "15px"
        else:
            ocup_label = "📊 Ocupação Ótima Média"
            ocup_valor_html = f"{ocup_otimo_media:.0f}%"
            ocup_font = "24px"

        # 6 cards numa única grid (5 principais + ocupação)
        st.markdown(f"""
            <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:25px;">
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">{label_output_title}</small><br>
                    <span style="font-size:24px;font-weight:bold;">{label_output}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">&#128100; FTE Total</small><br>
                    <span style="font-size:24px;font-weight:bold;">{total_fte_necessario:.1f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">&#9881;&#65039; WIP (un.)</small><br>
                    <span style="font-size:24px;font-weight:bold;">{wip_total_obras:.0f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">&#128640; Throughput Time (min)</small><br>
                    <span style="font-size:24px;font-weight:bold;">{tpt_total_min:.2f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">&#9203; Takt Time (min)</small><br>
                    <span style="font-size:24px;font-weight:bold;">{takt:.2f}</span>
                </div>
                <div style="background-color:{AZUL_KPMG};padding:20px;border-radius:10px;text-align:center;color:white;">
                    <small style="text-transform:uppercase;font-weight:bold;opacity:0.8;">{ocup_label}</small><br>
                    <span style="font-size:{ocup_font};font-weight:bold;">{ocup_valor_html}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # ==========================================
        # ABAS DE VISUALIZAÇÃO
        # ==========================================
        titulo_bal = "📊 Balanceamento (Atual vs Ótimo)" if tem_fte_atual else "📊 Balanceamento"
        titulo_ocup = "📈 Ocupação (Atual vs Ótimo)" if tem_fte_atual else "📈 Ocupação"
        tab_mapa, tab_bal, tab_ocup = st.tabs([
            "🌲 Mapa de Precedências & WIP", titulo_bal, titulo_ocup
        ])

        with tab_mapa:
            st.subheader("Fluxo Operacional, TC & WIP Ótimo | Foco Cumprimento Target")

            niveis = {}
            while len(niveis) < len(df_v):
                for _, row in df_v.iterrows():
                    atv = row["Atividade"]
                    preds = row["Precedência"]
                    if atv in niveis:
                        continue
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

            for _, row in df_v.iterrows():
                destino = row["Atividade"]
                for origem in row["Precedência"]:
                    if origem in pos:
                        x0, y0 = pos[origem]
                        x1, y1 = pos[destino]
                        fig_net.add_annotation(
                            x=x1 - 1.2, y=y1, ax=x0 + 1.2, ay=y0,
                            xref="x", yref="y", axref="x", ayref="y",
                            showarrow=True, arrowhead=3, arrowsize=1.2,
                            arrowwidth=2, arrowcolor=PRETO
                        )

            num_atividades = len(df_v)
            marker_size = 80 if num_atividades < 10 else 60

            for atv, (x, y) in pos.items():
                row_data = df_v[df_v['Atividade'] == atv]
                tc_atual = row_data['Tempo Ciclo (min)'].values[0]
                fte_atual = row_data['FTE_Necessario'].values[0]

                wip_base = math.ceil(fte_atual)
                wip_protecao = 0
                sucessores = df_v[df_v['Precedência'].apply(lambda p: atv in p)]
                for _, suc in sucessores.iterrows():
                    if suc['Tempo Ciclo (min)'] > 0 and suc['FTE_Necessario'] > 0:
                        cadencia_suc = suc['Tempo Ciclo (min)'] / suc['FTE_Necessario']
                        wip_protecao = max(wip_protecao, math.ceil(tc_atual / cadencia_suc))
                wip_final = max(wip_base, wip_protecao)

                # Destacar gargalo a laranja (ambos os modos)
                cor_caixa = LARANJA_BOTTLENECK if atv == operacao_gargalo else AZUL_KPMG

                fig_net.add_trace(go.Scatter(
                    x=[x], y=[y], mode="markers+text",
                    marker=dict(size=marker_size, symbol="square", color=cor_caixa,
                                line=dict(width=1, color=PRETO)),
                    text=f"<b>{atv}</b>",
                    textfont=dict(color="white", size=12),
                    textposition="middle center", hoverinfo="skip"
                ))

                fig_net.add_trace(go.Scatter(
                    x=[x], y=[y - (marker_size / 45)], mode="text",
                    text=f"⏱️ {tc_atual} min | FTE: {fte_atual}<br>WIP Inicial: {wip_final}",
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
            st.caption(f"🟠 Bottleneck: **{operacao_gargalo}**")

        with tab_bal:
            # Ordem já garantida pelo sort topológico feito antes
            ordem_ops = df_v['Atividade'].tolist()

            # Capacidade diária por operação (Ótimo)
            cap_otimo = (
                    t_util * df_v['FTE_Necessario'] /
                    df_v['Tempo Ciclo (min)'].replace(0, float('inf'))
            ).tolist()

            cores_barras = [
                LARANJA_BOTTLENECK if atv == operacao_gargalo else AZUL_KPMG
                for atv in ordem_ops
            ]

            if tem_fte_atual:
                # ---- COMPARAÇÃO ATUAL vs ÓTIMO ----
                fte_atual_vals = df_v['FTE Atual'].fillna(0).tolist()
                cap_atual = [
                    t_util * fa / tc if tc > 0 else 0
                    for fa, tc in zip(fte_atual_vals, df_v['Tempo Ciclo (min)'].tolist())
                ]

                fig_bal = make_subplots(
                    rows=1, cols=2,
                    subplot_titles=("FTE: Atual vs Ótimo", "Capacidade/Dia: Atual vs Ótimo"),
                    horizontal_spacing=0.10
                )

                # Painel esquerdo: FTE
                fig_bal.add_trace(go.Bar(
                    name="Atual", x=ordem_ops, y=fte_atual_vals,
                    marker_color='#AAAAAA', opacity=0.85,
                    legendgroup="atual", showlegend=True
                ), row=1, col=1)
                fig_bal.add_trace(go.Bar(
                    name="Ótimo", x=ordem_ops, y=df_v['FTE_Necessario'].tolist(),
                    marker_color=[LARANJA_BOTTLENECK if a == operacao_gargalo else AZUL_KPMG for a in ordem_ops],
                    opacity=0.85,
                    legendgroup="otimo", showlegend=True
                ), row=1, col=1)

                # Painel direito: Capacidade
                fig_bal.add_trace(go.Bar(
                    name="Atual", x=ordem_ops, y=cap_atual,
                    marker_color='#AAAAAA', opacity=0.85,
                    legendgroup="atual", showlegend=False
                ), row=1, col=2)
                fig_bal.add_trace(go.Bar(
                    name="Ótimo", x=ordem_ops, y=cap_otimo,
                    marker_color=[LARANJA_BOTTLENECK if a == operacao_gargalo else AZUL_KPMG for a in ordem_ops],
                    opacity=0.85,
                    legendgroup="otimo", showlegend=False
                ), row=1, col=2)

                # Linha target sem anotação — entra na legenda via scatter invisível
                fig_bal.add_hline(
                    y=output_resultado, line_dash="dash", line_color=PRETO,
                    row=1, col=2
                )
                # Entrada de legenda para o Target
                fig_bal.add_trace(go.Scatter(
                    x=[None], y=[None], mode='lines',
                    name="Target Output",
                    line=dict(color=PRETO, dash='dash', width=2),
                    legendgroup="target", showlegend=True
                ), row=1, col=2)

                fig_bal.update_layout(
                    barmode='group',
                    showlegend=True,
                    legend=dict(
                        orientation="h",
                        yanchor="top", y=-0.18,
                        xanchor="center", x=0.5,
                    ),
                    height=420,
                    yaxis=dict(title="FTE"),
                    yaxis2=dict(title="Capacidade (Obras/Dia)"),
                )

            else:
                # ---- SEM DADOS ATUAIS ----
                fig_bal = make_subplots(specs=[[{"secondary_y": True}]])

                fig_bal.add_trace(go.Bar(
                    x=ordem_ops, y=cap_otimo,
                    name="Capacidade/Dia",
                    marker_color=cores_barras,
                    opacity=0.85, showlegend=False
                ), secondary_y=False)

                fig_bal.add_trace(go.Scatter(
                    x=ordem_ops, y=df_v['FTE_Necessario'].tolist(),
                    name="FTE Ótimo", mode='lines+markers+text',
                    text=[f"{v:.1f}" for v in df_v['FTE_Necessario'].tolist()],
                    textposition="top center",
                    texttemplate='<b>%{text}</b>',
                    textfont=dict(size=14, color=VERMELHO_KPMG),
                    marker=dict(size=12, color=VERMELHO_KPMG),
                    line=dict(color=VERMELHO_KPMG, width=3),
                    showlegend=False
                ), secondary_y=True)

                # Linha target sem anotação — entra só na legenda
                fig_bal.add_hline(
                    y=output_resultado, line_dash="dash", line_color=PRETO,
                )
                fig_bal.add_trace(go.Scatter(
                    x=[None], y=[None], mode='lines',
                    name="Target Output",
                    line=dict(color=PRETO, dash='dash', width=2),
                    showlegend=True
                ), secondary_y=False)

                fig_bal.update_layout(
                    showlegend=True,
                    legend=dict(
                        orientation="h",
                        yanchor="top", y=-0.18,
                        xanchor="center", x=0.5,
                    ),
                    yaxis=dict(title="Capacidade (Obras/Dia)"),
                    yaxis2=dict(title="FTE Ótimo", showgrid=False, overlaying='y', side='right'),
                    xaxis=dict(title="Operações")
                )

            st.plotly_chart(fig_bal, use_container_width=True)
            st.caption(f"🟠 Bottleneck: **{operacao_gargalo}**")

        # ==========================================
        # ABA OCUPAÇÃO
        # ==========================================
        with tab_ocup:
            st.subheader("Taxa de Ocupação")

            ocup_otimo_vals = df_v['_ocup_otimo'].tolist()
            cores_ocup = [
                LARANJA_BOTTLENECK if atv == operacao_gargalo else AZUL_KPMG
                for atv in ordem_ops
            ]

            if tem_fte_atual:
                ocup_atual_vals = df_v['_ocup_atual'].tolist()

                fig_ocup = go.Figure()
                fig_ocup.add_trace(go.Bar(
                    name="Atual", x=ordem_ops, y=ocup_atual_vals,
                    marker_color='#AAAAAA', opacity=0.85,
                ))
                fig_ocup.add_trace(go.Bar(
                    name="Ótimo", x=ordem_ops, y=ocup_otimo_vals,
                    marker_color=cores_ocup, opacity=0.85,
                ))
                # Linha dos 100%
                fig_ocup.add_hline(
                    y=100, line_dash="dot", line_color=VERMELHO_KPMG,
                )
                fig_ocup.add_trace(go.Scatter(
                    x=[None], y=[None], mode='lines',
                    name="100% Ocupação",
                    line=dict(color=VERMELHO_KPMG, dash='dot', width=2),
                    showlegend=True
                ))
                fig_ocup.update_layout(
                    barmode='group',
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
                    yaxis=dict(title="Taxa de Ocupação (%)", range=[0, 110]),
                    xaxis=dict(title="Operações"),
                    height=400,
                )
            else:
                fig_ocup = go.Figure()
                fig_ocup.add_trace(go.Bar(
                    name="Ocupação Ótima", x=ordem_ops, y=ocup_otimo_vals,
                    marker_color=cores_ocup, opacity=0.85, showlegend=False,
                    text=[f"{v:.1f}%" for v in ocup_otimo_vals],
                    textposition="outside",
                ))
                fig_ocup.add_hline(
                    y=100, line_dash="dot", line_color=VERMELHO_KPMG,
                )
                fig_ocup.add_trace(go.Scatter(
                    x=[None], y=[None], mode='lines',
                    name="100% Ocupação",
                    line=dict(color=VERMELHO_KPMG, dash='dot', width=2),
                    showlegend=True
                ))
                fig_ocup.update_layout(
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
                    yaxis=dict(title="Taxa de Ocupação (%)", range=[0, 115]),
                    xaxis=dict(title="Operações"),
                    height=400,
                )

            st.plotly_chart(fig_ocup, use_container_width=True)
            st.caption(f"🟠 Bottleneck: **{operacao_gargalo}**")

        # Limpar colunas temporárias
        df_v.drop(columns=[c for c in ['_ocup_otimo', '_ocup_atual'] if c in df_v.columns], inplace=True)