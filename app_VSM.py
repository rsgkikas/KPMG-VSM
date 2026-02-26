import streamlit as st
import pandas as pd
import math
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta


# ==========================================
# 1. FUNÇÕES AUXILIARES E CONFIGURAÇÃO
# ==========================================

def arredondar_quarteirao(n):
    if n <= 0: return 0
    return math.ceil(n * 4) / 4


def extrair_dados_inteligente(df):
    colunas = df.columns.tolist()
    chaves_nome = ['atividade', 'operação', 'operacao', 'nome', 'task', 'descrição', 'processo']
    chaves_tempo = ['tempo', 'ciclo', 'minutos', 'segundos', 'duration', 'cycle', 't.c', 'tc']
    chaves_precedencia = ['precedência', 'precedencia', 'antecessora', 'dependência', 'dependencia', 'precedent']

    col_nome, col_tempo, col_precedencia = None, None, None
    for c in colunas:
        c_low = str(c).lower()
        if not col_nome and any(k in c_low for k in chaves_nome): col_nome = c
        if not col_tempo and any(k in c_low for k in chaves_tempo): col_tempo = c
        if not col_precedencia and any(k in c_low for k in chaves_precedencia): col_precedencia = c
    if not col_nome: col_nome = colunas[0]
    if not col_tempo: col_tempo = colunas[1] if len(colunas) > 1 else colunas[0]
    return col_nome, col_tempo, col_precedencia


def calcular_cronograma(df):
    tempos_ciclo = dict(zip(df['Atividade'], df['Tempo Ciclo (min)']))
    precedencias = dict(zip(df['Atividade'], df['Precedência']))
    tempos_fim = {}
    atividades_processadas = []
    lista_para_processar = df['Atividade'].tolist()
    max_iter = len(lista_para_processar) * 2
    iterações = 0
    while len(atividades_processadas) < len(lista_para_processar) and iterações < max_iter:
        progresso = False
        for atividade in lista_para_processar:
            if atividade not in atividades_processadas:
                deps = precedencias.get(atividade, [])
                deps_validas = [d for d in deps if d in tempos_ciclo]
                if not deps_validas or all(d in tempos_fim for d in deps_validas):
                    inicio = max([tempos_fim[d] for d in deps_validas], default=0)
                    tempos_fim[atividade] = inicio + tempos_ciclo[atividade]
                    atividades_processadas.append(atividade)
                    progresso = True
        iterações += 1
        if not progresso: break
    return tempos_fim


# Cores KPMG
AZUL_KPMG = '#00338D'
VERMELHO_KPMG = '#E30513'
PRETO = '#000000'
BRANCO = '#FFFFFF'

st.set_page_config(page_title="KPMG VSM Tool", layout="wide")
st.markdown('<h1 style="margin-bottom: 40px;">📂 KPMG VSM Tool</h1>', unsafe_allow_html=True)

# ==========================================
# 2. SIDEBAR E ESTADO
# ==========================================
st.sidebar.header("⚙️ Parâmetros de Produção")
vol_alvo = st.sidebar.number_input("Volume Alvo (Unidades/Dia)", min_value=1, value=100)
tempo_disp = st.sidebar.number_input("Tempo Disponível (min/Dia)", min_value=1, value=480)

if 'df_tarefas' not in st.session_state:
    st.session_state.df_tarefas = pd.DataFrame([{"Atividade": "", "Tempo Ciclo (min)": 0.0, "Precedência": []}])

# ==========================================
# 3. ENTRADA DE DADOS
# ==========================================
metodo = st.radio("Como pretende carregar os dados?", ["Manual", "Excel"], horizontal=True)

if metodo == "Excel":
    upload_file = st.file_uploader("Carregue o Excel", type=["xlsx"])
    if upload_file:
        try:
            df_bruto = pd.read_excel(upload_file)
            col_id_nome, col_id_tempo, col_id_prec = extrair_dados_inteligente(df_bruto)
            if col_id_prec:
                prec_data = df_bruto[col_id_prec].apply(
                    lambda x: [i.strip() for i in str(x).split(',')] if pd.notnull(x) and str(x).strip() != "" and str(
                        x).lower() != 'nan' else [])
            else:
                prec_data = [[]] * len(df_bruto)
            st.session_state.df_tarefas = pd.DataFrame({
                "Atividade": df_bruto[col_id_nome].astype(str),
                "Tempo Ciclo (min)": pd.to_numeric(df_bruto[col_id_tempo], errors='coerce').fillna(0.0),
                "Precedência": prec_data
            })
            st.success("🤖 IA: Colunas mapeadas!")
        except Exception as e:
            st.error(f"Erro: {e}")

st.markdown("---")

# ==========================================
# 4. TABELA E BOTÕES (DEFINIÇÃO DE 'CALCULAR')
# ==========================================
st.subheader("📝 Edição e Precedências")
opcoes_validas = [a for a in st.session_state.df_tarefas["Atividade"].tolist() if a and str(a).strip()]

df_editado = st.data_editor(
    st.session_state.df_tarefas,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Atividade": st.column_config.TextColumn("Operação", required=True),
        "Tempo Ciclo (min)": st.column_config.NumberColumn("T.C. (min)", format="%.3f"),
        "Precedência": st.column_config.MultiselectColumn("Precedência", options=opcoes_validas)
    },
    key="vsm_editor_final"
)

# Sincronização imediata
st.session_state.df_tarefas = df_editado

col_btn1, col_btn2 = st.columns([1, 5])
with col_btn1:
    # AQUI É DEFINIDA A VARIÁVEL 'calcular'
    calcular = st.button("🚀 Calcular")
with col_btn2:
    if st.button("🗑️ Reset"):
        st.session_state.df_tarefas = pd.DataFrame([{"Atividade": "", "Tempo Ciclo (min)": 0.0, "Precedência": []}])
        st.rerun()

# ==========================================
# 5. RESULTADOS (USO DE 'CALCULAR')
# ==========================================
if calcular:
    df_v = st.session_state.df_tarefas[st.session_state.df_tarefas["Atividade"].astype(str).str.strip() != ""].copy()

    if not df_v.empty:
        tempos_fim_calculados = calcular_cronograma(df_v)

        # Cálculos de FTE e Output
        df_v['FTE_Teorico'] = (df_v['Tempo Ciclo (min)'] * vol_alvo) / tempo_disp
        df_v['FTE_Real'] = df_v['FTE_Teorico'].apply(arredondar_quarteirao)
        df_v['Output_Real'] = (tempo_disp * df_v['FTE_Real']) / df_v['Tempo Ciclo (min)']

        # Métricas para os Cards
        lt_total = max(tempos_fim_calculados.values()) if tempos_fim_calculados else 0
        fte_total = df_v['FTE_Real'].sum()

        # --- CARDS PEQUENOS LADO A LADO ---
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"""<div style="background-color:{AZUL_KPMG}; padding:20px; border-radius:10px; text-align:center;">
                <p style="color:white; margin:0; font-size:16px;">⏱️ Lead Time Total</p>
                <h2 style="color:white; margin:0; font-size:32px;">{lt_total:.2f} min</h2>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div style="background-color:{AZUL_KPMG}; padding:20px; border-radius:10px; text-align:center;">
                <p style="color:white; margin:0; font-size:16px;">👥 Total FTE Necessário</p>
                <h2 style="color:white; margin:0; font-size:32px;">{fte_total:.2f}</h2>
            </div>""", unsafe_allow_html=True)

        # --- GANTT ---
        st.write("")
        st.subheader("📅 Gantt Ilustrativo")
        gantt_data = []
        base = datetime(2026, 1, 1, 8, 0)
        for _, row in df_v.iterrows():
            f = tempos_fim_calculados.get(row['Atividade'], 0)
            i = f - row['Tempo Ciclo (min)']
            gantt_data.append(
                dict(Task=row['Atividade'], Start=base + timedelta(minutes=i), Finish=base + timedelta(minutes=f)))

        fig_gantt = px.timeline(pd.DataFrame(gantt_data), x_start="Start", x_end="Finish", y="Task",
                                color_discrete_sequence=[AZUL_KPMG])
        fig_gantt.update_yaxes(autorange="reversed")
        st.plotly_chart(fig_gantt, use_container_width=True)

        # --- BALANCEAMENTO ---
        st.subheader("📊 Balanceamento")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(x=df_v['Atividade'], y=df_v['Output_Real'], name="Output", marker_color=AZUL_KPMG, opacity=0.7),
            secondary_y=False)
        fig.add_trace(go.Scatter(x=df_v['Atividade'], y=df_v['FTE_Real'], name="FTE", mode='lines+markers+text',
                                 marker=dict(size=12, color=PRETO), text=df_v['FTE_Real'].apply(lambda x: f'{x:.2f}')),
                      secondary_y=True)
        fig.add_hline(y=vol_alvo, line_dash="dash", line_color=VERMELHO_KPMG)
        st.plotly_chart(fig, use_container_width=True)