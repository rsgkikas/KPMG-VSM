import streamlit as st
import pandas as pd
import math
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta


# ==========================================
# 1. FUNÇÕES AUXILIARES E CORES
# ==========================================

def arredondar_quarteirao(n):
    if n <= 0: return 0
    return math.ceil(n * 4) / 4


# Cores Corporativas
AZUL_KPMG = '#00338D'
VERMELHO_KPMG = '#E30513'
PRETO = '#000000'
BRANCO = '#FFFFFF'

st.set_page_config(page_title="KPMG VSM Tool", layout="wide")

# Espaçamento do Título
st.markdown('<h1 style="margin-bottom: 50px;">📂 KPMG VSM Tool</h1>', unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# 2. SIDEBAR E INPUTS
# ==========================================

st.sidebar.header("⚙️ Parâmetros de Produção")
vol_alvo = st.sidebar.number_input("Volume Alvo (Unidades/Dia)", min_value=1, value=100)
tempo_disp = st.sidebar.number_input("Tempo Disponível (min/Dia)", min_value=1, value=480)

if 'df_tarefas' not in st.session_state:
    st.session_state.df_tarefas = pd.DataFrame(
        [{"Atividade": "", "Tempo Ciclo (min)": 0.0, "Precedência": []}]
    )

st.subheader("📝 Configuração de Sequenciamento e Tempos")
opcoes_validas = [a for a in st.session_state.df_tarefas["Atividade"].tolist() if a and a.strip()]

df_editado = st.data_editor(
    st.session_state.df_tarefas,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Atividade": st.column_config.TextColumn("Nome da Operação", required=True),
        "Tempo Ciclo (min)": st.column_config.NumberColumn("T.C. (min)", min_value=0.001, format="%.3f"),
        "Precedência": st.column_config.MultiselectColumn("Precedência", options=opcoes_validas)
    },
    key="vsm_editor_v5"
)

col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 4])
with col_btn1:
    if st.button("💾 Atualizar Lista"):
        st.session_state.df_tarefas = df_editado
        st.rerun()
with col_btn2:
    calcular = st.button("🚀 Calcular")
with col_btn3:
    if st.button("🗑️ Limpar Tabela"):
        st.session_state.df_tarefas = pd.DataFrame([{"Atividade": "", "Tempo Ciclo (min)": 0.0, "Precedência": []}])
        st.rerun()

# ==========================================
# 3. GANTT E PROCESSAMENTO
# ==========================================

if calcular:
    df_v = df_editado[df_editado["Atividade"].str.strip() != ""].copy()

    if not df_v.empty:
        # --- GANTT ILUSTRATIVO ---
        st.subheader("📅 Gantt Ilustrativo")

        gantt_data = []
        tempos_fim = {}

        # Ordenação simples para o Gantt respeitar a lógica de fluxo
        for _, row in df_v.iterrows():
            nome = row['Atividade']
            duracao = row['Tempo Ciclo (min)']
            precedencias = row['Precedência']

            inicio = 0
            if precedencias:
                inicio = max([tempos_fim.get(p, 0) for p in precedencias], default=0)

            fim = inicio + duracao
            tempos_fim[nome] = fim

            base = datetime(2026, 1, 1, 8, 0)
            gantt_data.append(dict(
                Task=nome,
                Start=base + timedelta(minutes=inicio),
                Finish=base + timedelta(minutes=fim)
            ))

        df_gantt = pd.DataFrame(gantt_data)
        fig_gantt = px.timeline(df_gantt, x_start="Start", x_end="Finish", y="Task",
                                color_discrete_sequence=[AZUL_KPMG])
        fig_gantt.update_yaxes(autorange="reversed")
        fig_gantt.update_layout(
            title="Sequenciamento Lógico de Operações",
            template="plotly_white",
            xaxis_title="Tempo Decorrido (HH:MM)"
        )
        st.plotly_chart(fig_gantt, use_container_width=True)

        # --- CÁLCULOS TÉCNICOS ---
        df_v['FTE_Teorico'] = (df_v['Tempo Ciclo (min)'] * vol_alvo) / tempo_disp
        df_v['FTE_Real'] = df_v['FTE_Teorico'].apply(arredondar_quarteirao)
        df_v['Output_Real'] = (tempo_disp * df_v['FTE_Real']) / df_v['Tempo Ciclo (min)']

        st.markdown("---")

        # --- GRÁFICO DE BALANCEAMENTO ---
        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # Barras de Output (Eixo Esquerdo)
        fig.add_trace(
            go.Bar(
                x=df_v['Atividade'], y=df_v['Output_Real'], name="Output Libertado",
                marker_color=AZUL_KPMG, opacity=0.7,
                text=df_v['Output_Real'].map('{:.0f} un'.format), textposition='inside'
            ),
            secondary_y=False
        )

        # Linha FTE (Eixo Direito - Preto Negrito)
        fig.add_trace(
            go.Scatter(
                x=df_v['Atividade'], y=df_v['FTE_Real'], name="Carga FTE",
                mode='lines+markers+text',
                marker=dict(size=14, symbol='circle', color=PRETO, line=dict(width=2, color=BRANCO)),
                text=df_v['FTE_Real'].apply(lambda x: f'<b>FTE: {x:.2f}</b>'),
                textposition='top center',
                textfont=dict(color=PRETO, size=13),
                line=dict(color=PRETO, width=3)
            ),
            secondary_y=True
        )

        # Linha de Meta (Eixo Esquerdo)
        fig.add_hline(y=vol_alvo, line_dash="dash", line_color=VERMELHO_KPMG,
                      annotation_text=f"Meta: {vol_alvo} un", secondary_y=False)

        # Layout e Escala Dinâmica
        max_fte = df_v['FTE_Real'].max()
        step_fte = 0.25 if max_fte <= 5 else None

        fig.update_layout(
            title_text="Análise de Produtividade vs Mão de Obra",
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.1, xanchor="center", x=0.5),
            margin=dict(t=100)
        )

        fig.update_yaxes(title_text="<b>Output</b> (Unidades/Dia)", secondary_y=False)
        fig.update_yaxes(
            title_text="<b>Carga FTE</b> (Escala 0.25)",
            secondary_y=True, dtick=step_fte, showgrid=False,
            range=[0, max_fte * 1.4]
        )

        st.plotly_chart(fig, use_container_width=True)

        # Tabela de Dados Final
        st.subheader("📋 Resumo Técnico")
        st.dataframe(df_v[['Atividade', 'Tempo Ciclo (min)', 'FTE_Teorico', 'FTE_Real', 'Output_Real']],
                     use_container_width=True)

    else:
        st.warning("Preencha a tabela e clique em 'Atualizar Lista' primeiro.")