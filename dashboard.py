"""
dashboard.py
============
Dashboard Streamlit para acompanhamento de projetos exportados do MS Project.

Execução:
    streamlit run dashboard.py
"""

import math
from datetime import date

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from processar_dados import ARQUIVO_PADRAO, carregar_dados

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard de Projetos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# CSS CUSTOMIZADO
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .metric-card {
        background: #1e293b;
        border-radius: 12px;
        padding: 18px 22px;
        border-left: 5px solid #3b82f6;
    }
    .metric-card.verde { border-left-color: #22c55e; }
    .metric-card.amarelo { border-left-color: #f59e0b; }
    .metric-card.vermelho { border-left-color: #ef4444; }
    .metric-card.azul { border-left-color: #3b82f6; }
    .metric-title { color: #94a3b8; font-size: 13px; margin-bottom: 4px; }
    .metric-value { color: #f1f5f9; font-size: 28px; font-weight: 700; }
    .metric-sub { color: #64748b; font-size: 12px; margin-top: 2px; }
    [data-testid="stSidebar"] { background-color: #0f172a; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# FUNÇÕES UTILITÁRIAS
# ---------------------------------------------------------------------------

@st.cache_data(ttl=300)
def _carregar(caminho: str) -> pd.DataFrame:
    return carregar_dados(caminho)


def card(titulo: str, valor: str, sub: str = "", cor: str = "azul") -> str:
    return f"""
    <div class="metric-card {cor}">
        <div class="metric-title">{titulo}</div>
        <div class="metric-value">{valor}</div>
        <div class="metric-sub">{sub}</div>
    </div>
    """


def cor_status(status: str) -> str:
    mapa = {
        "Concluída": "#22c55e",
        "Em andamento": "#3b82f6",
        "Atrasada": "#ef4444",
        "Não iniciada": "#94a3b8",
        "Indefinido": "#64748b",
    }
    return mapa.get(status, "#64748b")


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/48/project.png",
        width=48,
    )
    st.title("Filtros")
    st.divider()

    # Upload de arquivo (opcional – usa o padrão se não houver upload)
    arquivo_upload = st.file_uploader(
        "📂 Carregar arquivo do MS Project",
        type=["mpp", "csv", "xlsx", "xls"],
        help=(
            "Formatos aceitos:\n"
            "\u2022 .mpp  — arquivo nativo do MS Project (requer Java instalado)\n"
            "\u2022 .csv  — exportação CSV do MS Project (separador ;)\n"
            "\u2022 .xlsx / .xls  — exportação Excel do MS Project\n"
            "Deixe em branco para usar o arquivo padrão da pasta."
        ),
    )

    st.divider()

    # Carrega dados
    if arquivo_upload is not None:
        import pathlib
        import tempfile

        # Preserva a extensão original para que o processador detecte o formato
        ext = pathlib.Path(arquivo_upload.name).suffix or ".csv"
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(arquivo_upload.read())
            caminho_dados = tmp.name
    else:
        caminho_dados = str(ARQUIVO_PADRAO)

    try:
        df_completo = _carregar(caminho_dados)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        st.stop()

    # Filtro: Fase
    fases_disponiveis = sorted(df_completo["fase"].dropna().unique())
    fases_sel = st.multiselect(
        "🔷 Fase",
        options=fases_disponiveis,
        default=fases_disponiveis,
    )

    # Filtro: Status
    status_disponiveis = sorted(df_completo["status"].dropna().unique())
    status_sel = st.multiselect(
        "🔘 Status",
        options=status_disponiveis,
        default=status_disponiveis,
    )

    # Filtro: Recurso
    todos_recursos = sorted(set(
        r for lst in df_completo["recursos_lista"] for r in lst if r
    ))
    recurso_sel = st.multiselect(
        "👤 Recurso",
        options=todos_recursos,
        default=[],
        placeholder="Todos os recursos",
    )

    # Filtro: Nível hierárquico
    nivel_max = int(df_completo["nivel"].max())
    if nivel_max > 0:
        nivel_sel = st.slider(
            "📐 Nível hierárquico (máx)",
            min_value=0,
            max_value=nivel_max,
            value=nivel_max,
            step=1,
        )
    else:
        nivel_sel = 0
        st.info("Hierarquia de nível único detectada.")

    st.divider()
    st.caption(f"Dados atualizados em {date.today().strftime('%d/%m/%Y')}")


# ---------------------------------------------------------------------------
# APLICAR FILTROS
# ---------------------------------------------------------------------------
df = df_completo.copy()

if fases_sel:
    df = df[df["fase"].isin(fases_sel)]

if status_sel:
    df = df[df["status"].isin(status_sel)]

if recurso_sel:
    df = df[df["recursos_lista"].apply(
        lambda lst: any(r in recurso_sel for r in lst)
    )]

df = df[df["nivel"] <= nivel_sel]

df_tarefas = df[df["nivel"] > 0].copy()  # exclui linha-raiz do projeto


# ---------------------------------------------------------------------------
# CABEÇALHO
# ---------------------------------------------------------------------------
st.title("📊 Dashboard de Acompanhamento de Projetos")
projeto_nome = df_completo[df_completo["nivel"] == 0]["nome"].values
st.subheader(projeto_nome[0] if len(projeto_nome) else "Projeto")
st.divider()


# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
total = len(df_tarefas)
concluidas = (df_tarefas["status"] == "Concluída").sum()
em_andamento = (df_tarefas["status"] == "Em andamento").sum()
atrasadas = (df_tarefas["status"] == "Atrasada").sum()
nao_iniciadas = (df_tarefas["status"] == "Não iniciada").sum()

pct_medio = df_tarefas["pct_concluido"].dropna().mean()
pct_medio_str = f"{pct_medio:.1f}%" if not math.isnan(pct_medio) else "–"

c1, c2, c3, c4, c5 = st.columns(5)
c1.markdown(card("Total de Tarefas", str(total), "todas as tarefas filtradas", "azul"), unsafe_allow_html=True)
c2.markdown(card("Concluídas", str(concluidas), f"{concluidas/total*100:.0f}% do total" if total else "–", "verde"), unsafe_allow_html=True)
c3.markdown(card("Em Andamento", str(em_andamento), "em progresso", "azul"), unsafe_allow_html=True)
c4.markdown(card("Atrasadas", str(atrasadas), "prazo vencido sem 100%", "vermelho"), unsafe_allow_html=True)
c5.markdown(card("% Médio Geral", pct_medio_str, "média de conclusão", "amarelo"), unsafe_allow_html=True)

st.divider()


# ---------------------------------------------------------------------------
# LINHA 1: Progresso por Fase | Pizza de Status
# ---------------------------------------------------------------------------
col_esq, col_dir = st.columns([2, 1])

with col_esq:
    st.subheader("📈 Progresso por Fase")

    if not df_tarefas.empty:
        df_fase = (
            df_tarefas.groupby("fase")["pct_concluido"]
            .mean()
            .reset_index()
            .sort_values("pct_concluido", ascending=True)
        )
        df_fase.columns = ["Fase", "% Concluído"]

        fig_fase = px.bar(
            df_fase,
            x="% Concluído",
            y="Fase",
            orientation="h",
            color="% Concluído",
            color_continuous_scale=["#ef4444", "#f59e0b", "#22c55e"],
            range_color=[0, 100],
            text=df_fase["% Concluído"].apply(lambda v: f"{v:.1f}%"),
        )
        fig_fase.update_traces(textposition="outside")
        fig_fase.update_layout(
            coloraxis_showscale=False,
            xaxis=dict(range=[0, 115], title="% Média de Conclusão"),
            yaxis_title="",
            height=max(300, len(df_fase) * 42),
            margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#f1f5f9",
        )
        st.plotly_chart(fig_fase, use_container_width=True)
    else:
        st.info("Nenhuma tarefa disponível para os filtros selecionados.")

with col_dir:
    st.subheader("🔵 Status Geral")

    if not df_tarefas.empty:
        df_status_count = (
            df_tarefas["status"].value_counts().reset_index()
        )
        df_status_count.columns = ["Status", "Quantidade"]

        fig_pizza = px.pie(
            df_status_count,
            names="Status",
            values="Quantidade",
            color="Status",
            color_discrete_map={
                "Concluída": "#22c55e",
                "Em andamento": "#3b82f6",
                "Atrasada": "#ef4444",
                "Não iniciada": "#94a3b8",
                "Indefinido": "#64748b",
            },
            hole=0.45,
        )
        fig_pizza.update_traces(textinfo="percent+label", textfont_size=12)
        fig_pizza.update_layout(
            showlegend=False,
            margin=dict(l=10, r=10, t=30, b=10),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#f1f5f9",
            height=350,
        )
        st.plotly_chart(fig_pizza, use_container_width=True)

st.divider()


# ---------------------------------------------------------------------------
# GANTT
# ---------------------------------------------------------------------------
st.subheader("📅 Cronograma (Gantt)")

df_gantt = df_tarefas.dropna(subset=["inicio", "termino"]).copy()
df_gantt = df_gantt[df_gantt["nome"].str.strip() != ""].copy()

if not df_gantt.empty:
    # Limita para não sobrecarregar o gráfico
    LIMITE_GANTT = 60
    if len(df_gantt) > LIMITE_GANTT:
        st.caption(
            f"ℹ️ Exibindo as {LIMITE_GANTT} primeiras tarefas. "
            "Use os filtros para refinar a visualização."
        )
        df_gantt = df_gantt.head(LIMITE_GANTT)

    # Força datetime64[ns] — obrigatório para Plotly no pandas 2+
    df_gantt["inicio_dt"] = pd.to_datetime(df_gantt["inicio"].astype(str)).astype("datetime64[ns]")
    df_gantt["termino_dt"] = pd.to_datetime(df_gantt["termino"].astype(str)).astype("datetime64[ns]")
    df_gantt["pct_label"] = df_gantt["pct_concluido"].apply(
        lambda v: f"{v:.0f}%" if v is not None and not math.isnan(v) else "–"
    )

    fig_gantt = px.timeline(
        df_gantt,
        x_start="inicio_dt",
        x_end="termino_dt",
        y="nome",
        color="status",
        color_discrete_map={
            "Concluída": "#22c55e",
            "Em andamento": "#3b82f6",
            "Atrasada": "#ef4444",
            "Não iniciada": "#94a3b8",
            "Indefinido": "#64748b",
        },
        hover_data={
            "pct_label": True,
            "fase": True,
            "recursos": True,
            "inicio_dt": "|%d/%m/%Y",
            "termino_dt": "|%d/%m/%Y",
            "nome": True,
        },
        labels={
            "pct_label": "% Concluído",
            "inicio_dt": "Início",
            "termino_dt": "Término",
            "nome": "Tarefa",
            "fase": "Fase",
            "recursos": "Recursos",
        },
    )

    # Linha vertical "hoje" — Plotly timeline exige ms Unix no eixo x
    hoje_ms = pd.Timestamp(date.today()).value // 1_000_000
    fig_gantt.add_vline(
        x=hoje_ms,
        line_dash="dash",
        line_color="#f59e0b",
        annotation_text="Hoje",
        annotation_font_color="#f59e0b",
    )

    fig_gantt.update_yaxes(autorange="reversed", tickfont=dict(size=11))
    fig_gantt.update_layout(
        xaxis_title="",
        yaxis_title="",
        legend_title="Status",
        height=max(400, len(df_gantt) * 30),
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font_color="#f1f5f9",
    )
    st.plotly_chart(fig_gantt, use_container_width=True)
else:
    st.info("Sem datas suficientes para exibir o Gantt.")

st.divider()


# ---------------------------------------------------------------------------
# LINHA 2: Tarefas por Recurso | Tarefas Atrasadas
# ---------------------------------------------------------------------------
col_r1, col_r2 = st.columns([1, 1])

with col_r1:
    st.subheader("👤 Tarefas por Recurso")

    # Explode recursos
    df_rec = df_tarefas.explode("recursos_lista").rename(columns={"recursos_lista": "recurso"})
    df_rec = df_rec[df_rec["recurso"].str.strip() != ""]

    if not df_rec.empty:
        df_rec_count = (
            df_rec.groupby(["recurso", "status"])
            .size()
            .reset_index(name="qtd")
            .sort_values("qtd", ascending=False)
        )
        top_recursos = df_rec_count.groupby("recurso")["qtd"].sum().nlargest(15).index
        df_rec_count = df_rec_count[df_rec_count["recurso"].isin(top_recursos)]

        fig_rec = px.bar(
            df_rec_count,
            x="recurso",
            y="qtd",
            color="status",
            color_discrete_map={
                "Concluída": "#22c55e",
                "Em andamento": "#3b82f6",
                "Atrasada": "#ef4444",
                "Não iniciada": "#94a3b8",
                "Indefinido": "#64748b",
            },
            barmode="stack",
            labels={"recurso": "Recurso", "qtd": "Nº de Tarefas", "status": "Status"},
        )
        fig_rec.update_layout(
            xaxis_tickangle=-35,
            height=380,
            margin=dict(l=10, r=10, t=10, b=80),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#f1f5f9",
            legend_title="Status",
        )
        st.plotly_chart(fig_rec, use_container_width=True)
    else:
        st.info("Nenhum recurso associado às tarefas filtradas.")

with col_r2:
    st.subheader("🚨 Tarefas Atrasadas")

    df_atrasadas = df_tarefas[df_tarefas["status"] == "Atrasada"].copy()
    if not df_atrasadas.empty:
        df_atrasadas["Atraso (dias)"] = df_atrasadas["termino"].apply(
            lambda t: (date.today() - t).days if pd.notna(t) else 0
        )
        df_atrasadas_show = df_atrasadas[[
            "nome", "fase", "pct_concluido", "termino", "recursos", "Atraso (dias)"
        ]].rename(columns={
            "nome": "Tarefa",
            "fase": "Fase",
            "pct_concluido": "% Concluído",
            "termino": "Prazo",
            "recursos": "Recursos",
        }).sort_values("Atraso (dias)", ascending=False)

        st.dataframe(
            df_atrasadas_show,
            use_container_width=True,
            height=360,
            column_config={
                "% Concluído": st.column_config.ProgressColumn(
                    "% Concluído",
                    min_value=0,
                    max_value=100,
                    format="%.0f%%",
                ),
                "Prazo": st.column_config.DateColumn("Prazo", format="DD/MM/YYYY"),
                "Atraso (dias)": st.column_config.NumberColumn("Atraso (dias)", format="%d dias"),
            },
            hide_index=True,
        )
    else:
        st.success("✅ Nenhuma tarefa atrasada com os filtros atuais!")

st.divider()


# ---------------------------------------------------------------------------
# TABELA COMPLETA
# ---------------------------------------------------------------------------
st.subheader("📋 Tabela de Tarefas")

with st.expander("Ver / ocultar tabela completa", expanded=False):
    df_tabela = df_tarefas[[
        "nome", "nivel", "fase", "pct_concluido", "duracao_dias",
        "inicio", "termino", "status", "recursos", "observacao"
    ]].rename(columns={
        "nome": "Tarefa",
        "nivel": "Nível",
        "fase": "Fase",
        "pct_concluido": "% Concluído",
        "duracao_dias": "Duração (dias)",
        "inicio": "Início",
        "termino": "Término",
        "status": "Status",
        "recursos": "Recursos",
        "observacao": "Observação",
    })

    st.dataframe(
        df_tabela,
        use_container_width=True,
        height=500,
        column_config={
            "% Concluído": st.column_config.ProgressColumn(
                "% Concluído",
                min_value=0,
                max_value=100,
                format="%.0f%%",
            ),
            "Início": st.column_config.DateColumn("Início", format="DD/MM/YYYY"),
            "Término": st.column_config.DateColumn("Término", format="DD/MM/YYYY"),
            "Duração (dias)": st.column_config.NumberColumn(
                "Duração (dias)", format="%d dias"
            ),
        },
        hide_index=True,
    )

    # Botão de download
    csv_export = df_tabela.to_csv(index=False, sep=";", encoding="utf-8-sig")
    st.download_button(
        label="⬇️ Baixar tabela como CSV",
        data=csv_export,
        file_name=f"tarefas_{date.today().isoformat()}.csv",
        mime="text/csv",
    )

st.divider()

# ---------------------------------------------------------------------------
# RODAPÉ
# ---------------------------------------------------------------------------
st.caption(
    "Dashboard de Projetos · Fonte: MS Project (.mpp / .csv / .xlsx) · "
    f"Gerado em {date.today().strftime('%d/%m/%Y')}"
)
