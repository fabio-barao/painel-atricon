import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ===============================
# Configurações iniciais
# ===============================
st.set_page_config(page_title="Painel Bibliotecas - Atricon", layout="wide")

# CSS para deixar a sidebar branca
st.markdown("""
    <style>
        [data-testid="stSidebar"] {
            background-color: white;
        }
    </style>
""", unsafe_allow_html=True)

# ===============================
# Carregar base
# ===============================
@st.cache_data
def carregar_dados():
    df = pd.read_excel("Base_Painel.xlsx")
    return df

df = carregar_dados()
df["Ano"] = df["Ano"].astype(str)  # evita valores tipo 2022.5

# ===============================
# Sidebar com logo e filtros
# ===============================
with st.sidebar:
    st.image("Logo Atricon.jpg", use_container_width=True)
    st.header("Filtros")

# 1) Selecione o bloco de análise
ordem_blocos = ["Por Etapa de Ensino", "Por Rede de Ensino", "Por Região", "Por Estado"]
bloco = st.sidebar.radio(
    "1) Selecione o bloco de análise:",
    ordem_blocos,
    index=0
)

# 2) Selecione o ano
anos = st.sidebar.multiselect(
    "2) Selecione o(s) Ano(s):",
    sorted(df["Ano"].unique()),
    default=sorted(df["Ano"].unique())
)

# 3) Selecione as categorias
categorias_disponiveis = [c for c in sorted(df[df["Bloco"] == bloco]["Categoria"].unique()) if "Total" not in c]
categorias_disponiveis = ["TODOS"] + categorias_disponiveis

categorias = st.sidebar.multiselect(
    "3) Selecione as categorias de análise:",
    categorias_disponiveis,
    default=["TODOS"]
)

if "TODOS" in categorias or len(categorias) == 0:
    categorias = [c for c in categorias_disponiveis if c != "TODOS"]
    if len(categorias) == 0:
        st.sidebar.warning("⚠ Pelo menos uma categoria deve ser selecionada. Selecionando todas por padrão.")

# 4) Tipo de visualização
tipo_grafico = st.sidebar.radio(
    "4) Selecione o tipo de visualização:",
    (
        "Escolas Públicas Ativas por Existência ou Não de Biblioteca e/ou Sala de Leitura",
        "Escolas Públicas Ativas com Biblioteca e/ou Sala de Leitura, por Existência ou Não de Bibliotecário/Monitor"
    )
)

# ===============================
# Aplicar filtros
# ===============================
df_filtrado = df[
    (df["Ano"].isin(anos)) &
    (df["Bloco"] == bloco) &
    (df["Categoria"].isin(categorias))
]

# ===============================
# Preparar dados
# ===============================
def preparar_dados(tipo):
    if tipo.startswith("Escolas Públicas Ativas por Existência"):
        df1 = df_filtrado.melt(
            id_vars=["Ano", "Categoria"],
            value_vars=["Escolas com Biblioteca", "Escolas sem Biblioteca"],
            var_name="Tipo", value_name="Valor"
        )
        titulo1 = "Escolas Públicas Ativas no Ano, por Existência ou Não de Biblioteca e/ou Sala de Leitura"

        df2 = df_filtrado.melt(
            id_vars=["Ano", "Categoria"],
            value_vars=["Matrículas com Biblioteca", "Matrículas sem Biblioteca"],
            var_name="Tipo", value_name="Valor"
        )
        titulo2 = "Número de Alunos Matriculados em Escolas Públicas Ativas, Equipadas ou Não com Biblioteca e/ou Sala de Leitura"

        legenda = {
            "Escolas com Biblioteca": "Escolas com Biblioteca e/ou Sala de Leitura",
            "Escolas sem Biblioteca": "Escolas sem Biblioteca e/ou Sala de Leitura",
            "Matrículas com Biblioteca": "Matrículas em Escolas com Biblioteca e/ou Sala de Leitura",
            "Matrículas sem Biblioteca": "Matrículas em Escolas sem Biblioteca e/ou Sala de Leitura"
        }

    else:
        df1 = df_filtrado.melt(
            id_vars=["Ano", "Categoria"],
            value_vars=["Escolas com Bibliotecário", "Escolas sem Bibliotecário"],
            var_name="Tipo", value_name="Valor"
        )
        titulo1 = "Escolas Públicas Ativas com Biblioteca e/ou Sala de Leitura, por Existência ou Não de Bibliotecário/Monitor"

        df2 = df_filtrado.melt(
            id_vars=["Ano", "Categoria"],
            value_vars=["Matrículas com Bibliotecário", "Matrículas sem Bibliotecário"],
            var_name="Tipo", value_name="Valor"
        )
        titulo2 = "Número de Alunos Matriculados em Escolas Públicas Ativas com Biblioteca e/ou Sala de Leitura, por Existência ou Não de Bibliotecário/Monitor"

        legenda = {
            "Escolas com Bibliotecário": "Escolas com Bibliotecário/Monitor",
            "Escolas sem Bibliotecário": "Escolas sem Bibliotecário/Monitor",
            "Matrículas com Bibliotecário": "Matrículas em Escolas com Bibliotecário/Monitor",
            "Matrículas sem Bibliotecário": "Matrículas em Escolas sem Bibliotecário/Monitor"
        }

    # Ajusta legenda no DataFrame
    df1["Tipo"] = df1["Tipo"].map(legenda)
    df2["Tipo"] = df2["Tipo"].map(legenda)

    return df1, titulo1, df2, titulo2

# ===============================
# Gerar gráfico
# ===============================
def gerar_grafico(df_long, titulo):
    fig = px.bar(
        df_long,
        x="Categoria",
        y="Valor",
        color="Tipo",
        barmode="stack",
        facet_col="Ano",
        text=df_long["Valor"].apply(lambda x: f"{x:,.0f}"),
        title=titulo
    )

    fig.update_traces(
        textposition="inside",
        insidetextanchor="middle",
        hovertemplate="<b>%{customdata[0]}</b><br>Ano: %{customdata[1]}<br>Categoria: %{customdata[2]}<br>Valor: %{y:,.0f}",
        customdata=df_long[["Tipo", "Ano", "Categoria"]]
    )

    fig.update_layout(
        height=500,
        legend_title_text="",
        yaxis_tickformat=","
    )

    fig.for_each_xaxis(lambda axis: axis.update(tickangle=0, title_text=""))

    # Deixar só o ano nas anotações, mais visível e com cor personalizada
    fig.for_each_annotation(
        lambda a: a.update(
            text=f"<b>{a.text.split('=')[-1]}</b>",
            font=dict(size=18, color="#0071BC")
        )
    )

    return fig

# ===============================
# Exibir gráficos
# ===============================
df1, titulo1, df2, titulo2 = preparar_dados(tipo_grafico)

fig1 = gerar_grafico(df1, titulo1)
fig2 = gerar_grafico(df2, titulo2)

st.plotly_chart(fig1, use_container_width=True)
st.plotly_chart(fig2, use_container_width=True)

# ===============================
# Download Excel (versão blindada)
# ===============================
def to_excel(df):
    # Garantir que nomes de colunas sejam string
    df.columns = [str(col) for col in df.columns]

    # Criar buffer
    output = BytesIO()

    # Usar xlsxwriter para evitar erro de serialização
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados Filtrados")

        # Ajustar largura automática das colunas
        workbook = writer.book
        worksheet = writer.sheets["Dados Filtrados"]
        for i, col in enumerate(df.columns):
            col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

    output.seek(0)
    return output

st.sidebar.download_button(
    label="📥 Baixar dados filtrados em Excel",
    data=to_excel(df_filtrado),
    file_name="dados_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
