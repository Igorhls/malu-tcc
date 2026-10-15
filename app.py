import io
import os
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from gerar_graficos import (
    DADOS,
    PONTOS,
    COR_MANHA,
    COR_TARDE,
    COR_NOITE,
    COR_TEXTO,
    COR_GRADE,
    COR_BORDA,
    configurar_estilo,
    criar_dataframe,
    gerar_planilha_excel
)

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Monitoramento de Temperatura - BOSQUE",
    page_icon="🌡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Aplicar estilo às fontes
configurar_estilo()


def render_grafico(df, info, periodos, com_marcadores=False):
    """Renderiza a figura do matplotlib de acordo com os filtros selecionados."""
    fig, ax = plt.subplots(figsize=(8.0, 4.0), dpi=200)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    marcadores = {'Manha': 'o', 'Tarde': 's', 'Noite': '^'}
    cores = {'Manha': COR_MANHA, 'Tarde': COR_TARDE, 'Noite': COR_NOITE}

    for periodo in periodos:
        m = marcadores[periodo] if com_marcadores else None
        ax.plot(
            df['Pontos'], df[periodo],
            label=periodo,
            color=cores[periodo],
            linewidth=2.2,
            marker=m,
            markersize=6
        )

    ax.set_title(info["titulo_grafico"], fontsize=13, color=COR_TEXTO, pad=25)

    if periodos:
        leg = ax.legend(
            loc='upper center', bbox_to_anchor=(0.5, 1.15),
            ncol=len(periodos), frameon=False, fontsize=10,
            handlelength=1.5, handletextpad=0.5, columnspacing=1.6
        )
        for text in leg.get_texts():
            text.set_color(COR_TEXTO)

    ax.set_ylabel("Temperatura do ar (°C)", fontsize=10, color=COR_TEXTO, labelpad=8)
    ax.set_xlabel("Pontos", fontsize=10, color=COR_TEXTO, labelpad=8)

    ax.set_ylim(info["ylim"])
    ax.yaxis.set_major_locator(ticker.MultipleLocator(2))

    ax.grid(axis='y', color=COR_GRADE, linestyle='-', linewidth=0.8)
    ax.grid(axis='x', visible=False)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color(COR_BORDA)
    ax.spines['bottom'].set_linewidth(0.8)

    ax.tick_params(axis='y', which='both', length=0, labelsize=10, labelcolor=COR_TEXTO)
    ax.tick_params(axis='x', which='both', length=3, color=COR_BORDA, labelsize=10, labelcolor=COR_TEXTO)

    plt.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# BARRA LATERAL (SIDEBAR)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🌡️ Painel de Controle")
    st.write("**Projeto: TCC - BOSQUE**")
    st.markdown("---")

    visualizacao = st.radio(
        "Modo de visualização:",
        ["Dia Individual", "Todos os Dias (Painel)", "Tabela Comparativa"]
    )

    st.markdown("---")
    st.subheader("Configurações do Gráfico")
    com_marcadores = st.checkbox("Exibir marcadores nos pontos", value=False)
    mostrar_tabela = st.checkbox("Exibir tabela de dados brutos", value=True)
    mostrar_estatisticas = st.checkbox("Exibir métricas e estatísticas", value=True)

    st.markdown("---")
    st.subheader("📥 Downloads")

    # Gerar a planilha Excel sob demanda ou fornecer a existente
    excel_path = "dados_e_graficos_bosque.xlsx"
    if not os.path.exists(excel_path):
        gerar_planilha_excel(excel_path)

    with open(excel_path, "rb") as f:
        st.download_button(
            label="📊 Baixar Planilha Excel (.xlsx)",
            data=f.read(),
            file_name="dados_e_graficos_bosque.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# ---------------------------------------------------------------------------
# CONTEÚDO PRINCIPAL
# ---------------------------------------------------------------------------
st.title("🌳 Análise de Temperatura do Ar - BOSQUE")
st.caption("Visualização e análise de dados microclimáticos coletados nos pontos P1 a P10.")

if visualizacao == "Dia Individual":
    col_sel1, col_sel2 = st.columns([1, 2])

    with col_sel1:
        dia_selecionado = st.selectbox(
            "Selecione o Dia:",
            options=["14", "16", "19"],
            format_func=lambda x: DADOS[x]["dia_nome"]
        )

    with col_sel2:
        periodos_selecionados = st.multiselect(
            "Períodos do dia:",
            options=["Manha", "Tarde", "Noite"],
            default=["Manha", "Tarde", "Noite"]
        )

    if not periodos_selecionados:
        st.warning("Selecione pelo menos um período (Manhã, Tarde ou Noite) para exibir o gráfico.")
    else:
        info_dia = DADOS[dia_selecionado]
        df_dia = criar_dataframe(dia_selecionado)

        # Métricas em destaque
        if mostrar_estatisticas:
            temp_max = df_dia[periodos_selecionados].max().max()
            temp_min = df_dia[periodos_selecionados].min().min()
            temp_media = df_dia[periodos_selecionados].mean().mean()

            m1, m2, m3 = st.columns(3)
            m1.metric("Temperatura Média", f"{temp_media:.1f} °C")
            m2.metric("Temperatura Máxima", f"{temp_max:.1f} °C")
            m3.metric("Temperatura Mínima", f"{temp_min:.1f} °C")

        # Gráfico
        fig = render_grafico(df_dia, info_dia, periodos_selecionados, com_marcadores)
        st.pyplot(fig)

        # Botão para baixar imagem do gráfico atual
        buf_png = io.BytesIO()
        fig.savefig(buf_png, format="png", dpi=300, bbox_inches="tight")
        st.download_button(
            label=f"📥 Baixar Gráfico ({info_dia['dia_nome']}) em Alta Resolução (PNG 300 DPI)",
            data=buf_png.getvalue(),
            file_name=f"grafico_dia_{dia_selecionado}.png",
            mime="image/png"
        )
        plt.close(fig)

        # Tabela de dados
        if mostrar_tabela:
            st.markdown("### 📋 Tabela de Dados")
            cols_exibir = ['Pontos'] + periodos_selecionados
            st.dataframe(df_dia[cols_exibir].style.format({c: "{:.1f} °C" for c in periodos_selecionados}), use_container_width=True)

elif visualizacao == "Todos os Dias (Painel)":
    st.subheader("Painel Comparativo - Todos os Dias")
    st.markdown("Visualização sequencial dos três dias de medições no BOSQUE:")

    tabs = st.tabs(["14 de Agosto", "16 de Agosto", "19 de Agosto", "Painel Completo"])

    for idx, (dia_key, tab) in enumerate(zip(["14", "16", "19"], tabs[:3])):
        with tab:
            info = DADOS[dia_key]
            df = criar_dataframe(dia_key)
            fig = render_grafico(df, info, ["Manha", "Tarde", "Noite"], com_marcadores)
            st.pyplot(fig)
            plt.close(fig)

            if mostrar_tabela:
                st.dataframe(df.style.format({c: "{:.1f} °C" for c in ["Manha", "Tarde", "Noite"]}), use_container_width=True)

    with tabs[3]:
        painel_path = "graficos_gerados/graficos_todos_dias.png"
        if os.path.exists(painel_path):
            st.image(painel_path, caption="Painel Integrado - 14, 16 e 19 de Agosto", use_container_width=True)
            with open(painel_path, "rb") as f:
                st.download_button(
                    label="📥 Baixar Painel Completo em PNG (300 DPI)",
                    data=f.read(),
                    file_name="graficos_todos_dias.png",
                    mime="image/png"
                )

elif visualizacao == "Tabela Comparativa":
    st.subheader("Resumo Estatístico por Período e Dia")
    
    resumo_list = []
    for dia_key, info in DADOS.items():
        df = criar_dataframe(dia_key)
        for p in ["Manha", "Tarde", "Noite"]:
            resumo_list.append({
                "Dia": info["dia_nome"],
                "Período": p,
                "Média (°C)": round(df[p].mean(), 2),
                "Mínima (°C)": round(df[p].min(), 2),
                "Máxima (°C)": round(df[p].max(), 2),
                "Amplitude (°C)": round(df[p].max() - df[p].min(), 2)
            })

    df_resumo = pd.DataFrame(resumo_list)
    st.dataframe(df_resumo, use_container_width=True)
