import io
import os
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO GERAL E ESTILO VZN
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="VZN EcoClima - Studio de Análise Térmica",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta visual fiel ao modelo original
COR_MANHA = '#265b79'   # Azul petróleo
COR_TARDE = '#e27641'   # Laranja terracota
COR_NOITE = '#236532'   # Verde escuro
COR_TEXTO = '#595959'   # Cinza escuro para títulos e eixos
COR_GRADE = '#e6e6e6'   # Cinza claro para linhas de grade
COR_BORDA = '#d9d9d9'   # Cinza suave para contornos

# Fontes
plt.rcParams['font.sans-serif'] = ['Cambria', 'Times New Roman', 'DejaVu Serif', 'serif']
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['axes.edgecolor'] = COR_BORDA
plt.rcParams['axes.linewidth'] = 0.8

# Dados pré-carregados padrão
PRESETS = {
    "Dia 14 de Agosto": {
        "titulo_dia": "Dia 14 de agosto",
        "titulo_grafico": "DIA 14 DE AGOSTO - BOSQUE",
        "min_y": 24.0,
        "max_y": 32.0,
        "df": pd.DataFrame({
            'Pontos': ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10'],
            'Manha': [26.2, 26.5, 26.5, 26.9, 27.2, 27.4, 27.1, 27.2, 27.3, 27.6],
            'Tarde': [28.5, 29.0, 29.5, 29.6, 30.4, 29.9, 30.9, 31.4, 30.3, 30.5],
            'Noite': [25.5, 26.1, 26.2, 26.4, 26.9, 26.7, 26.8, 26.9, 27.0, 27.1]
        })
    },
    "Dia 16 de Agosto": {
        "titulo_dia": "Dia 16 de agosto",
        "titulo_grafico": "DIA 16 DE AGOSTO - BOSQUE",
        "min_y": 24.0,
        "max_y": 32.0,
        "df": pd.DataFrame({
            'Pontos': ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10'],
            'Manha': [26.0, 25.3, 25.3, 26.6, 26.2, 25.5, 25.4, 26.2, 26.0, 26.3],
            'Tarde': [28.5, 28.6, 28.9, 31.0, 30.8, 29.5, 30.4, 30.6, 30.0, 29.9],
            'Noite': [25.8, 26.0, 26.1, 26.7, 26.8, 27.0, 27.0, 26.9, 26.6, 27.1]
        })
    },
    "Dia 19 de Agosto": {
        "titulo_dia": "Dia 19 de agosto",
        "titulo_grafico": "DIA 19 DE AGOSTO - BOSQUE",
        "min_y": 22.0,
        "max_y": 32.0,
        "df": pd.DataFrame({
            'Pontos': ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10'],
            'Manha': [24.7, 24.0, 23.4, 23.6, 23.8, 23.6, 23.5, 23.8, 24.2, 24.4],
            'Tarde': [28.2, 28.5, 28.9, 31.5, 31.0, 28.9, 31.0, 31.8, 30.0, 30.6],
            'Noite': [27.1, 25.5, 25.3, 26.7, 26.2, 26.1, 25.8, 27.0, 27.1, 25.6]
        })
    }
}


def sanitizar_dados_tabela(df):
    """
    Corrige automaticamente os erros clássicos de colagem do Excel em português:
    1. Vírgula que virou zero/centena ao colar (ex: 27,4 -> 274.0 vira 27.4).
    2. Valores None/vazios na coluna Pontos quando linhas extras são coladas.
    """
    modificado = False
    df_novo = df.copy()

    # 1. Garantir que a coluna Pontos não tenha valores None/vazios
    for idx in range(len(df_novo)):
        val_ponto = df_novo.iloc[idx]['Pontos']
        if pd.isna(val_ponto) or str(val_ponto).strip().lower() in ['none', 'nan', '']:
            df_novo.iloc[idx, df_novo.columns.get_loc('Pontos')] = f"P{idx + 1}"
            modificado = True

    # 2. Corrigir números onde a vírgula foi engolida pelo navegador (ex: 274.0 -> 27.4)
    for col in ['Manha', 'Tarde', 'Noite']:
        if col in df_novo.columns:
            for idx in range(len(df_novo)):
                val = df_novo.iloc[idx][col]
                if pd.notnull(val):
                    try:
                        v_float = float(val)
                        # Valores entre 60.0 e 600.0 são temperaturas sem vírgula (ex: 274.0 -> 27.4)
                        if 60.0 < v_float <= 600.0:
                            df_novo.iloc[idx, df_novo.columns.get_loc(col)] = round(v_float / 10.0, 1)
                            modificado = True
                        elif 600.0 < v_float <= 6000.0:
                            df_novo.iloc[idx, df_novo.columns.get_loc(col)] = round(v_float / 100.0, 1)
                            modificado = True
                    except (ValueError, TypeError):
                        pass

    return df_novo, modificado


def gerar_figura_matplotlib(df, titulo_grafico, min_y=None, max_y=None, com_marcadores=False):
    """Gera a figura Matplotlib com estética idêntica ao modelo original."""
    fig, ax = plt.subplots(figsize=(6.8, 3.6), dpi=200)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    pontos = df['Pontos'].astype(str).tolist()
    m_marker = 'o' if com_marcadores else None
    t_marker = 's' if com_marcadores else None
    n_marker = '^' if com_marcadores else None

    # Linhas dos períodos
    if 'Manha' in df.columns:
        ax.plot(pontos, df['Manha'], label='Manha', color=COR_MANHA, linewidth=2.0, marker=m_marker, markersize=5)
    if 'Tarde' in df.columns:
        ax.plot(pontos, df['Tarde'], label='Tarde', color=COR_TARDE, linewidth=2.0, marker=t_marker, markersize=5)
    if 'Noite' in df.columns:
        ax.plot(pontos, df['Noite'], label='Noite', color=COR_NOITE, linewidth=2.0, marker=n_marker, markersize=5)

    ax.set_title(titulo_grafico, fontsize=12, color=COR_TEXTO, pad=25)

    leg = ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.15),
                    ncol=3, frameon=False, fontsize=9.5,
                    handlelength=1.5, handletextpad=0.5, columnspacing=1.5)
    if leg:
        for text in leg.get_texts():
            text.set_color(COR_TEXTO)

    ax.set_ylabel("Temperatura do ar (°C)", fontsize=9.5, color=COR_TEXTO, labelpad=8)
    ax.set_xlabel("Pontos", fontsize=9.5, color=COR_TEXTO, labelpad=8)

    # Configuração de limites Y
    todas_temps = []
    for col in ['Manha', 'Tarde', 'Noite']:
        if col in df.columns:
            todas_temps.extend(pd.to_numeric(df[col], errors='coerce').dropna().tolist())

    if todas_temps:
        real_min = min(todas_temps)
        real_max = max(todas_temps)
        pad_min = (int(real_min // 2) * 2) if min_y is None else min_y
        pad_max = (int(np.ceil(real_max / 2)) * 2) if max_y is None else max_y
        if pad_max <= pad_min:
            pad_max = pad_min + 2
        ax.set_ylim(pad_min, pad_max)
        ax.yaxis.set_major_locator(ticker.MultipleLocator(2))

    ax.grid(axis='y', color=COR_GRADE, linestyle='-', linewidth=0.8)
    ax.grid(axis='x', visible=False)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color(COR_BORDA)
    ax.spines['bottom'].set_linewidth(0.8)

    ax.tick_params(axis='y', which='both', length=0, labelsize=9.5, labelcolor=COR_TEXTO)
    ax.tick_params(axis='x', which='both', length=3, color=COR_BORDA, labelsize=9.5, labelcolor=COR_TEXTO)

    plt.tight_layout()
    return fig


def gerar_excel_bytes(df, titulo_dia, titulo_grafico, min_y, max_y):
    """Gera uma planilha Excel em memória (.xlsx) com a tabela e o gráfico nativo ao lado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados e Grafico"

    borda_fina = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )

    # Cabeçalho mesclado do dia
    ws.merge_cells("B2:E2")
    header = ws.cell(row=2, column=2, value=titulo_dia)
    header.font = Font(name="Cambria", size=11, bold=True)
    header.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(2, 6):
        ws.cell(row=2, column=c).border = borda_fina

    # Nomes das colunas
    colunas = ["Pontos", "Manha", "Tarde", "Noite"]
    for idx, col in enumerate(colunas, start=2):
        c = ws.cell(row=3, column=idx, value=col)
        c.font = Font(name="Cambria", size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borda_fina

    # Linhas de dados
    for r_idx, row in df.iterrows():
        r = 4 + r_idx
        ws.cell(row=r, column=2, value=str(row['Pontos'])).alignment = Alignment(horizontal="center")
        for col_idx, col_name in enumerate(["Manha", "Tarde", "Noite"], start=3):
            val = row.get(col_name, np.nan)
            cell = ws.cell(row=r, column=col_idx, value=float(val) if pd.notnull(val) else None)
            cell.alignment = Alignment(horizontal="center")
            cell.border = borda_fina
            if cell.value is not None:
                cell.number_format = '0.0'
        ws.cell(row=r, column=2).border = borda_fina

    fim_dados = 3 + len(df)

    # Adicionar Gráfico Nativo do Excel ao lado (coluna G)
    chart = LineChart()
    chart.title = titulo_grafico
    chart.style = 10
    chart.y_axis.title = "Temperatura do ar (°C)"
    chart.x_axis.title = "Pontos"
    if min_y is not None:
        chart.y_axis.scaling.min = min_y
    if max_y is not None:
        chart.y_axis.scaling.max = max_y
    chart.y_axis.majorUnit = 2
    chart.width = 16
    chart.height = 9.5

    dados_ref = Reference(ws, min_col=3, min_row=3, max_col=5, max_row=fim_dados)
    cats_ref = Reference(ws, min_col=2, min_row=4, max_row=fim_dados)
    chart.add_data(dados_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "G2")

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 11
    ws.column_dimensions['C'].width = 11
    ws.column_dimensions['D'].width = 11
    ws.column_dimensions['E'].width = 11
    ws.column_dimensions['F'].width = 4

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# BARRA LATERAL (VZN BRANDING & CONFIGURAÇÕES)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 🌿 **VZN EcoClima**")
    st.caption("Studio de Visualização Microclimática • TCC")
    st.markdown("---")

    st.subheader("⚡ Modelos Prontos")
    preset_escolhido = st.selectbox(
        "Carregar dados pré-definidos:",
        ["Criar Novo / Em Branco"] + list(PRESETS.keys())
    )

    st.markdown("---")
    st.subheader("📁 Importar Arquivo de Planilha")
    uploaded_file = st.file_uploader(
        "Carregar arquivo (.xlsx ou .csv):",
        type=["xlsx", "csv"],
        help="A planilha deve conter colunas: Pontos, Manha, Tarde, Noite"
    )

    st.markdown("---")
    st.subheader("🎨 Opções Visuais")
    opt_marcadores = st.checkbox("Exibir pontos/marcadores na linha", value=False)
    ajustar_escala = st.checkbox("Definir escala Y manualmente", value=True)

# ---------------------------------------------------------------------------
# INICIALIZAÇÃO DO ESTADO DA SESSÃO (SESSION STATE)
# ---------------------------------------------------------------------------
if "editor_version" not in st.session_state:
    st.session_state.editor_version = 0

if "ultimo_preset" not in st.session_state:
    st.session_state.ultimo_preset = "Dia 14 de Agosto"
    st.session_state.titulo_dia = PRESETS["Dia 14 de Agosto"]["titulo_dia"]
    st.session_state.titulo_grafico = PRESETS["Dia 14 de Agosto"]["titulo_grafico"]
    st.session_state.min_y = PRESETS["Dia 14 de Agosto"]["min_y"]
    st.session_state.max_y = PRESETS["Dia 14 de Agosto"]["max_y"]
    st.session_state.df_atual = PRESETS["Dia 14 de Agosto"]["df"].copy()

# Se o usuário trocou o preset no seletor
if preset_escolhido != st.session_state.ultimo_preset:
    st.session_state.ultimo_preset = preset_escolhido
    st.session_state.editor_version += 1
    if preset_escolhido in PRESETS:
        p = PRESETS[preset_escolhido]
        st.session_state.titulo_dia = p["titulo_dia"]
        st.session_state.titulo_grafico = p["titulo_grafico"]
        st.session_state.min_y = p["min_y"]
        st.session_state.max_y = p["max_y"]
        st.session_state.df_atual = p["df"].copy()
    else:
        st.session_state.titulo_dia = "Novo Dia"
        st.session_state.titulo_grafico = "NOVO DIA - BOSQUE"
        st.session_state.min_y = 20.0
        st.session_state.max_y = 35.0
        st.session_state.df_atual = pd.DataFrame({
            'Pontos': [f"P{i}" for i in range(1, 11)],
            'Manha': [25.0] * 10,
            'Tarde': [30.0] * 10,
            'Noite': [26.0] * 10
        })

# Se o usuário fez upload de uma planilha
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_up = pd.read_csv(uploaded_file)
        else:
            df_up = pd.read_excel(uploaded_file)

        col_map = {}
        for c in df_up.columns:
            c_clean = str(c).strip().lower()
            if 'ponto' in c_clean:
                col_map[c] = 'Pontos'
            elif 'manh' in c_clean or 'manha' in c_clean:
                col_map[c] = 'Manha'
            elif 'tarde' in c_clean:
                col_map[c] = 'Tarde'
            elif 'noite' in c_clean:
                col_map[c] = 'Noite'

        df_up = df_up.rename(columns=col_map)
        colunas_necessarias = ['Pontos', 'Manha', 'Tarde', 'Noite']
        if all(c in df_up.columns for c in colunas_necessarias):
            st.session_state.df_atual = df_up[colunas_necessarias].copy()
            st.session_state.editor_version += 1
            nome_base = os.path.splitext(uploaded_file.name)[0]
            st.session_state.titulo_dia = nome_base
            st.session_state.titulo_grafico = f"{nome_base.upper()} - BOSQUE"
            st.sidebar.success(f"Planilha '{uploaded_file.name}' carregada com sucesso!")
        else:
            st.sidebar.error("A planilha precisa ter as colunas: Pontos, Manha, Tarde e Noite.")
    except Exception as e:
        st.sidebar.error(f"Erro ao ler planilha: {e}")

# ---------------------------------------------------------------------------
# CABEÇALHO PRINCIPAL
# ---------------------------------------------------------------------------
st.title("🌿 VZN EcoClima — Studio de Análise Térmica")
st.markdown(
    "Edite os valores na **tabela à esquerda** (ou cole direto do Excel) e veja o **gráfico ao lado atualizando instantaneamente em tempo real**."
)
st.markdown("---")

# ---------------------------------------------------------------------------
# ÁREA DE TRABALHO: TABELA E GRÁFICO LADO A LADO
# ---------------------------------------------------------------------------
col_tabela, col_grafico = st.columns([1, 1.2], gap="large")

with col_tabela:
    st.subheader("📋 Tabela de Dados (Editável)")

    # Ferramenta para Colar Direto do Excel sem erro de vírgula
    with st.expander("📋 **Colar Dados do Excel (Ctrl+C no Excel -> Ctrl+V aqui)**", expanded=False):
        st.caption("Copie as células no Excel e cole no campo abaixo. As vírgulas são convertidas automaticamente!")
        
        destino_colar = st.selectbox(
            "Onde colar?",
            ["Apenas Manhã", "Apenas Tarde", "Apenas Noite", "Manhã, Tarde e Noite (3 colunas)", "Tabela Completa (Pontos + 3 colunas)"]
        )
        texto_colado = st.text_area("Cole aqui (Ctrl+V):", height=90, placeholder="Exemplo:\n27,4\n28,0\n27,6\n...")

        if st.button("📥 Aplicar Dados na Tabela", use_container_width=True):
            if texto_colado.strip():
                linhas_raw = [l.strip() for l in texto_colado.strip().split('\n') if l.strip()]
                df_temp = st.session_state.df_atual.copy()

                if destino_colar in ["Apenas Manhã", "Apenas Tarde", "Apenas Noite"]:
                    col_alvo = "Manha" if "Manhã" in destino_colar else ("Tarde" if "Tarde" in destino_colar else "Noite")
                    valores = []
                    for l in linhas_raw:
                        val_str = l.replace(',', '.').strip()
                        try:
                            v = float(val_str)
                            if 60.0 < v <= 600.0:
                                v = v / 10.0
                            valores.append(v)
                        except ValueError:
                            valores.append(np.nan)

                    # Ajustar tamanho do DataFrame se necessário
                    if len(valores) > len(df_temp):
                        linhas_extras = len(valores) - len(df_temp)
                        novos_pontos = [f"P{len(df_temp) + i + 1}" for i in range(linhas_extras)]
                        df_extra = pd.DataFrame({'Pontos': novos_pontos, 'Manha': [np.nan]*linhas_extras, 'Tarde': [np.nan]*linhas_extras, 'Noite': [np.nan]*linhas_extras})
                        df_temp = pd.concat([df_temp, df_extra], ignore_index=True)

                    for i, v in enumerate(valores):
                        df_temp.at[i, col_alvo] = v

                elif destino_colar == "Manhã, Tarde e Noite (3 colunas)":
                    novas_linhas = []
                    for l in linhas_raw:
                        partes = [p.strip().replace(',', '.') for p in l.replace(';', '\t').split('\t')]
                        nums = []
                        for p in partes[:3]:
                            try:
                                v = float(p)
                                if 60.0 < v <= 600.0:
                                    v = v / 10.0
                                nums.append(v)
                            except ValueError:
                                nums.append(np.nan)
                        while len(nums) < 3:
                            nums.append(np.nan)
                        novas_linhas.append(nums)

                    df_temp = pd.DataFrame(novas_linhas, columns=['Manha', 'Tarde', 'Noite'])
                    df_temp.insert(0, 'Pontos', [f"P{i+1}" for i in range(len(df_temp))])

                elif destino_colar == "Tabela Completa (Pontos + 3 colunas)":
                    novas_linhas = []
                    for l in linhas_raw:
                        partes = [p.strip().replace(',', '.') for p in l.replace(';', '\t').split('\t')]
                        ponto_nome = partes[0]
                        nums = []
                        for p in partes[1:4]:
                            try:
                                v = float(p)
                                if 60.0 < v <= 600.0:
                                    v = v / 10.0
                                nums.append(v)
                            except ValueError:
                                nums.append(np.nan)
                        while len(nums) < 3:
                            nums.append(np.nan)
                        novas_linhas.append([ponto_nome] + nums)

                    df_temp = pd.DataFrame(novas_linhas, columns=['Pontos', 'Manha', 'Tarde', 'Noite'])

                st.session_state.df_atual = df_temp
                st.session_state.editor_version += 1
                st.success("Dados colados com sucesso e vírgulas convertidas!")
                st.rerun()

    t_col1, t_col2 = st.columns(2)
    with t_col1:
        titulo_dia_input = st.text_input("Título da Tabela:", value=st.session_state.titulo_dia)
    with t_col2:
        titulo_grafico_input = st.text_input("Título do Gráfico:", value=st.session_state.titulo_grafico)

    if ajustar_escala:
        e_col1, e_col2 = st.columns(2)
        with e_col1:
            min_y_val = st.number_input("Escala Y Mínima (°C):", value=float(st.session_state.min_y), step=1.0)
        with e_col2:
            max_y_val = st.number_input("Escala Y Máxima (°C):", value=float(st.session_state.max_y), step=1.0)
    else:
        min_y_val = None
        max_y_val = None

    # Botões auxiliares rápidos
    b_col1, b_col2 = st.columns(2)
    if b_col1.button("🔄 Renumerar Pontos (P1..Pn)", use_container_width=True):
        st.session_state.df_atual['Pontos'] = [f"P{i+1}" for i in range(len(st.session_state.df_atual))]
        st.session_state.editor_version += 1
        st.rerun()

    if b_col2.button("🪄 Corrigir Vírgulas/Zeros", use_container_width=True, help="Divide por 10 automaticamente qualquer valor > 60°C (ex: 274 -> 27.4)"):
        df_corr, _ = sanitizar_dados_tabela(st.session_state.df_atual)
        st.session_state.df_atual = df_corr
        st.session_state.editor_version += 1
        st.rerun()

    st.caption("💡 *Dica: você pode colar células diretamente aqui. Se o navegador engolir a vírgula (ex: 27,4 virar 274), o sistema corrige sozinho automaticamente!*")

    # Editor de dados interativo com chave versionada
    editor_key = f"editor_principal_{st.session_state.editor_version}"
    df_bruto = st.data_editor(
        st.session_state.df_atual,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Pontos": st.column_config.TextColumn("Pontos", width="small", required=True),
            "Manha": st.column_config.NumberColumn("Manha (°C)", format="%.1f", step=0.1),
            "Tarde": st.column_config.NumberColumn("Tarde (°C)", format="%.1f", step=0.1),
            "Noite": st.column_config.NumberColumn("Noite (°C)", format="%.1f", step=0.1),
        },
        height=380,
        key=editor_key
    )

    # Auto-correção em tempo real se a usuária colar números com vírgula do Excel que o browser transformou em centena (ex: 274 -> 27.4)
    df_sanitizado, foi_corrigido = sanitizar_dados_tabela(df_bruto)
    if foi_corrigido:
        st.session_state.df_atual = df_sanitizado
        st.session_state.editor_version += 1
        st.toast("🪄 Vírgula do Excel corrigida automaticamente (ex: 274 -> 27.4)!", icon="✨")
        st.rerun()
    else:
        df_editado = df_sanitizado

with col_grafico:
    st.subheader("📈 Gráfico em Tempo Real (Preview)")

    # Gera a figura sincronizada com os dados editados
    fig_preview = gerar_figura_matplotlib(
        df=df_editado,
        titulo_grafico=titulo_grafico_input,
        min_y=min_y_val,
        max_y=max_y_val,
        com_marcadores=opt_marcadores
    )
    st.pyplot(fig_preview)

    # -----------------------------------------------------------------------
    # BOTÕES DE EXPORTAÇÃO E DOWNLOAD
    # -----------------------------------------------------------------------
    st.markdown("##### 📥 Opções de Exportação:")
    btn_col1, btn_col2, btn_col3 = st.columns(3)

    # 1. Download PNG 300 DPI
    buf_png = io.BytesIO()
    fig_preview.savefig(buf_png, format="png", dpi=300, bbox_inches="tight")
    btn_col1.download_button(
        label="🖼️ Baixar Gráfico PNG (300 DPI)",
        data=buf_png.getvalue(),
        file_name=f"{titulo_dia_input.replace(' ', '_').lower()}_grafico.png",
        mime="image/png",
        use_container_width=True
    )

    # 2. Download PDF Vetorial
    buf_pdf = io.BytesIO()
    fig_preview.savefig(buf_pdf, format="pdf", bbox_inches="tight")
    btn_col2.download_button(
        label="📄 Baixar Gráfico em PDF",
        data=buf_pdf.getvalue(),
        file_name=f"{titulo_dia_input.replace(' ', '_').lower()}_grafico.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    # 3. Download Planilha Excel Completa (.xlsx)
    excel_bytes = gerar_excel_bytes(
        df=df_editado,
        titulo_dia=titulo_dia_input,
        titulo_grafico=titulo_grafico_input,
        min_y=min_y_val,
        max_y=max_y_val
    )
    btn_col3.download_button(
        label="📊 Exportar para Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"{titulo_dia_input.replace(' ', '_').lower()}_com_grafico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    plt.close(fig_preview)

# ---------------------------------------------------------------------------
# RESUMO ESTATÍSTICO INFORMATIVO
# ---------------------------------------------------------------------------
st.markdown("---")
with st.expander("📊 Ver Resumo Estatístico dos Dados Atuais", expanded=False):
    s_col1, s_col2, s_col3 = st.columns(3)
    for idx, (periodo, s_col) in enumerate(zip(['Manha', 'Tarde', 'Noite'], [s_col1, s_col2, s_col3])):
        if periodo in df_editado.columns:
            serie = pd.to_numeric(df_editado[periodo], errors='coerce').dropna()
            if not serie.empty:
                s_col.markdown(f"#### Período: **{periodo}**")
                s_col.write(f"- **Média:** {serie.mean():.2f} °C")
                s_col.write(f"- **Mínima:** {serie.min():.2f} °C")
                s_col.write(f"- **Máxima:** {serie.max():.2f} °C")
                s_col.write(f"- **Amplitude:** {(serie.max() - serie.min()):.2f} °C")
