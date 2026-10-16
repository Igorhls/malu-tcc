# 🚀 Vespera Sci — Documento de Contexto e Arquitetura para Agentes de IA (Handoff Blueprint)

> **Instrução para o Agente de IA:**  
> Leia este documento integralmente antes de propor ou executar qualquer ação. Ele contém todo o histórico, regras de negócio, dados experimentais, especificações gráficas e o plano de trabalho acordado com o usuário para a construção do software desktop **Vespera Sci**.

---

## ⚠️ REGRA DE OURO DE INTERAÇÃO COM O USUÁRIO
1. **Consulte sempre antes de executar:** O usuário solicitou expressamente ser consultado **antes** do início ou execução de cada etapa. Apresente o que será feito, aguarde o "de acordo" e só então gere/modifique os arquivos.
2. **Idioma:** Português (Brasil).
3. **Repositórios:**
   * **Repositório Legado/Web (Streamlit — Mantido intacto para o TCC):** `https://github.com/Igorhls/malu-tcc.git`
   * **Novo Repositório Oficial (Aplicação Desktop — Vespera Sci):** `https://github.com/Igorhls/vespera-sci.git`

---

## 1. Visão Geral e Origem do Projeto

O projeto nasceu como **VZN EcoClima** (`malu-tcc`), uma aplicação em Streamlit desenvolvida para auxiliar no Trabalho de Conclusão de Curso (TCC) da **Malu** (Monitoramento Microclimático: análise de temperatura do ar no Bosque vs. Estacionamento).

Com a evolução das funcionalidades — controle milimétrico de estilo padrão *OriginLab Pro*, múltiplos tipos de gráficos, colagem inteligente do Excel, gerenciador de projetos e exportação nativa para `.xlsx`, `.pdf` e `.png` (300 DPI) — o arquivo `app.py` atingiu mais de **2.200 linhas**, esbarrando nas limitações arquiteturais do Streamlit (reexecução completa do script a cada interação, ausência de arraste nativo com o mouse na figura e dependência de navegador/terminal).

Daí surge o **Vespera Sci**:
* Um **software desktop científico local, modular e expansível**, construído em **Python + PySide6 (Qt for Python)** com **Matplotlib interativo (`FigureCanvasQTAgg`)**.
* Projetado para começar atendendo 100% das demandas de gráficos do TCC da Malu, mas com arquitetura pronta para crescer como uma ferramenta geral de plotagem científica (estilo *OriginLab* / *GraphPad Prism*).
* Possui um **sistema de atualização automática via GitHub**: a usuária instala/baixa na máquina dela e, através de um botão **"Verificar Atualizações"** na interface, o programa busca as novidades que o desenvolvedor subiu no repositório `https://github.com/Igorhls/vespera-sci.git` e se atualiza automaticamente.

---

## 2. Plano de Execução em 3 Etapas (Acordado com o Usuário)

O desenvolvimento do **Vespera Sci** está dividido em 3 etapas sequenciais (sempre validando com o usuário antes de cada uma):

### Etapa 1: Setup do Repositório `vespera-sci` e Arquitetura Modular
* Inicialização da estrutura de diretórios limpa no repositório `https://github.com/Igorhls/vespera-sci.git`.
* Criação do `README.md` oficial do **Vespera Sci** (modelo completo incluído na Seção 6 deste documento).
* Separação do motor matemático/gráfico (`core/`) da camada visual (`ui/`).

### Etapa 2: Construção do Aplicativo Desktop Nativo (PySide6 + Matplotlib)
* Janela principal moderna e sóbria (compatível com tema claro/escuro).
* Tabela editável com suporte a `Ctrl+C` / `Ctrl+V` direto do Excel e sanitização automática de decimais brasileiros.
* Canvas do Matplotlib integrado ao Qt com suporte a **legenda arrastável diretamente com o mouse** (`leg.set_draggable(True)`), além de controles finos na barra lateral.
* Suporte completo aos modos de gráficos já existentes (Linhas por Pontos $P_1\text{–}P_n$ e Barras Agrupadas de Média e Desvio Padrão) + arquitetura aberta para novos tipos (Boxplot, Dispersão, etc.).
* Exportação imediata em PNG (300/600 DPI), PDF Vetorial, SVG e Planilha Excel (`.xlsx` com gráficos nativos via `openpyxl`).
* Sistema de salvar/abrir projetos (`.vespera` / `.json`).

### Etapa 3: Sistema de Auto-Update via GitHub e Empacotamento Local
* Implementação do módulo `core/updater.py` conectado ao repositório `Igorhls/vespera-sci`.
* Botão na interface **"🔄 Verificar Atualizações"** com barra de progresso:
  * Consulta a API do GitHub (`releases/latest` ou arquivo de versão/zip da branch `main`), baixa os arquivos novos, aplica a atualização e reinicia o aplicativo automaticamente.
* Configuração de script de empacotamento (PyInstaller ou pacote portátil com lançador `.exe` / `.bat`) para que a usuária final rode com 2 cliques sem precisar usar terminal.

---

## 3. Arquitetura Modular Recomendada para o `vespera-sci`

```text
vespera-sci/
├── main.py                         # Ponto de entrada da aplicação Desktop (PySide6)
├── version.json                    # Versão atual do sistema (ex: {"version": "1.0.0"})
├── requirements.txt                # PySide6, matplotlib, pandas, numpy, openpyxl, requests
├── README.md                       # Documentação oficial do Vespera Sci
│
├── config/
│   ├── __init__.py
│   ├── presets.py                  # Dados pré-definidos (Dias 14, 16, 19 Ago + TA Barras)
│   └── themes.py                   # Paletas (OriginLab Pro, Pastel, Nature, Monocromático) e Fontes
│
├── core/
│   ├── __init__.py
│   ├── sanitizer.py                # Limpeza de dados colados do Excel e formatação PT-BR
│   ├── plot_engine.py              # Funções puras de geração das figuras Matplotlib (Linhas e Barras)
│   ├── excel_exporter.py           # Geração de planilhas .xlsx com gráficos nativos (openpyxl)
│   ├── project_manager.py          # Leitura/Gravação de projetos (.vespera / .json)
│   └── updater.py                  # Cliente de atualização automática via GitHub (Igorhls/vespera-sci)
│
└── ui/
    ├── __init__.py
    ├── main_window.py              # Janela Principal (Menus, Barra de Projetos, Abas/Splitters)
    ├── sidebar_controls.py         # Painel lateral de estilo (Fontes, Cores, Bordas, Ticks, Legenda)
    ├── data_table_widget.py        # Tabela interativa com colagem rápida do Excel e calculadora
    ├── plot_canvas_widget.py       # FigureCanvasQTAgg com legenda arrastável no mouse e barra de zoom
    └── update_dialog.py            # Modal de verificação e progresso de atualização do GitHub
```

---

## 4. Especificações Técnicas e Regras de Negócio (Extraídas do Sistema Atual)

Para garantir compatibilidade total com tudo o que já foi construído no `malu-tcc`, o agente deve preservar as seguintes especificações exatas:

### 4.1. Dados Pré-Cadastrados (Presets)

#### A) Modo Linhas (`PRESETS_LINHAS` — Pontos `P1` a `P10`, colunas `Pontos`, `Manha`, `Tarde`, `Noite`)
1. **Dia 14 de Agosto** (`titulo_dia`: `"Dia 14 de agosto"`, `titulo_grafico`: `"DIA 14 DE AGOSTO - BOSQUE"`, `min_y`: `24.0`, `max_y`: `33.0`):
   * `Manha`: `[26.2, 26.5, 26.5, 26.9, 27.2, 27.4, 27.1, 27.2, 27.3, 27.6]`
   * `Tarde`: `[28.5, 29.0, 29.5, 29.6, 30.4, 29.9, 30.9, 31.4, 30.3, 30.5]`
   * `Noite`: `[25.5, 26.1, 26.2, 26.4, 26.9, 26.7, 26.8, 26.9, 27.0, 27.1]`
2. **Dia 16 de Agosto** (`titulo_dia`: `"Dia 16 de agosto"`, `titulo_grafico`: `"DIA 16 DE AGOSTO - BOSQUE"`, `min_y`: `24.0`, `max_y`: `33.0`):
   * `Manha`: `[26.0, 25.3, 25.3, 26.6, 26.2, 25.5, 25.4, 26.2, 26.0, 26.3]`
   * `Tarde`: `[28.5, 28.6, 28.9, 31.0, 30.8, 29.5, 30.4, 30.6, 30.0, 29.9]`
   * `Noite`: `[25.8, 26.0, 26.1, 26.7, 26.8, 27.0, 27.0, 26.9, 26.6, 27.1]`
3. **Dia 19 de Agosto** (`titulo_dia`: `"Dia 19 de agosto"`, `titulo_grafico`: `"DIA 19 DE AGOSTO - BOSQUE"`, `min_y`: `22.0`, `max_y`: `33.0`):
   * `Manha`: `[24.7, 24.0, 23.4, 23.6, 23.8, 23.6, 23.5, 23.8, 24.2, 24.4]`
   * `Tarde`: `[28.2, 28.5, 28.9, 31.5, 31.0, 28.9, 31.0, 31.8, 30.0, 30.6]`
   * `Noite`: `[27.1, 25.5, 25.3, 26.7, 26.2, 26.1, 25.8, 27.0, 27.1, 25.6]`

#### B) Modo Barras Agrupadas (`PRESETS_BARRAS` — Média e Desvio Padrão `TA`)
* **TA — Média e Desvio Padrão (Agosto e Dezembro)**:
  * `titulo_dia`: `"TA"`, `titulo_grafico`: `""`, `rotulo_x`: `"TA"`, `rotulo_y`: `"Média (%)"`, `min_y`: `0.0`, `max_y`: `35.0`
  * Colunas: `Mes`, `Bosque`, `Desvio_Bosque`, `Estacionamento`, `Desvio_Estacionamento`
  * Linha 1 (`Agosto`): `Bosque=27.32`, `Desvio_Bosque=2.30`, `Estacionamento=28.50`, `Desvio_Estacionamento=4.87`
  * Linha 2 (`Dezembro`): `Bosque=28.60`, `Desvio_Bosque=1.42`, `Estacionamento=29.40`, `Desvio_Estacionamento=2.19`

### 4.2. Paletas de Cores e Tipografia
* **Paletas de Linhas:**
  * `"OriginLab Pro (Preto, Vermelho, Azul)"`: `{"Manha": "#222222", "Tarde": "#e74c3c", "Noite": "#2980b9"}`
  * `"Acadêmico Monocromático (Escala de Cinza)"`: `{"Manha": "#111111", "Tarde": "#666666", "Noite": "#aaaaaa"}`
  * `"Nature / Science (Tons Frios & Quentes)"`: `{"Manha": "#2c3e50", "Tarde": "#d35400", "Noite": "#16a085"}`
  * `"Excel Clássico"`: `{"Manha": "#265b79", "Tarde": "#e27641", "Noite": "#236532"}`
* **Paletas de Barras:**
  * `"OriginLab Pastel (Padrão do Modelo)"`: `{"Bosque": "#fdc086", "Desvio_Bosque": "#8fd194", "Estacionamento": "#c2b2d6", "Desvio_Estacionamento": "#ffff99"}`
  * `"OriginLab Clássico (Cores Vibrantes)"`: `{"Bosque": "#f39c12", "Desvio_Bosque": "#27ae60", "Estacionamento": "#8e44ad", "Desvio_Estacionamento": "#f1c40f"}`
  * `"Acadêmico Monocromático (Escala de Cinza)"`: `{"Bosque": "#333333", "Desvio_Bosque": "#777777", "Estacionamento": "#aaaaaa", "Desvio_Estacionamento": "#dddddd"}`
  * `"Nature / Científico (Azul & Verde)"`: `{"Bosque": "#2b5c8f", "Desvio_Bosque": "#6baed6", "Estacionamento": "#31a354", "Desvio_Estacionamento": "#a1d99b"}`
* **Fontes Disponíveis:** `Arial`, `Times New Roman`, `Cambria`, `Garamond`, `Calibri`.

### 4.3. Regras de Sanitização ao Colar do Excel
* Em português, números como `27,4` muitas vezes são interpretados como `274` ao colar.
* Na tabela de **Linhas (Temperaturas)**: se `60.0 < valor <= 600.0`, dividir automaticamente por `10.0`; se `600.0 < valor <= 6000.0`, dividir por `100.0`.
* Na tabela de **Barras**: se `100.0 < valor <= 1000.0`, dividir por `10.0`; se `1000.0 < valor <= 10000.0`, dividir por `100.0`.
* Formatação de rótulos nas barras em padrão brasileiro: `formatar_numero_ptbr` substitui `.` por `,` e permite escolher entre decimais automáticos (`27,32` e `2,3`) ou 2 casas fixas (`27,32` e `2,30`).

---

## 5. Especificação do Sistema de Auto-Update (`core/updater.py`)

O usuário deseja que o programa instalado na máquina tenha um botão para atualizar conforme novas versões forem enviadas ao GitHub (`https://github.com/Igorhls/vespera-sci.git`).

### Como implementar de forma dupla e infalível:
1. **Verificação de Versão Remota:**
   * O app consulta `https://raw.githubusercontent.com/Igorhls/vespera-sci/main/version.json` (para atualizações rápidas de código direto da branch `main`) **OU** `https://api.github.com/repos/Igorhls/vespera-sci/releases/latest` (para binários `.exe` empacotados).
2. **Download e Aplicação:**
   * **Modo Código Portátil (Recomendado para agilidade máxima):** Baixa o `.zip` da branch `main` (`https://github.com/Igorhls/vespera-sci/archive/refs/heads/main.zip`), extrai os arquivos `.py` atualizados preservando os projetos salvos do usuário (`projetos_salvos.json`) e reinicia o processo Python via `os.execl(sys.executable, sys.executable, *sys.argv)`.
   * **Modo Executável Compilado (`.exe`):** Baixa o novo `.exe` da Release do GitHub, gera um `updater.bat` temporário que aguarda 2 segundos após o fechamento do app, substitui o `.exe` antigo pelo novo e reabre o programa.

---

## 6. Modelo Pronto de `README.md` para o Repositório `vespera-sci`

*(Ao iniciar a Etapa 1 no novo repositório, utilize o conteúdo abaixo como base para o `README.md` do `vespera-sci`)*:

```markdown
# 🌌 Vespera Sci — Scientific Plotting & Data Analysis Studio

**Vespera Sci** é uma plataforma desktop científica para análise de dados experimentais, estatística descritiva e produção de gráficos de alta precisão com padrão de publicação internacional (*OriginLab Pro*, *Nature*, *ABNT*).

Repositório Oficial: [https://github.com/Igorhls/vespera-sci.git](https://github.com/Igorhls/vespera-sci.git)

---

## ✨ Principais Recursos

- 🖥️ **Aplicação Desktop Nativa (PySide6 / Qt):** Desempenho instantâneo, funcionamento 100% local/offline e interface adaptativa.
- 🖱️ **Canvas Matplotlib Interativo:** Arraste a legenda diretamente com o mouse sobre o gráfico, dê zoom, pan e ajuste escalas em tempo real.
- 📊 **Múltiplos Modos Gráficos:**
  - **Séries Espaço-Temporais (Linhas):** Monitoramento por pontos amostrais com controle individual de séries, marcadores e estilos de traço.
  - **Comparativos de Média e Desvio Padrão (Barras Agrupadas):** Rótulos automáticos em padrão nacional (vírgula decimal) e barras customizáveis.
- 📋 **Integração Fluida com Excel:** Copie e cole blocos de células diretamente do Microsoft Excel com sanitização automática de casas decimais.
- 💾 **Gerenciador de Projetos (`.vespera` / `.json`):** Salve e retome sessões completas de trabalho (dados + estilização gráfica).
- 📥 **Exportação para Publicação Científica:**
  - **PNG (300 / 600 DPI)** para artigos e monografias.
  - **PDF e SVG Vetorial** para escalabilidade sem perda de qualidade.
  - **Planilha Excel (`.xlsx`)** com tabelas formatadas e gráficos nativos editáveis.
- 🔄 **Auto-Update Integrado via GitHub:** Atualize o software com 1 clique diretamente pela interface sempre que uma nova versão for publicada no repositório.
```
