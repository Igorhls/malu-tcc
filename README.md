# 📐 VZN EcoClima — Sistema Computacional de Análise Microclimática e Produção de Figuras Científicas

> **Trabalho de Conclusão de Curso (TCC)**  
> **Área de Concentração:** Climatologia Urbana, Conforto Térmico e Monitoramento Microclimático  
> **Repositório Web (Streamlit):** [https://github.com/Igorhls/malu-tcc.git](https://github.com/Igorhls/malu-tcc.git)  
> **Projeto Sucessor (Aplicação Desktop Nativa):** [Vespera Sci (`vespera-sci`)](https://github.com/Igorhls/vespera-sci.git)

---

## 1. Resumo e Contextualização Acadêmica

Este repositório reúne a infraestrutura computacional desenvolvida em linguagem **Python** para o processamento estatístico, tabulação normatizada e renderização gráfica de alta precisão dos dados experimentais de **Temperatura do Ar ($\text{TA}$, em $\text{°C}$)** coletados em campo para fins de Trabalho de Conclusão de Curso (TCC).

A pesquisa investiga o comportamento microclimático e o efeito de atenuação térmica proporcionado pela vegetação arbórea (**Bosque**) em comparação a superfícies antropizadas impermeáveis (**Estacionamento**). O delineamento amostral contempla duas vertentes analíticas principais:

1. **Variação Espaço-Temporal por Pontos Amostrais ($P_1$ a $P_{10}$):** Monitoramento da temperatura do ar ao longo de 10 pontos distribuídos no interior do Bosque durante os períodos **Manhã**, **Tarde** e **Noite**, realizado nas campanhas de levantamento de **14 de agosto**, **16 de agosto** e **19 de agosto**.
2. **Análise Comparativa Sazonal e Inter-Ambientes ($\text{TA}$ — Média e Desvio Padrão):** Avaliação comparativa entre os ambientes **Bosque** e **Estacionamento** nos meses de **Agosto** e **Dezembro**, expressa por meio da média aritmética ($\bar{x}$) e do desvio padrão amostral ($s$).

---

## 2. Fundamentação Estatística e Metodológica

Para cada conjunto de observações amostrais $X = \{x_1, x_2, \dots, x_n\}$, o sistema calcula automaticamente os seguintes estimadores estatísticos descritivos:

* **Média Aritmética Amostral ($\bar{x}$):**
  $$\bar{x} = \frac{1}{n} \sum_{i=1}^{n} x_i$$

* **Desvio Padrão Amostral ($s$):**
  $$s = \sqrt{\frac{1}{n - 1} \sum_{i=1}^{n} (x_i - \bar{x})^2}$$

* **Amplitude Térmica ($\Delta T$):**
  $$\Delta T = T_{\max} - T_{\min}$$

---

## 3. Base de Dados Experimentais de Referência

### 3.1. Monitoramento por Pontos Amostrais — Bosque ($\text{°C}$)

| Ponto Amostral | 14/Ago (Manhã) | 14/Ago (Tarde) | 14/Ago (Noite) | 16/Ago (Manhã) | 16/Ago (Tarde) | 16/Ago (Noite) | 19/Ago (Manhã) | 19/Ago (Tarde) | 19/Ago (Noite) |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **P1** | 26,2 | 28,5 | 25,5 | 26,0 | 28,5 | 25,8 | 24,7 | 28,2 | 27,1 |
| **P2** | 26,5 | 29,0 | 26,1 | 25,3 | 28,6 | 26,0 | 24,0 | 28,5 | 25,5 |
| **P3** | 26,5 | 29,5 | 26,2 | 25,3 | 28,9 | 26,1 | 23,4 | 28,9 | 25,3 |
| **P4** | 26,9 | 29,6 | 26,4 | 26,6 | 31,0 | 26,7 | 23,6 | 31,5 | 26,7 |
| **P5** | 27,2 | 30,4 | 26,9 | 26,2 | 30,8 | 26,8 | 23,8 | 31,0 | 26,2 |
| **P6** | 27,4 | 29,9 | 26,7 | 25,5 | 29,5 | 27,0 | 23,6 | 28,9 | 26,1 |
| **P7** | 27,1 | 30,9 | 26,8 | 25,4 | 30,4 | 27,0 | 23,5 | 31,0 | 25,8 |
| **P8** | 27,2 | 31,4 | 26,9 | 26,2 | 30,6 | 26,9 | 23,8 | 31,8 | 27,0 |
| **P9** | 27,3 | 30,3 | 27,0 | 26,0 | 30,0 | 26,6 | 24,2 | 30,0 | 27,1 |
| **P10** | 27,6 | 30,5 | 27,1 | 26,3 | 29,9 | 27,1 | 24,4 | 30,6 | 25,6 |
| **Média ($\bar{x}$)** | **26,99** | **30,00** | **26,56** | **25,88** | **29,82** | **26,60** | **23,90** | **30,04** | **26,24** |
| **Desvio ($s$)** | **0,45** | **0,88** | **0,50** | **0,46** | **0,92** | **0,46** | **0,43** | **1,34** | **0,69** |

### 3.2. Síntese Sazonal Comparativa ($\text{TA}$ — Bosque vs. Estacionamento)

| Mês de Referência | Bosque — Média ($\text{°C}$) | Bosque — Desvio/p ($\pm\text{°C}$) | Estacionamento — Média ($\text{°C}$) | Estacionamento — Desvio/p ($\pm\text{°C}$) |
| :---: | :---: | :---: | :---: | :---: |
| **Agosto** | 27,32 | 2,30 | 28,50 | 4,87 |
| **Dezembro** | 28,60 | 1,42 | 29,40 | 2,19 |

---

## 4. Padrão de Representação Gráfica Científica

As figuras produzidas pelo sistema seguem rigorosamente as diretrizes estéticas de softwares científicos de referência internacional (*OriginLab Pro* e normas editoriais da *ABNT* / periódicos *Nature* e *Science*):

* **Tipografia Padronizada:** Suporte nativo a *Times New Roman* (padrão ABNT para trabalhos acadêmicos) e *Arial* (padrão internacional para figuras científicas).
* **Moldura Fechada (*Box Frame*):** Eixos delimitados nas 4 faces (superior, inferior, esquerda e direita) com controle individual de visibilidade e espessura em pontos tipográficos ($\text{pt}$).
* **Traços de Escala (*Major & Minor Ticks*):** Marcações principais e subdivisões secundárias configuráveis (voltadas para o interior `in` nos gráficos de linhas ou para o exterior `out` nos gráficos de barras agrupadas).
* **Notação Decimal Brasileira:** Rótulos de dados formatados segundo a convenção nacional com separador de vírgula (ex.: `27,32`, `2,3`).
* **Exportação para Publicação Impressa e Digital:**
  * **PDF Vetorial (`.pdf`):** Escalabilidade infinita sem perda de definição para inclusão direta em documentos LaTeX, Microsoft Word ou Adobe InDesign.
  * **PNG de Alta Densidade (`300 DPI`):** Resolução exigida por bancas examinadoras e revistas científicas indexadas.
  * **Planilha Microsoft Excel (`.xlsx`):** Tabelas formatadas em padrão acadêmico acompanhadas de gráficos editáveis nativos gerados via `openpyxl`.

---

## 5. Arquitetura do Sistema e Funcionalidades

O repositório é composto por dois módulos operacionais independentes:

### 5.1. Aplicação Web Interativa (`app.py` — Streamlit)
Ambiente visual completo para manipulação em tempo real dos dados e parâmetros gráficos:
* **Alternância de Modos Analíticos:** Gráficos de Linhas ($P_1\text{–}P_{10}$) e Gráficos de Barras Agrupadas (Média e Desvio Padrão).
* **Importação Direta de Área de Transferência (Excel Copy-Paste):** Conversão automática de vírgulas decimais e normalização inteligente de valores colados diretamente de planilhas externas.
* **Gerenciador de Projetos Persistente:** Salvamento local em estrutura `.json` (`projetos_salvos.json`) permitindo retomar sessões de trabalho, exportar backups e carregar diferentes cenários experimentais.
* **Controle Milimétrico de Legenda e Eixos:** Posicionamento livre da legenda por coordenadas $(X, Y)$, seleção de paletas científicas ou monocromáticas (para impressão P&B) e calculadora integrada de média e desvio padrão.

### 5.2. Pipeline Autônomo em Lote (`gerar_graficos.py` — CLI)
Script de linha de comando para reprodutibilidade automatizada sem necessidade de interface gráfica. Gera de uma só vez todas as figuras individuais (`.png` e `.pdf`), o painel triplo combinado dos dias de coleta, o gráfico comparativo de barras (`TA`) e a planilha consolidada `dados_e_graficos_bosque.xlsx`.

---

## 6. Estrutura de Diretórios

```text
Malu-TCC/
├── app.py                          # Aplicação web interativa (Streamlit)
├── gerar_graficos.py               # Script autônomo de geração em lote (CLI)
├── dados_e_graficos_bosque.xlsx    # Planilha Excel consolidada com gráficos nativos
├── graficos_gerados/               # Diretório de saída das figuras de publicação (300 DPI e PDF)
│   ├── grafico_dia_14.png / .pdf   # Perfil térmico espacial — 14 de agosto
│   ├── grafico_dia_16.png / .pdf   # Perfil térmico espacial — 16 de agosto
│   ├── grafico_dia_19.png / .pdf   # Perfil térmico espacial — 19 de agosto
│   ├── graficos_todos_dias.png/.pdf# Painel comparativo integrado (3 dias)
│   └── grafico_media_desvio_ta.*   # Gráfico de barras agrupadas (Média e Desvio Padrão - TA)
├── requirements.txt                # Especificação das bibliotecas científicas Python
├── VESPERA_SCI_AGENT_CONTEXT.md    # Documento de arquitetura e transição para o Vespera Sci Desktop
└── README.md                       # Documentação acadêmica e técnica do projeto
```

---

## 7. Instruções de Execução e Reprodutibilidade

### 7.1. Pré-requisitos e Instalação
Certifique-se de possuir o **Python 3.10+** instalado. No terminal, instale as dependências listadas em `requirements.txt`:

```bash
py -m pip install -r requirements.txt
```

### 7.2. Executando o Painel Interativo (Streamlit)
Para abrir a interface visual no navegador:

```bash
py -m streamlit run app.py
```

### 7.3. Executando a Geração Automatizada em Lote (Terminal)
Para regerar instantaneamente todos os arquivos `.png`, `.pdf` e `.xlsx` de referência do TCC:

```bash
py gerar_graficos.py
```

---

## 8. Tecnologias e Bibliotecas Utilizadas

* **[Python 3](https://www.python.org/):** Linguagem base de processamento científico.
* **[Matplotlib](https://matplotlib.org/):** Motor de renderização vetorial e matricial com precisão tipográfica.
* **[Pandas](https://pandas.pydata.org/) & [NumPy](https://numpy.org/):** Estruturação matricial e cálculos estatísticos descritivos.
* **[OpenPyXL](https://openpyxl.readthedocs.io/):** Construção programática de planilhas Excel (`.xlsx`) e gráficos nativos.
* **[Streamlit](https://streamlit.io/):** Interface web reativa para exploração e customização dos experimentos.

---

> **Nota de Evolução do Projeto:**  
> Como desdobramento tecnológico deste TCC, o motor gráfico aqui desenvolvido deu origem ao projeto **[Vespera Sci](https://github.com/Igorhls/vespera-sci.git)**, uma plataforma desktop científica modular com suporte a múltiplos tipos de gráficos e atualização automatizada via GitHub.
