# 🌿 VZN EcoClima — Monitoramento Microclimático (TCC)

Aplicação e scripts em Python para visualização, análise e exportação de dados de temperatura do ar por pontos amostrais (**P1 a P10**) para os períodos **Manhã**, **Tarde** e **Noite** nos dias:
- **14 de agosto**
- **16 de agosto**
- **19 de agosto**

Repositório oficial: [https://github.com/Igorhls/malu-tcc.git](https://github.com/Igorhls/malu-tcc.git)

---

## 🌐 Aplicação Web Interativa (Streamlit)

O projeto conta com um painel web completo e responsivo construído com **Streamlit**.

### Como rodar localmente:
```bash
py -m streamlit run app.py
```
*(ou simplesmente `streamlit run app.py`)*

### Recursos do aplicativo:
- 💾 **Salvar e Retomar Projetos**: Botão de salvamento rápido que guarda todos os dados da tabela e personalizações do gráfico (cores, fontes, bordas, legenda e escalas) para continuar o trabalho posteriormente, além de backup/importação em `.json`.
- 📅 **Seleção de Dias**: Filtro por dia específico ou painel geral integrado.
- 🕒 **Filtros de Período**: Seleção individual ou combinada de Manhã, Tarde e Noite.
- 📈 **Gráficos em Alta Resolução**: Renderização vetorial dos gráficos idênticos ao modelo.
- 📊 **Estatísticas Descritivas**: Temperatura média, máxima, mínima e amplitude térmica.
- 📥 **Exportação Rápida**: Download direto das figuras em PNG (300 DPI), PDF vetorial e planilha Excel (.xlsx).

---

## 📁 Arquivos e Estrutura do Projeto

- **`app.py`**: Aplicação web Streamlit interativa.
- **`gerar_graficos.py`**: Script Python autônomo para geração das figuras e planilha.
- **`dados_e_graficos_bosque.xlsx`**: Planilha Excel com as tabelas e gráficos embutidos nativamente.
- **`graficos_gerados/`**:
  - `grafico_dia_14.png` e `.pdf`
  - `grafico_dia_16.png` e `.pdf`
  - `grafico_dia_19.png` e `.pdf`
  - `graficos_todos_dias.png` e `.pdf` (Painel com todos os dias)
- **`requirements.txt`**: Dependências do projeto (compatível com deploy no Streamlit Cloud).

---

## 🚀 Como gerar os gráficos via terminal

Se preferir rodar apenas o gerador de arquivos sem abrir o navegador:
```bash
py gerar_graficos.py
```

---

## 📦 Instalação das Dependências

```bash
py -m pip install -r requirements.txt
```
