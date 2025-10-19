# Otimizador de Horários Escolares (Co-criado com IA)

## ⚡ Impacto: Processo manual de 5 horas reduzido para 10 minutos ⚡

Este é um script de Automação de Processos (RPA) e Engenharia de Dados (ETL) que resolve um problema de negócios do mundo real: a alocação manual de horários escolares.

O script foi desenvolvido para automatizar a tarefa de transcrever horários de um PDF complexo e não estruturado (gerado por um sistema legado) para uma planilha Excel limpa, organizada e visualmente formatada.

**Nota de Desenvolvimento:** Este projeto foi arquitetado, desenvolvido e depurado usando Engenharia de Prompts Avançada, com o Google Gemini servindo como um parceiro de co-criação de código.

---

## 🚀 Funcionalidades Principais (O Processo ETL)

1.  **(Extract) Extração de Dados:**
    * Utiliza a biblioteca `PDFPlumber` para ler tabelas complexas distribuídas por várias páginas de um arquivo PDF.
    * Mapeia com precisão as salas (que estão fora das tabelas) para os dados corretos.

2.  **(Transform) Transformação e Limpeza:**
    * Utiliza `Pandas` para limpar, normalizar e estruturar os dados brutos extraídos.
    * Expande aulas geminadas (duplas/triplas) em entradas individuais.
    * Normaliza nomes de professores (removendo acentos, corrigindo erros comuns) e disciplinas.
    * Reorganiza os dados em uma matriz (pivot table) de **Dia/Hora vs. Local**.

3.  **(Load) Carregamento e Formatação:**
    * Utiliza `Openpyxl` para salvar os dados em uma planilha Excel (`.xlsx`).
    * Aplica formatação condicional avançada: **colore automaticamente** cada aula com base no professor (cores quentes para professoras, frias para professores), facilitando a visualização.
    * Ajusta o alinhamento, fontes e largura das colunas para criar um relatório final pronto para distribuição.

---

## 🛠️ Stack de Tecnologia

* **Python**
* **Pandas:** Para toda a transformação e pivoteamento dos dados.
* **PDFPlumber:** Para a extração de dados de tabelas do PDF.
* **Openpyxl:** Para a formatação avançada (cores, fontes, bordas) da planilha Excel final.
* **Engenharia de Prompts (Gemini):** Para co-criação de código, lógica de ETL e depuração.

---

## ⚙️ Como Executar

1.  Coloque o arquivo PDF-fonte na mesma pasta do script e renomeie-o para `horarios.pdf`.
2.  Instale as dependências:
    `pip install pandas pdfplumber openpyxl`
3.  Execute o script:
    `python seu_script.py` (ou o nome do seu arquivo principal)
4.  Um novo arquivo, `Horario_ATUALIZADO.xlsx`, será criado na pasta.
