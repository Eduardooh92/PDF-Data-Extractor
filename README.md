# PDF Data Extractor & Automated Form Filler

## 1. Visão Geral

Este projeto é um sistema de automação ETL (Extract, Transform, Load) desenvolvido em Python. Sua função principal é monitorar uma pasta de entrada, extrair informações cadastrais de múltiplos arquivos PDF (como Cartão CNPJ e Sintegra/Inscrição Estadual), consolidar esses dados, aplicar regras de negócio e, por fim, preencher uma planilha Excel (`.xlsx`) pré-formatada.

O sistema foi projetado para ser robusto e desacoplado, utilizando um arquivo de configuração externo (`config.ini`) para gerenciar todos os caminhos e parâmetros, além de manter um log detalhado de todas as operações.

## 2. Funcionalidades Principais

- **Extração de Dados de PDF:** Utiliza a biblioteca `PyMuPDF` para ler e extrair texto de documentos PDF.
- **Análise com Regex:** Emprega expressões regulares para identificar e isolar dados específicos como CNPJ, Razão Social, Endereço, etc.
- **Configuração Centralizada:** Todos os caminhos de pastas (entrada, saída, processados, erros) e arquivos são gerenciados externamente pelo `config.ini`, permitindo fácil adaptação do ambiente sem alterar o código-fonte.
- **Gerenciamento de Arquivos:** O script move automaticamente os arquivos PDF da pasta de entrada para pastas de "processados" ou "erros", garantindo que um arquivo seja processado apenas uma vez.
- **Logging Robusto:** Gera um arquivo de log (`processamento.log`) com rotação de tamanho, registrando cada passo da execução, de sucessos a avisos e erros críticos, facilitando o diagnóstico de problemas.
- **Preenchimento de Excel:** Utiliza a biblioteca `openpyxl` para inserir os dados extraídos em células específicas de um template Excel, preservando a formatação original.

## 3. Tecnologias Utilizadas

- **Linguagem:** Python 3.x
- **Bibliotecas Principais:**
  - `PyMuPDF (fitz)`: Leitura e extração de texto de arquivos PDF.
  - `openpyxl`: Manipulação de planilhas Excel (`.xlsx`).
  - `configparser`: Gerenciamento do arquivo de configuração `.ini`.
  - `logging`: Sistema de registro de eventos.
  - `re`: Operações com expressões regulares.

## 4. Como Executar o Sistema

### Pré-requisitos

- Python 3.8 ou superior instalado.
- Acesso ao terminal ou prompt de comando.
- Um template de planilha Excel (`.xlsx`) com a estrutura esperada.

### Instalação

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/nome-do-repositorio.git](https://github.com/seu-usuario/nome-do-repositorio.git)
    cd nome-do-repositorio
    ```

2.  **Instale as dependências:**
    É altamente recomendável usar um ambiente virtual (`venv`) para isolar as dependências do projeto.
    ```bash
    # Crie um ambiente virtual (opcional, mas recomendado)
    python -m venv venv
    
    # Ative o ambiente virtual
    # Windows
    .\venv\Scripts\activate
    # Linux / macOS
    source venv/bin/activate
    
    # Instale as bibliotecas necessárias
    pip install PyMuPDF openpyxl
    ```

### Configuração

1.  **Crie as pastas:** Crie a estrutura de pastas que será usada pelo sistema (entrada, saída, processados, erros).

2.  **Edite o `config.ini`:**
    Abra o arquivo `config.ini` e ajuste os caminhos (`Paths`) para que correspondam à estrutura de pastas que você criou no seu ambiente.

    ```ini
    [Paths]
    InputFolder = /caminho/para/sua/pasta/de/entrada
    OutputFolder = /caminho/para/sua/pasta/de/saida
    ProcessedFolder = /caminho/para/sua/pasta/de/processados
    ErrorFolder = /caminho/para/sua/pasta/de/erros
    ExcelTemplate = /caminho/para/seu/template.xlsx

    [Settings]
    LogFile = processamento.log
    ```

### Execução

1.  **Adicione os arquivos PDF** que você deseja processar na pasta definida em `InputFolder`.

2.  **Execute o script** a partir do seu terminal:
    ```bash
    python processador_ficha.py
    ```

3.  O script irá processar os arquivos, gerar a planilha preenchida na pasta de `OutputFolder` e mover os PDFs para as pastas `ProcessedFolder` ou `ErrorFolder`, conforme o resultado da operação. O arquivo `processamento.log` será criado no mesmo diretório do script com todos os detalhes da execução.
