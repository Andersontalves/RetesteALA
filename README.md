# 📡 Analisador de Sinal FTTH

Ferramenta desktop (Python + Tkinter) para cruzar uma base de clientes FTTH com dados exportados da OLTCloud.

Normaliza sinais em milli-dBm, classifica a qualidade do sinal e gera duas abas formatadas diretamente na planilha Excel de entrada.

---

## ✨ Funcionalidades

- Detecta automaticamente o arquivo `.xlsx` na pasta
- Cruza a **1ª aba (Clientes)** com a **2ª aba (OLTCloud)** pelo número de contrato
- Converte sinais em milli-dBm para dBm automaticamente
- Classifica o sinal como: **BOM**, **RUIM**, **SINAL ALTO** ou **SEM DADOS**
- Gera a aba **RESULTADO** — todos os clientes com status de sinal
- Gera a aba **SOMENTE_BONS** — apenas clientes com sinal dentro do critério
- Interface gráfica (GUI) com log em tempo real e seleção de arquivo via janela

---

## 📋 Pré-requisitos da Planilha

O arquivo `.xlsx` deve ter **pelo menos 2 abas**:

| Posição | Conteúdo | Colunas obrigatórias |
|---|---|---|
| 1ª aba | Base de clientes | `contrato` |
| 2ª aba | Exportação da OLTCloud | `External Contract ID`, `RX ONU`, `RX OLT`, `Status`, `OLT`, `SN ONU`, `Modelo` |

---

## 🚀 Como Usar

### Modo GUI (recomendado)

```bash
python app_gui.py
```

1. Clique em **Buscar Arquivo** e selecione a planilha `.xlsx`
2. Clique em **▶ INICIAR PROCESSAMENTO**
3. Escolha onde salvar o arquivo de resultado
4. Aguarde a conclusão — o log mostrará o progresso em tempo real

### Modo linha de comando

Coloque o script na mesma pasta do arquivo `.xlsx` e execute:

```bash
python processar.py
```

---

## 📦 Instalação

```bash
pip install -r requirements.txt
```

---

## 🏗️ Gerar Executável (.exe)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=alares.ico --add-data "alares.ico;." app_gui.py
```

O executável será gerado em `dist/app_gui.exe`.

---

## 📂 Estrutura do Projeto

```
.
├── app_gui.py        # Aplicação com interface gráfica
├── processar.py      # Lógica de processamento (também usável via CLI)
├── alares.ico        # Ícone da aplicação
├── requirements.txt  # Dependências Python
└── README.md
```

---

## 📊 Critérios de Classificação de Sinal

| Canal | BOM | SINAL ALTO | RUIM |
|---|---|---|---|
| RX ONU | -24,99 a -10 dBm | acima de -10 dBm | abaixo de -24,99 dBm |
| RX OLT | -26,99 a -10 dBm | acima de -10 dBm | abaixo de -26,99 dBm |

> Valores em milli-dBm são **convertidos automaticamente** para dBm antes da classificação.

---

## 🛠️ Tecnologias

- Python 3.x
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- Tkinter (incluso no Python padrão)
