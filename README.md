# Relatório Excluir Aniel

Este projeto contém um script Python para processar e mesclar dados de cancelamentos do sistema Voalle com dados do painel de serviços do sistema Aniel, gerando relatórios filtrados para exclusão.

## � Ambiente Virtual

**IMPORTANTE:** Este projeto agora está configurado com um ambiente virtual Python para isolar as dependências.

### Dependências Instaladas:
- pandas - Manipulação de dados
- tqdm - Barra de progresso
- openpyxl - Leitura/escrita de arquivos Excel
- tabulate - Formatação de tabelas
- numpy - Operações numéricas

### Como Executar:

#### Método 1: Script Automático (PowerShell)
```powershell
.\executar.ps1
```

#### Método 2: Script Batch (CMD)
```cmd
ativar_venv.bat
python Exclusao.py
```

#### Método 3: Manual
```powershell
.\venv\Scripts\Activate.ps1
python Exclusao.py
```

## 📋 Descrição

O script `Exclusao.py` automatiza o processo de:
- Leitura de dados de cancelamentos do Voalle (CSV)
- Leitura de dados do painel de serviços do Aniel (Excel)
- Mesclagem dos dados baseada no número do protocolo/ordem de serviço
- Filtragem de registros conforme critérios específicos
- Geração de relatórios em formato Excel

## 🔧 Pré-requisitos

- Python 3.7 ou superior
- Os seguintes pacotes Python:
  - `pandas`
  - `tqdm`
  - `tabulate`
  - `openpyxl`

## 📦 Instalação

1. Clone ou baixe este repositório
2. Instale as dependências necessárias:

```bash
pip install pandas tqdm tabulate openpyxl
```

## 📁 Estrutura de Arquivos

```
excluir-aniel/
├── RelatorioExcluirAniel.py
├── Excluir Aniel/
│   ├── Canceladas Voalle.csv
│   ├── Painel de Serviços.xlsx
│   └── Exclusão/
│       ├── Excluir_Aniel.xlsx
│       ├── Excluir_Aniel_com_Status.xlsx
│       └── Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx
└── README.md
```

## 📄 Arquivos de Entrada

### Canceladas Voalle.csv
- Arquivo CSV com dados de protocolos cancelados no sistema Voalle
- Formato: separado por ponto e vírgula (;)
- Codificação: UTF-8 ou Latin-1
- Campos principais: `Protocolo`, `Status`

### Painel de Serviços.xlsx
- Arquivo Excel com dados do painel de serviços do Aniel
- Formato: Excel (.xlsx)
- Header na linha 2
- Campos principais: `Nº. Ordem Serviço`, `Contrato`, `Projeto`, `Identificação do Cliente`, `Data/Hora Criação`, `Status`

## 🚀 Como Usar

1. Certifique-se de que os arquivos de entrada estão na pasta `Excluir Aniel/`:
   - `Canceladas Voalle.csv`
   - `Painel de Serviços.xlsx`

2. Execute o script:
```bash
python RelatorioExcluirAniel.py
```

3. O script mostrará uma barra de progresso com as seguintes etapas:
   - Lendo CSV Canceladas Voalle.csv
   - Lendo Excel Painel de Serviços.xlsx
   - Mesclando Planilhas
   - Montando DataFrame final
   - Filtrando Status
   - Salvando arquivos

## 📊 Arquivos de Saída

O script gera três arquivos na pasta `Excluir Aniel/Exclusão/`:

### 1. Excluir_Aniel.xlsx
- Contém os registros filtrados **sem** as colunas de status
- Campos: Contrato, Projeto, Identificação do Cliente, Ordem de Serviço, Data de Criação

### 2. Excluir_Aniel_com_Status.xlsx
- Contém os registros filtrados **com** as colunas de status
- Campos: Contrato, Projeto, Identificação do Cliente, Ordem de Serviço, Data de Criação, Status Voalle, Status Aniel

### 3. Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx
- Contém apenas registros onde:
  - Status Voalle = "Cancelado"
  - Status Aniel = "Fechada Produtiva"

## 🔍 Critérios de Filtragem

O script aplica os seguintes filtros:

1. **Mesclagem**: Junta dados do Voalle e Aniel baseado no campo Protocolo/Ordem de Serviço
2. **Exclusão de Status**: Remove registros com Status Aniel igual a:
   - "Fechada Improdutiva"
   - "Fechada Produtiva"
   - "Cancelado"
3. **Arquivo Especial**: Cria arquivo separado para registros com Status Voalle "Cancelado" E Status Aniel "Fechada Produtiva"

## 📅 Formatação de Data

As datas são automaticamente formatadas para o padrão brasileiro: `DD/MM/AAAA HH:MM:SS`

## ⚠️ Tratamento de Erros

O script inclui tratamento para:
- Arquivos não encontrados
- Problemas de codificação (tenta UTF-8, depois Latin-1)
- Linhas mal formatadas no CSV

## 🖥️ Interface

- Barra de progresso visual durante a execução
- Tabela formatada mostrando preview dos resultados
- Mensagens de confirmação ao final
- Indicação dos arquivos gerados

## 📝 Notas

- O script cria automaticamente a pasta `Excluir Aniel/Exclusão/` se ela não existir
- Todos os arquivos de saída são sobrescritos a cada execução
- O script para a execução se os arquivos de entrada não forem encontrados

## 🐛 Solução de Problemas

### Arquivo não encontrado
Verifique se os arquivos estão na pasta correta:
- `Excluir Aniel/Canceladas Voalle.csv`
- `Excluir Aniel/Painel de Serviços.xlsx`

### Erro de codificação
O script tenta automaticamente UTF-8 e Latin-1. Se ainda houver problemas, verifique a codificação do arquivo CSV.

### Dependências não instaladas
Execute: `pip install pandas tqdm tabulate openpyxl`

## 👥 Contribuições

Para contribuir com este projeto:
1. Faça um fork do repositório
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Faça um push para a branch
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob licença MIT. Veja o arquivo LICENSE para mais detalhes.
