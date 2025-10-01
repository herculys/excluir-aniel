# Excluir Aniel

Script Python para processar dados entre sistemas Voalle e Aniel, gerando relatórios de exclusão filtrados.

## Como Usar

1. Coloque os arquivos na pasta `Excluir Aniel/`:
   - `QTD Solicitações _ Recolhimento.csv`
   - `RECOLHIMENTO.xlsx`
   - `RECOLHIMENTO AGENDADO.xlsx`
   - `RECOLHIMENTO RÁDIO.xlsx`

2. Execute:
   ```powershell
   .\executar.ps1
   ```

## Saída

O script cria uma pasta `Exclusão XX - DD-MM-AAAA` com:
- `Excluir_Aniel.xlsx` - Dados principais
- `Excluir_Aniel_com_Status.xlsx` - Com status detalhado
- `Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx` - Casos especiais
- `Source/` - Arquivos originais movidos

## Filtros

Inclui apenas registros com Status Voalle:
- Cancelado
- Fechamento  
- Encerramento

## Requisitos

- Python 3.7+
- Ambiente virtual já configurado
- Dependências: pandas, tqdm, openpyxl, tabulate