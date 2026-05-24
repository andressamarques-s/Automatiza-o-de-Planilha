# Automatiza-o-de-Planilha

Automação desenvolvida em **JavaScript (Google Apps Script)** para processar e organizar respostas de formulários de pesquisa de satisfação, utilizada em ambiente real de trabalho.

##  Contexto

Sistema desenvolvido para o **Centro Médico Adventista**, com o objetivo de automatizar o tratamento de dados de uma pesquisa de lista de presença e satisfação aplicada em múltiplas áreas de serviço.

##  Funcionalidades

- Processamento automático de novas respostas via trigger `onFormSubmit`
- Limpeza e padronização dos dados brutos
- Organização em aba estruturada com cabeçalho formatado
- Cálculo de **NPS** (Net Promoter Score) por respondente
- Identificação automática da área de serviço (SST, AMA, GPS, SPS)
- Formatação de CPF, nomes e datas
- Reprocessamento completo do histórico quando necessário

##  Estrutura dos Dados

| Campo | Descrição |
|-------|-----------|
| `nome` | Nome capitalizado automaticamente |
| `cpf` | CPF formatado (000.000.000-00) |
| `area_servico` | Identificada automaticamente pelo formulário |
| `nps_recomendacao` | Nota de 0 a 10 |
| `nps_categoria` | Promotor / Neutro / Detrator |
| `resolvido` | Se o problema foi solucionado |

## Conceitos Aplicados

- Google Apps Script (JavaScript)
- Triggers automáticos por envio de formulário
- Manipulação de planilhas com a API do Google Sheets
- Tratamento e limpeza de dados
- Lógica condicional para múltiplas áreas de serviço

## Como usar

1. Abra o Google Sheets vinculado ao formulário
2. Vá em **Extensões → Apps Script**
3. Cole o código e salve
4. Execute `configurarTrigger()` uma vez para ativar a automação
5. A partir daí, cada nova resposta é processada automaticamente

## 👩‍💻 Autora

Desenvolvido como solução real de automação no ambiente de trabalho.
