# Automação Integrada — SAP & Cargo Heroes

Automação para criação massiva de Requisições de Compra no SAP (transação ME51N) e atualização de registros logísticos no Cargo Heroes, com dados lidos de uma planilha Google Sheets.

## Requisitos

- Python 3.10 ou superior
- SAP GUI com scripting habilitado e acesso à transação ME51N
- Google Chrome + ChromeDriver no PATH (para o fluxo Cargo Heroes)
- Arquivo de credenciais do Google Sheets: `credentials.json` (service account) na mesma pasta do script
- Conta com acesso ao [Cargo Heroes](https://cargo-heroes.appslatam.com)

## Instalação

```bash
pip install -r requirements.txt
```

## Como Usar

Execute o arquivo principal:

```bash
python APP_UNIFICADO_SAP_CH.py
```

### Aba Configurações

Preencha e salve antes de iniciar qualquer automação:

| Campo | Descrição |
|---|---|
| Caminho Logon.exe | Caminho para o executável do SAP Logon |
| Sistema | Identificador do sistema SAP (ex: `PRD`) |
| Usuário / Senha | Credenciais SAP (senha salva no Windows Credential Manager) |
| Credenciais | Caminho para o arquivo `credentials.json` |
| Planilha | Nome da planilha no Google Sheets |
| Aba | Nome da aba a ser processada |
| Email / Senha CH | Credenciais para login no Cargo Heroes |

### Aba Automação

- **Iniciar SAP** — conecta ao SAP GUI, lê as linhas com `Status` em branco e cria as Requisições de Compra (ME51N), gravando o número da RC e o status de volta na planilha.
- **Atualizar CH** — abre o Chrome, faz login SSO no Cargo Heroes e atualiza os registros logísticos via API interna, marcando `CH OK = OK` ao concluir.
- **Parar Automação** — interrompe o ciclo atual de forma segura ao fim do item em processamento.

## Estrutura da Planilha Google Sheets

### Aba principal (SAP + CH normal)

| Coluna | Descrição |
|---|---|
| PN / Material ID | Número do material |
| QTD / Quantidade | Quantidade |
| ORIGEM / Origem Sigla | Centro/base de origem |
| DESTINO / Destino Sigla | Centro/base de destino |
| DATA REMESSA | Data de remessa (DD/MM/AAAA) |
| TEXTO / Logística | Descrição e horários de embarque/pouso |
| Tipo de Transporte | `Aéreo` ou `Terrestre` |
| Status | Deixar em branco para itens novos; preenchido automaticamente |
| REQUISIÇÃO | Número da RC criada (preenchido automaticamente) |
| CH OK | Preenchido com `OK` ou `ERRO` pelo fluxo Cargo Heroes |

### Aba `MAPEAMENTO` (opcional)

Usada para ajustar o status de equipamentos no Cargo Heroes. Colunas relevantes:

| Coluna | Valores esperados |
|---|---|
| Material ID | Código do equipamento |
| ORIGEM | `NA BASE` ou `ZERO` |
| CH OK | Deixar em branco para processar |
