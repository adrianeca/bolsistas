# BRASAS Bolsistas

Google Apps Script — painel de gestão de alunos bolsistas.

## Arquivos

- `Code.gs` — backend (Apps Script): autenticação, leitura/edição de dados, geração de PDF, envio de e-mail
- `Index.html` — frontend: interface completa (HTML + CSS + JS inline)

## Planilhas

| Constante | ID | Uso |
|---|---|---|
| `USERS_SHEET_ID` | `1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc` | Sessões, usuários, roles |
| `BOLSISTAS_SHEET_ID` | `1G7RfHP_8j7-6VPqC8ReYvpScyWvBnVyRXu8QAdeL4bQ` | Dados dos bolsistas |

**Abas em USERS_SHEET:**
- `SESSOES` — tokens de sessão com expiração
- `ROLES` — permissões por role (coluna `bolsistas` = TRUE/FALSE) — não usado em `_hasAccess` atualmente
- `USUARIOS` — cadastro de usuários; colunas: A e-mail, B nome, C role, D unidade (restrição de acesso), E ativo, F e G = `acessos_dashboards`. `bolsistas` na lista de acessos é só o token de entrada no app (como `mat` para matrículas) — não concede acesso a nenhuma página sozinho; cada página exige seu próprio token (`bolsistas_alunos`, `bolsistas_relatorios`, `bolsistas_historico`, `bolsistas_motivo` — singular, `bolsistas_liberacao`)

**Abas em BOLSISTAS_SHEET:**
- `Bolsistas App` — dados principais lidos/editados pelo painel
- `bolsistas_consolidado` — fonte do ciclo mensal
- `db_bolsistas_mes_anterior` — backup do mês anterior
- `Parceiros` — e-mails e logos dos parceiros de bolsa
- `Feriados` — feriados para cálculo de dias úteis

## Hub

URL do hub (sistema de autenticação central):
```
https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec
```

O painel recebe o token de sessão via parâmetro `?s=TOKEN` na URL.

## Roles e permissões

```javascript
const EDIT_ROLES  = ['admin', 'secretaria', 'diretor', 'b2b'];
const EMAIL_ROLES = ['admin', 'b2b'];
const ADMIN_ROLES = ['admin'];
```

## Arquitetura do frontend (Index.html)

### Boot em 2 fases
1. `validateSession(token)` — só valida sessão (rápido, lê SESSOES). Mostra o app.
2. `initApp(token)` — carrega todos os dados em background. Popula a tabela.

### Globals principais
- `FULL_ROWS` — todos os dados carregados (respeitando restrição de unidade)
- `ALL_ROWS` — subconjunto após filtros do cliente
- `USER` — objeto com `{ email, nome, role, unidade, canEdit, canSendEmail, isAdmin }`
- `_currentPDFFiles` — `{ [chavePDF]: { fileId, viewUrl } }` — PDFs gerados na sessão

### Filtros
Todos client-side via `applyFilters()` sobre `FULL_ROWS`. Multi-select: Ano, Mês, Unidade, Origem. Sidebar de anos (AppSheet-style). Grupos por Mês na tabela (apenas quando sem ordenação).

### Campos editáveis na tabela / painel de detalhe
`data1aAula`, `frequencia`, `dataInicioBook`, `horario`, `diasPrevistos`, `diasAssistidos`, `mt`, `wt`, `oc`, `ot`, `aproveitamento`, `observacoes`

### Campos calculados (client-side, não persistidos)
- `_nota` — `IF(WT≠"", (WT+OC+OT)/3, MT)`
- `_aprovAjustado` — E/MB/B/R/I/N baseado em `_nota` e `diasAssistidos`
- `_faltados` — dias faltados acumulados por livro (book-level)

### Fluxo PDF / E-mail
1. Abre modal → anos/meses carregados de `FULL_ROWS` (client-side, sem servidor)
2. Seleciona Ano + Mês → lista de empresas (groupadas por `r.chavePDF`) aparece instantaneamente
3. "📄 Gerar PDF" → chama `gerarPDF(token, chavePDF)` → armazena `{ fileId, viewUrl }` em `_currentPDFFiles`
4. "👁 Ver PDF" aparece → abre Drive em nova aba para preview
5. "📧 Enviar" → modal de e-mail mostra link do PDF + destinatários + campo de obs
6. "Enviar E-mail" → chama `enviarEmail` com `fileId` do PDF já gerado

**Nota:** cada empresa é enviada individualmente. Não há "enviar tudo".

## Backend (Code.gs) — funções principais

| Função | Descrição |
|---|---|
| `validateSession(token)` | Valida token na aba SESSOES; retorna objeto do usuário |
| `initApp(token)` | Valida sessão + retorna rows + opções de filtro (CacheService 300s) |
| `updateBolsista(token, payload)` | Edita campo de um bolsista; invalida cache de dados |
| `gerarPDF(token, chavePDF)` | Gera Google Doc → PDF → Drive; retorna `{ fileId, viewUrl }` |
| `getEmailPreview(token, chavePDF)` | Retorna info do parceiro + destinatários para preview no modal |
| `enviarEmail(token, payloadJson)` | Envia e-mail com PDF anexado para o parceiro |
| `getChavesPDF(token, mes, ano)` | Lista chaves PDF disponíveis (ainda existe, não mais usado pelo frontend) |
| `runCicloMensal(token)` | Executa ciclo mensal (backup + consolidado → Bolsistas App) |

### CacheService
- `user_TOKEN` → objeto do usuário (600s) — invalidado automaticamente ao expirar
- `appdata_TOKEN` → resultado de `initApp` serializado (300s) — invalidado por `updateBolsista`

### Restrição de unidade
`_getUserUnidade(ss, email)` lê coluna D da aba USUARIOS. Se não-vazio, o usuário só vê linhas cuja `unidade` bate com a lista (suporta múltiplas unidades separadas por vírgula).

## Implementado (histórico de features)

- ~~Simplificar `_hasAccess`~~ — lê apenas aba USUARIOS, elimina leitura de ROLES
- ~~Performance: 2 fases de boot~~ — `validateSession` rápido + `initApp` em background
- ~~CacheService~~ — cache de usuário (600s) e dados (300s)
- ~~Restrição por unidade~~ — filtra linhas pelo campo UNIDADE da aba USUARIOS
- ~~PDF preview antes de enviar~~ — botão "Ver PDF" por empresa + link no modal de e-mail
- ~~Envio por empresa individual~~ — cada empresa tem seus próprios botões independentes

## Line endings

`Index.html` tem line endings mistos: CRLF no JS regular, LF dentro de template literals. O Edit tool do Claude Code falha em strings que cruzam essa fronteira. Usar PowerShell com `[System.IO.File]::ReadAllBytes` + `String.Replace` para edições nessas áreas.
