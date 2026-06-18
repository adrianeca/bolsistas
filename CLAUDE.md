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
- `ROLES` — permissões por role (coluna `bolsistas` = TRUE/FALSE)
- `USUARIOS` — cadastro de usuários (coluna `extraDashboards` para acesso individual)

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

## Melhorias planejadas (não implementadas)

### Performance — carregamento lento
O carregamento está lento por 3 leituras sequenciais de planilha na validação + 2 roundtrips ao servidor.

**Plano:**
1. **Fundir chamadas iniciais** — criar `initApp(token)` que valida sessão E retorna opções de filtro em um único roundtrip (elimina uma chamada ao servidor)
2. **CacheService** — cachear objeto do usuário por ~10 min após primeira validação; chamadas seguintes leem do cache em vez de abrir planilha
3. **Simplificar `_hasAccess`** — usar coluna `dashboards` (ou `extraDashboards`) na aba USUARIOS, buscando pelo e-mail; elimina leitura da aba ROLES

> **Pendente:** confirmar nome exato da coluna na aba USUARIOS antes de implementar o item 3.
