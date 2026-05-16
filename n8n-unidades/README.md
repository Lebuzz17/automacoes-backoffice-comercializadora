# N8N — Criação de Unidades Consumidoras

Workflow N8N que cadastra unidades em massa no ERP interno via API REST.

## Como importar

1. N8N → Workflows → Import from File
2. Selecione `workflow.json`
3. Configure as credenciais HTTP que aparecerão pendentes

## Endpoints usados (sanitizados)

| Endpoint | Método | Função |
|----------|--------|--------|
| `https://api-erp-interno.exemplo.com/authenticate` | POST | Login e geração de token JWT |
| `https://api-erp-interno.exemplo.com/rest/dadosSistema/cityByState` | GET | Resolve `cidade + estado → cityId` |
| `https://api-erp-interno.exemplo.com/rest/unit` | POST | Cria a unidade |

⚠️ Substitua `api-erp-interno.exemplo.com` pela URL da API do seu ERP.

## Estrutura do workflow

```
Form Trigger (upload planilha)
    ↓
Split into Items (1 item por linha)
    ↓
Authenticate (POST /authenticate)
    ↓
Set credenciais no contexto
    ↓
Loop:
    ├─ Lookup cityId (GET /cityByState)
    ├─ Build payload da unidade
    └─ POST /rest/unit
    ↓
Aggregate resultados
    ↓
Output (resumo)
```

## Decisões de design

- **Auth uma vez fora do loop**: token JWT serve para o loop inteiro, evita re-autenticar para cada unidade.
- **Lookup de city por estado**: a API do ERP recebe `cityId` (inteiro), não nome de cidade. O lookup `cityByState` cacheado na memória do workflow resolve isso.
- **Continue On Fail nos nodes POST**: uma unidade que falha (CNPJ duplicado, e.g.) não para o lote — gera erro registrado e segue.

## Limitações

- Não trata rate limit da API explicitamente (depende dos limites do seu ERP).
- Não faz idempotência (se rodar 2x a mesma planilha, cria duplicatas, a menos que a API rejeite).
