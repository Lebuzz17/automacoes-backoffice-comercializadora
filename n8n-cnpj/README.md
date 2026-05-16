# N8N — Consulta de CNPJ e Estruturação de Endereços

Workflow N8N que consulta dados públicos de CNPJ e retorna endereço estruturado.

## Como importar

1. N8N → Workflows → Import from File
2. Selecione `workflow.json`
3. Configure credenciais (token de `api.cnpja.com` se for usar a fonte primária)

## APIs usadas

| API | Tipo | Documentação |
|-----|------|--------------|
| `api.cnpja.com` | Privada (free tier) | https://docs.cnpja.com |
| `brasilapi.com.br/api/cnpj/v1/` | Pública, gratuita | https://brasilapi.com.br |

## Como funciona

1. **Input**: lista de CNPJs (pode vir de qualquer trigger — manual, webhook, planilha).
2. **Loop por CNPJ**: consulta a API com tratamento de erros.
3. **Output**: razão social, endereço completo, CEP, atividade principal, situação cadastral.

## Decisões de design

- **Cache opcional**: para CNPJs já consultados na mesma sessão, retorna do cache em vez de bater na API.
- **Fallback entre fontes**: se a API primária falha (rate limit, instabilidade), tenta a secundária.
- **Saída estruturada**: campos prontos para uso (não retorna o JSON cru da API).

## Limitações

- API `cnpja.com` no plano free tem rate limit de poucas requisições por minuto. Para volumes maiores, use o plano pago ou a BrasilAPI como fonte principal.
- Não faz validação prévia de CNPJ (módulo 11). Se o CNPJ é inválido, a API retorna erro.
