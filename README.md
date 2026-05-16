# Backoffice Automations — Comercializadora ACL

Quatro automações de backoffice usadas em uma comercializadora de energia no mercado livre (ACL):

1. **VBA RETUSD** — geração automática de Notas de Débito de ressarcimento TUSD em PDF
2. **N8N Criação de Unidades** — cadastro em massa de unidades consumidoras no ERP via API
3. **N8N Consulta CNPJ** — enriquecimento de cadastro com dados públicos de CNPJ
4. **PDF Consolidator** — coleta em lote de PDFs em estrutura de pastas (Python)

**Impacto combinado**: substitui processos manuais repetitivos do backoffice. Para os 3 fluxos correlatos da área (RETUSD + planilhas auxiliares de modulação), o tempo de execução caiu de ~1h para ~5min por ciclo mensal. O PDF Consolidator reduziu uma demanda pontual de 2-3h para ~1min.

> Versões anonimizadas (URLs internas substituídas por placeholders, dados de clientes removidos). Lógica e estrutura preservadas para fins de portfólio técnico.

---

## 1. VBA RETUSD — `vba-retusd/`

### O que faz

A **Nota de Débito de Ressarcimento TUSD (RETUSD)** é um documento mensal emitido pela comercializadora para cada cliente, baseado em relatórios da CCEE sobre desconto de TUSD. O processo manual envolvia:

1. Baixar relatórios da CCEE com saldos de desconto TUSD
2. Consolidar dados em planilha
3. Para cada cliente, copiar dados manualmente para um molde de ND
4. Salvar o documento como PDF
5. Repetir para todos os clientes do book

O VBA automatiza tudo a partir de uma planilha consolidada: para cada linha de cliente, um botão (checkbox) executa a geração da ND a partir do molde, com numeração sequencial, data de vencimento configurável, e exportação direta para PDF.

### Técnicas usadas

- **`SomaCorFonte`** (Módulo1): função UDF que soma valores em um range filtrando por cor de fonte. Útil quando a planilha consolidada usa cores para distinguir cliente A vs B na mesma coluna.
- **Eventos de checkbox por linha**: cada cliente tem seu próprio checkbox; ao marcar, dispara a macro com a linha correspondente.
- **Geração a partir de molde oculto**: copia a aba-molde, renomeia com numeração sequencial (`ND 2025 - 015 - <SUFIXO>`), preenche campos a partir da linha do cliente.
- **`InputBox` opcional**: data de vencimento pode ser informada ou cancelada (devolve o checkbox).
- **`ExportAsFixedFormat`** para PDF com nome padronizado.

### Arquivos

```
vba-retusd/
├── 01_sum_by_font_color.bas      # UDF SomaCorFonte
├── 02_generate_nd_from_template.bas  # Geração principal (Check_GerarND_RETUSD)
├── 03_pdf_export.bas             # Exportação PDF
└── 04_helpers.bas                # Funções auxiliares (parse de nome, etc.)
```

### Como usar (em workbook próprio)

1. Importe os 4 `.bas` no VBE (Alt+F11 → File → Import File)
2. Crie uma aba-molde oculta nomeada como `ND <ano> - <numero> - <sufixo>` (ex.: `ND 2025 - 014 - TEMPLATE`)
3. Em uma aba operacional, adicione checkboxes ActiveX por linha de cliente
4. Configure as colunas que serão lidas (B = sufixo do nome do cliente, AD = nome para D22, AE = CNPJ para E26)

---

## 2. N8N — Criação de Unidades — `n8n-unidades/`

### O que faz

Cadastra unidades consumidoras em massa no ERP interno via API REST, a partir de planilha de input. Fluxo:

1. **Form Trigger** com upload de planilha
2. **Parse** das linhas
3. **Authenticate** no ERP
4. **Loop**: para cada unidade, consulta cidade via `cityByState` (API auxiliar) e faz POST em `/rest/unit`
5. **Output**: resumo de unidades criadas com sucesso × erros

Antes: cadastro manual de cada unidade via UI do ERP, ~3 minutos por unidade × dezenas de unidades por ciclo.

### Arquivos

```
n8n-unidades/
├── workflow.json    # Export do workflow N8N (URLs sanitizadas)
└── README.md        # Como importar e configurar
```

### Sanitização para esta versão pública

- URL do ERP interno → `api-erp-interno.exemplo.com`
- Credenciais hardcoded → removidas (precisam ser reconfiguradas no N8N)
- IDs/UUIDs de nodes mantidos (não vazam informação)

---

## 3. N8N — Consulta CNPJ — `n8n-cnpj/`

### O que faz

Enriquece cadastro de clientes consultando dados públicos de CNPJ. Usa:

- **`api.cnpja.com`** como fonte primária (com token)
- **`brasilapi.com.br`** como fallback / fonte alternativa

Saída estruturada: razão social, endereço, CEP, atividades.

Antes: pesquisa manual cliente por cliente em sites de consulta CNPJ.

### Arquivos

```
n8n-cnpj/
├── workflow.json    # Export do workflow N8N (API pública, sem sanitização especial)
└── README.md
```

---

## 4. PDF Consolidator — `pdf-consolidator/`

### O que faz

Percorre uma estrutura de pastas (uma por cliente), localiza PDFs em subpastas padronizadas e consolida tudo em uma única pasta de destino. Usado em demanda pontual com ~370 pastas de clientes.

Não é só `os.walk + shutil.copy` — tem engenharia de produção:

- Validação dupla de PDF (extensão + cabeçalho binário `%PDF-`)
- Long path prefix do Windows (`\\?\`) para caminhos > 260 chars
- Suporte UNC (`\\Servidor\Compartilhamento`)
- Batches + workers paralelos configuráveis
- Dry-run, verificação de espaço livre, logging em arquivo

### Arquivos

```
pdf-consolidator/
├── pdf_consolidator.py    # Script principal
└── README.md              # Configuração e uso
```

---

## Como importar workflows N8N

1. Abra seu N8N (cloud ou self-hosted)
2. Workflows → Import from File → selecione `workflow.json`
3. Configure credenciais que aparecerão pendentes (botões com cadeado vermelho)
4. Para o workflow de unidades: ajuste a URL base da API do seu ERP

---

## Decisões de design

**Por que N8N e não Python puro?** N8N oferece interface visual + tratamento natural de erros por node + retry built-in. Para fluxos com 3-5 chamadas de API e poucos branches, o ganho de manutenibilidade vale mais que a flexibilidade de código.

**Por que VBA e não migrar pra Python?** O usuário final do RETUSD opera em Excel. Migrar pra Python obrigaria a equipe a abrir terminal/script, o que adiciona fricção. VBA dentro do próprio workbook mantém o fluxo "abre Excel, aperta botão" — a automação fica invisível ao usuário, que é o ponto.

**Por que Python (e não VBA/N8N) no PDF Consolidator?** Operação sobre filesystem em larga escala (~370 pastas) pede paralelismo e logging robustos. VBA é lento para esse tipo de coisa; N8N seria overkill. Python `concurrent.futures` resolve em dezenas de linhas.

---

## Sobre o autor

Analista de dados/energia em comercializadora do ACL. Automações implantadas em produção e em uso diário pela área. Esta versão pública preserva a lógica com dados/credenciais removidos.

## Licença

MIT.
