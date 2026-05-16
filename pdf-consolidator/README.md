# PDF Consolidator — Coleta em Lote de PDFs em Estrutura de Pastas

Script Python que percorre uma estrutura de pastas (uma por cliente), localiza arquivos PDF dentro de uma subpasta padronizada e os consolida em uma pasta única de destino.

**Caso de uso original**: backoffice de comercializadora com ~370 pastas de clientes, demanda pontual de consolidar PDFs de TUSD para envio a auditoria/contabilidade.

**Redução de tempo**: ~2-3 horas (cópia manual) → **~1 min** (execução do script).

---

## Como funciona

Dada uma estrutura como:

```
<BASE_CLIENTS_DIR>/
├── Cliente A/
│   └── subpasta_alvo/
│       ├── doc1.pdf
│       └── doc2.pdf
├── Cliente B/
│   └── subpasta_alvo/
│       └── doc3.pdf
└── Cliente C/
    └── subpasta_alvo/
        └── doc4.pdf
```

O script:

1. Descobre automaticamente todas as pastas de cliente em `BASE_CLIENTS_DIR`
2. Para cada cliente, navega para a `subpasta_alvo`
3. Valida cada `.pdf` por extensão **e** cabeçalho binário (`%PDF-`)
4. Copia para o destino, gerando nome único em caso de colisão

## Engenharia de robustez

Não é só `os.walk + shutil.copy`. As decisões abaixo distinguem script de fim de semana de ferramenta de produção:

- **Validação dupla**: extensão `.pdf` + cabeçalho `%PDF-` no arquivo. Defesa contra arquivos malformados ou renomeados.
- **Long path prefix do Windows** (`\\?\`): suporta caminhos > 260 caracteres. Pastas reais frequentemente extrapolam esse limite.
- **Suporte UNC**: caminhos de rede como `\\Servidor\Compartilhamento` são tratados corretamente.
- **Batches + workers configuráveis**: cópia paralelizada por threads, com tamanho de lote configurável (default 500). Para volumes grandes em discos lentos, ajuda muito.
- **Dry-run**: simula a operação sem copiar. Essencial antes do primeiro run em produção.
- **Verificação de espaço livre**: aborta antes de começar se o destino não comporta o volume total.
- **Logging em arquivo + console**: registra cada cópia individual, com tempo total e percentual de progresso.

## Configuração

Edite o topo do arquivo:

```python
BASE_CLIENTS_DIR = r"C:\caminho\para\pastas\de\clientes\AAAA\MM-MES"
SUBPATH_IN_CLIENT = r"subpasta_alvo"   # subpasta padronizada dentro de cada cliente
```

E rode:

```bash
# Simulação:
python pdf_consolidator.py --dest "D:\Consolidado" --dry-run

# Execução real, lote padrão:
python pdf_consolidator.py --dest "D:\Consolidado"

# Com 4 threads:
python pdf_consolidator.py --dest "D:\Consolidado" --workers 4

# Caminho UNC:
python pdf_consolidator.py --dest "\\Servidor\Compart\PDFs" --workers 4
```

## Limitações

- **Sem deduplicação por hash**: dois PDFs idênticos vindos de pastas diferentes são tratados como arquivos distintos (sufixados `(src-clienteX, dup1)`). Se quiser deduplicar por conteúdo, adicione um pre-pass com `hashlib.sha256`.
- **Não preserva estrutura de origem**: tudo vai pra mesma pasta. Útil quando o objetivo é entregar todos os PDFs num único lote; ruim quando você precisa rastrear a origem.
- **Windows-first**: prefixo de long path e UNC são específicos do Windows. Em Linux/Mac funciona, mas o prefixo `\\?\` é ignorado.
