#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
copy_pdfs.py
--------------------------------------
Consolida SOMENTE arquivos .pdf de múltiplas pastas de clientes para uma pasta
de destino. Descobre automaticamente as pastas de clientes a partir de um
caminho-base e adiciona o subcaminho fixo "subpasta_alvo".

⚠️ Segurança: copia EXCLUSIVAMENTE .pdf — valida extensão e cabeçalho binário "%PDF-".

Exemplos:
  python copy_pdfs.py --dest "D:\\Consolidado" --dry-run
  python copy_pdfs.py --dest "D:\\Consolidado" --batch-size 500
  python copy_pdfs.py --dest "\\\\Servidor\\Compart\\PDFs" --workers 4
"""

import argparse
import logging
import os
import shutil
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Generator, Iterable, List, Optional, Tuple

# ================== CONFIGURAÇÃO (EDITE AQUI SE PRECISAR) ===================

# 1) Pasta-base que contém todas as pastas de clientes do mês/ano:
BASE_CLIENTS_DIR = r"C:\caminho\para\pastas\de\clientes\AAAA\MM-MES"

# 2) Subpasta fixa (igual em todos os clientes) onde estão os PDFs:
SUBPATH_IN_CLIENT = r"subpasta_alvo"

# 3) Ajustes de desempenho:
DEFAULT_BATCH_SIZE = 500        # tamanho do lote
DEFAULT_WORKERS = 0             # 0/1=sequencial; >1 paraleliza I/O

# ===========================================================================

def add_long_path_prefix(path: str) -> str:
    """Adiciona prefixo \\?\\ para suportar caminhos > 260 caracteres no Windows."""
    if os.name == "nt":
        ap = os.path.abspath(path)
        if ap.startswith("\\\\?\\"):
            return ap
        if ap.startswith("\\\\"):  # UNC
            return "\\\\?\\UNC\\" + ap.lstrip("\\\\")
        return "\\\\?\\" + ap
    return path

def is_pdf_file(path: str) -> bool:
    """Aceita somente arquivos .pdf e com cabeçalho '%PDF-'."""
    _, ext = os.path.splitext(path)
    if ext.lower() != ".pdf":
        return False
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False

def build_source_map(base_clients_dir: str, subpath_in_client: str) -> Dict[str, str]:
    """
    Cria SOURCE_MAP automaticamente:
      { '<nome_cliente>': '<base>/<nome_cliente>/<subpath_in_client>' }
    Inclui somente clientes onde esse subcaminho existe.
    """
    result: Dict[str, str] = {}
    base = add_long_path_prefix(base_clients_dir)
    if not os.path.isdir(base):
        raise RuntimeError(f"Pasta-base não encontrada: {base_clients_dir}")
    for entry in os.listdir(base):
        full = os.path.join(base, entry)
        try:
            if not os.path.isdir(full):
                continue
            candidate = os.path.join(full, subpath_in_client)
            if os.path.isdir(add_long_path_prefix(candidate)):
                # usa o nome da pasta do cliente como chave
                result[entry] = candidate
            else:
                logging.debug("Subpasta não encontrada (ignorado): %s", candidate)
        except PermissionError:
            logging.warning("Sem permissão para acessar: %s", full)
        except Exception as exc:
            logging.warning("Falha ao inspecionar '%s': %s", full, exc)
    return result

def walk_pdfs(source_map: Dict[str, str]) -> Generator[Tuple[str, str, int], None, None]:
    """Gera (src_key, caminho_pdf, tamanho) para PDFs VÁLIDOS em todas as origens."""
    for src_key, root in source_map.items():
        if not root:
            continue
        root_lp = add_long_path_prefix(root)
        if not os.path.exists(root_lp):
            logging.warning("Origem inexistente: %s (%s)", src_key, root)
            continue
        for dirpath, _dirnames, filenames in os.walk(root_lp):
            for name in filenames:
                full = os.path.join(dirpath, name)
                try:
                    if os.path.splitext(name)[1].lower() != ".pdf":
                        continue
                    if not is_pdf_file(full):
                        logging.warning("Ignorado (não é PDF válido): %s", full)
                        continue
                    size = os.path.getsize(full)
                    yield (src_key, full, size)
                except Exception as exc:
                    logging.error("Falha ao inspecionar '%s': %s", full, exc)

def human_bytes(n: int) -> str:
    step = 1024.0
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if n < step:
            return f"{n:,.2f} {unit}".replace(",", ".")
        n /= step
    return f"{n:.2f} PB"

def ensure_dir(path: str) -> None:
    os.makedirs(add_long_path_prefix(path), exist_ok=True)

def disk_free_bytes(path: str) -> int:
    probe = path
    while probe and not os.path.exists(probe):
        probe = os.path.dirname(probe) or None
    if not probe:
        probe = os.path.dirname(path) or os.getcwd()
    usage = shutil.disk_usage(add_long_path_prefix(probe))
    return usage.free

def unique_dest_path(dest_dir: str, base_name: str, src_key: str) -> str:
    """Gera nome único em caso de colisão, preservando o original."""
    name, ext = os.path.splitext(base_name)
    assert ext.lower() == ".pdf"
    candidate = os.path.join(dest_dir, base_name)
    if not os.path.exists(add_long_path_prefix(candidate)):
        return candidate
    n = 1
    while True:
        alt = f"{name} (src-{src_key}, dup{n}){ext}"
        alt_path = os.path.join(dest_dir, alt)
        if not os.path.exists(add_long_path_prefix(alt_path)):
            return alt_path
        n += 1

def copy_one(src_key: str, src_path: str, dest_dir: str) -> Tuple[bool, str, Optional[str]]:
    """Copia um PDF (copy2) e retorna (ok, destino, erro)."""
    try:
        if not is_pdf_file(src_path):  # dupla verificação
            return (False, "", f"Bloqueado: não é PDF válido -> {src_path}")
        base = os.path.basename(src_path)
        dest_path = unique_dest_path(dest_dir, base, src_key)
        ensure_dir(dest_dir)
        shutil.copy2(add_long_path_prefix(src_path), add_long_path_prefix(dest_path))
        return (True, dest_path, None)
    except Exception as exc:
        return (False, "", f"Erro ao copiar '{src_path}': {exc}")

def setup_logging(dest_dir: str, verbose: bool = True) -> None:
    ensure_dir(dest_dir)
    log_file = os.path.join(dest_dir, "copy_pdfs.log")
    handlers = [logging.FileHandler(log_file, encoding="utf-8")]
    if verbose:
        handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)s | %(message)s",
                        handlers=handlers)
    logging.info("Log em: %s", log_file)

def copy_in_batches(
    items: Iterable[Tuple[str, str, int]],
    dest_dir: str,
    batch_size: int,
    workers: int,
    dry_run: bool,
    total_count: int,
) -> None:
    copied = 0
    failed = 0
    processed = 0
    start = time.time()

    def _process(seq: List[Tuple[str, str, int]]):
        nonlocal copied, failed
        if dry_run:
            for (src_key, src_path, _size) in seq:
                logging.info("[DRY-RUN] %s -> %s", src_path, dest_dir)
                copied += 1
            return
        if workers and workers > 1:
            with ThreadPoolExecutor(max_workers=workers) as ex:
                futs = [ex.submit(copy_one, k, p, dest_dir) for (k, p, _s) in seq]
                for fut in as_completed(futs):
                    ok, dst, err = fut.result()
                    if ok:
                        copied += 1
                        logging.info("Copiado: %s", dst)
                    else:
                        failed += 1
                        logging.error("%s", err)
        else:
            for (k, p, _s) in seq:
                ok, dst, err = copy_one(k, p, dest_dir)
                if ok:
                    copied += 1
                    logging.info("Copiado: %s", dst)
                else:
                    failed += 1
                    logging.error("%s", err)

    batch: List[Tuple[str, str, int]] = []
    for item in items:
        batch.append(item)
        if len(batch) >= batch_size:
            _process(batch); processed += len(batch); batch.clear()
            pct = processed / max(total_count, 1) * 100.0
            logging.info("Progresso: %.2f%% (%d/%d) | Falhas: %d | Tempo: %.1fs",
                         pct, processed, total_count, failed, time.time()-start)
    if batch:
        _process(batch); processed += len(batch)

    logging.info("FINALIZADO. Total: %d | Copiados: %d | Falhas: %d | Tempo: %.1fs",
                 total_count, copied, failed, time.time()-start)

def main():
    parser = argparse.ArgumentParser(
        description="Copia SOMENTE .pdf de múltiplas pastas de clientes (auto-descobertas) para uma pasta destino."
    )
    parser.add_argument("--dest", default=r"C:\caminho\para\consolidado",
                    help="Pasta de destino onde os PDFs serão consolidados (default acima).")
    parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE, help="Tamanho do lote.")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="Threads de cópia (0/1=sequencial).")
    parser.add_argument("--dry-run", action="store_true", help="Apenas simula; não copia.")
    parser.add_argument("--ignore-space-check", action="store_true", help="Ignora verificação de espaço livre.")
    args = parser.parse_args()

    dest_dir = args.dest
    setup_logging(dest_dir, verbose=True)

    # Gera o SOURCE_MAP automaticamente
    logging.info("Base: %s", BASE_CLIENTS_DIR)
    logging.info("Subpasta fixa: %s", SUBPATH_IN_CLIENT)
    try:
        source_map = build_source_map(BASE_CLIENTS_DIR, SUBPATH_IN_CLIENT)
    except Exception as e:
        logging.error(str(e))
        sys.exit(2)

    if not source_map:
        logging.error("Nenhuma origem válida encontrada. Verifique BASE_CLIENTS_DIR e SUBPATH_IN_CLIENT.")
        sys.exit(2)

    logging.info("Pastas de origem detectadas: %d", len(source_map))

    # 1) Contagem/estimativa de tamanho
    total_files = 0
    total_bytes = 0
    for (_k, _p, size) in walk_pdfs(source_map):
        total_files += 1
        total_bytes += size
    logging.info("PDFs válidos: %d | Tamanho total: %s", total_files, human_bytes(total_bytes))
    if total_files == 0:
        logging.info("Nada a copiar.")
        return

    # 2) Espaço livre
    if not args.ignore_space_check and not args.dry_run:
        free = disk_free_bytes(dest_dir)
        logging.info("Espaço livre no destino: %s", human_bytes(free))
        if total_bytes > free:
            logging.error("Espaço insuficiente (necessário %s, disponível %s). Abortando.",
                          human_bytes(total_bytes), human_bytes(free))
            sys.exit(3)

    # 3) Cópia em lotes (baixa memória)
    copy_in_batches(
        items=walk_pdfs(source_map),
        dest_dir=dest_dir,
        batch_size=max(1, args.batch_size),
        workers=max(0, args.workers),
        dry_run=args.dry_run,
        total_count=total_files,
    )

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.", file=sys.stderr)
        sys.exit(130)
