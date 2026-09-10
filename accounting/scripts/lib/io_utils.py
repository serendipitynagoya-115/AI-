"""出力・ログ・二重計上防止用の処理済み台帳を扱う共通ユーティリティ。

accounting/CLAUDE.md のルールに対応:
- data/・output/・logs/ はGit管理対象外(実データのため)。
- 同じ店舗・同じ月を再実行しても二重計上しない(常に上書き、追記しない)。
- 実行のたびにログを残す。
"""
from __future__ import annotations

import csv
import json
import logging
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2].parent  # accounting/scripts/lib -> accounting/scripts -> accounting -> repo root
ACCOUNTING_ROOT = REPO_ROOT / "accounting"
DATA_DIR = ACCOUNTING_ROOT / "data"
OUTPUT_DIR = ACCOUNTING_ROOT / "output"
LOGS_DIR = ACCOUNTING_ROOT / "logs"
LEDGER_PATH = DATA_DIR / "processed_ledger.json"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_json(path: Path, data) -> None:
    ensure_dir(path.parent)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    ensure_dir(path.parent)
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def setup_logger(name: str, store_id: str, year_month: str) -> tuple[logging.Logger, Path]:
    ensure_dir(LOGS_DIR)
    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    log_path = LOGS_DIR / f"{store_id}_{year_month}_{ts}.log"

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(sh)

    return logger, log_path


def load_ledger() -> dict:
    if not LEDGER_PATH.exists():
        return {}
    with open(LEDGER_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def update_ledger(store_id: str, year_month: str, source_file_sha256: str,
                   source_file_path: str, output_paths: list[str]) -> dict:
    """処理済み台帳を更新する。同じ店舗×月×ファイルハッシュで再実行しても、
    レコードを上書きするだけで追記(二重計上)はしない。"""
    ledger = load_ledger()
    key = f"{store_id}:{year_month}"
    previous = ledger.get(key)
    is_rerun_same_source = (
        previous is not None and previous.get("source_file_sha256") == source_file_sha256
    )
    ledger[key] = {
        "store_id": store_id,
        "year_month": year_month,
        "source_file_sha256": source_file_sha256,
        "source_file_path": source_file_path,
        "last_run_at": datetime.now(timezone.utc).isoformat(),
        "output_paths": output_paths,
        "run_count": (previous.get("run_count", 0) + 1) if previous else 1,
    }
    write_json(LEDGER_PATH, ledger)
    return {"is_rerun_same_source": is_rerun_same_source, "previous": previous}
