"""Monthly Accounting Layer(月次会計確定レイヤー)の実行スクリプト。

既存の物販監査(product_audit.py / run_product_audit.py)の出力(JSON)を「上から」
読み取り、店舗ごとの月次確定データJSON、および全社統合サマリー(JSON/CSV)を生成する。

このスクリプトはExcel原本を読み取り専用でのみ開く(AC/AD直接合計・月報集計セルの
状態確認のため)。書き込み・再計算・保存・外部リンク更新は一切行わない。

使い方:
    python3 scripts/run_monthly_accounting.py --year-month 2026-08
    python3 scripts/run_monthly_accounting.py --year-month 2026-08 --store-id moriyama
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib import io_utils, monthly_accounting


def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--year-month", default="2026-08")
    p.add_argument(
        "--store-id", action="append", default=None,
        help="対象店舗(複数指定可)。未指定なら全店舗(monthly_accounting.STORE_IDS_2026_08)。",
    )
    return p.parse_args()


def main():
    args = parse_args()
    year_month = args.year_month
    store_ids = args.store_id or monthly_accounting.STORE_IDS_2026_08

    logger, log_path = io_utils.setup_logger("monthly_accounting", "company", year_month)
    logger.info("=== Monthly Accounting Layer 開始 ===")
    logger.info(f"対象年月: {year_month} / 対象店舗: {store_ids}")

    store_master = monthly_accounting.load_store_master()
    for store_id in store_ids:
        if store_id not in store_master:
            logger.error(f"{store_id}: config/stores.yamlに登録がありません。")
        elif store_master[store_id].get("store_type") is None:
            logger.error(f"{store_id}: store_typeが未設定です(config/stores.yaml)。")

    out_dir = io_utils.OUTPUT_DIR / "monthly" / year_month
    io_utils.ensure_dir(out_dir)

    store_records = []
    for store_id in store_ids:
        logger.info(f"--- {store_id} ---")
        try:
            record = monthly_accounting.build_store_monthly_record(store_id, year_month, store_master)
        except FileNotFoundError as e:
            logger.error(f"{store_id}: {e}")
            continue
        store_records.append(record)

        store_path = out_dir / f"{store_id}.json"
        io_utils.write_json(store_path, record)
        logger.info(f"月次確定データを出力: {store_path}")

        for key, check in record["reconciliation"].items():
            status = check["status"]
            if status == "fail":
                logger.error(f"  [{key}] FAIL: lhs={check['lhs']} rhs={check['rhs']} diff={check['diff']}")
            elif status == "known_difference":
                logger.info(f"  [{key}] known_difference: diff={check['diff']} ({check['note']})")
            elif status == "pass":
                logger.info(f"  [{key}] pass (diff={check['diff']})")
            else:
                logger.info(f"  [{key}] unavailable")

        msr = record["monthly_summary_reconciliation"]
        logger.info(
            f"  [monthly_summary_reconciliation] status={msr['status']} / "
            f"source_workbook_reported={msr['source_workbook_reported']} / "
            f"authoritative_daily_total={msr['authoritative_daily_total']} / diff={msr['difference']}"
        )
        rr = record["retail_reconciliation"]
        logger.info(
            f"  [retail_reconciliation] raw_daily_af_total={rr['raw_daily_af_total']} / "
            f"transaction_rounded_total={rr['transaction_rounded_total']} / diff={rr['rounding_difference']}"
        )
        gp = record["retail_gross_profit"]
        margin = record["retail_gross_margin"]
        logger.info(
            f"  source_current_sales.total={record['source_current_sales']['total']} / "
            f"management_accounting_sales.total={record['management_accounting_sales']['total']} / "
            f"retail_gross_profit={gp} / retail_gross_margin={margin}"
        )
        pr = record["pl_readiness"]
        logger.info(f"  [pl_readiness] {pr}")

    if not store_records:
        logger.error("有効な店舗が1件もありません。company_summaryは生成しません。")
        return

    company_summary = monthly_accounting.build_company_summary(store_records, year_month)
    company_json_path = out_dir / "company_summary.json"
    io_utils.write_json(company_json_path, company_summary)
    logger.info(f"全社統合サマリー(JSON)を出力: {company_json_path}")

    csv_rows = monthly_accounting.company_summary_csv_rows(store_records)
    company_csv_path = out_dir / "company_summary.csv"
    io_utils.write_csv(company_csv_path, csv_rows, fieldnames=monthly_accounting.CSV_FIELDNAMES)
    logger.info(f"全社統合サマリー(CSV)を出力: {company_csv_path}")

    logger.info(
        f"authoritative_management_company_sales_total="
        f"{company_summary['authoritative_management_company_sales_total']} / "
        f"source_workbook_reported_total={company_summary['source_workbook_reported_total']}"
    )
    logger.info(
        f"company retail_gross_profit={company_summary['company_retail_gross_profit']} / "
        f"company retail_gross_margin={company_summary['company_retail_gross_margin']} / "
        f"confirmed_cogs_total={company_summary['retail_cogs_summary']['confirmed_cogs_total']} / "
        f"unconfirmed_cogs_sales_total={company_summary['retail_cogs_summary']['unconfirmed_cogs_sales_total']} / "
        f"cogs_fully_confirmed_all_stores={company_summary['retail_cogs_summary']['cogs_fully_confirmed_all_stores']}"
    )
    logger.info(f"pl_readiness_summary={company_summary['pl_readiness_summary']}")
    logger.info(f"stores_with_unconfirmed_retail_cogs={company_summary['stores_with_unconfirmed_retail_cogs']}")
    logger.info(f"stores_with_payment_pending={company_summary['stores_with_payment_pending']}")
    logger.info(f"stores_with_tax_pending={company_summary['stores_with_tax_pending']}")
    logger.info(f"stores_with_known_structural_difference={company_summary['stores_with_known_structural_difference']}")
    logger.info("=== 完了 ===")
    logger.info(f"ログファイル: {log_path}")


if __name__ == "__main__":
    main()
