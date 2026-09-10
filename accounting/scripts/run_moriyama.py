#!/usr/bin/env python3
"""守山店 月次 日報集計・照合スクリプト(読み取り専用)。

やること:
  1. Google Driveミラーとして取得済みのローカルExcel(日報)を読み取る(書き込みはしない)。
  2. 日別シート(1〜31)から、新規・既存・物販を税抜で集計する(顧客名の有無では絞り込まない)。
  3. 区分(新規/既存)に矛盾がある取引・数式エラーの取引は「未確定」として別管理し、
     新規・既存いずれにも計上しない(合計には含める)。
  4. 日別合計・月合計を作る。
  5. 月報集計シートの数値と突き合わせる。
  6. 「2026 売上現状」のローカルスナップショット(事前にMCP経由で取得したもの)と突き合わせ、
     分類差異・金額差異を区別して検出する。
  7. 結果を accounting/output/ へ、ログを accounting/logs/ へ出力する。

やらないこと:
  - Google スプレッドシートへの書き込み(まだ実装しない)。
  - Dropbox・Google Driveミラー・「2026 売上現状」「2026店舗収支」の変更。

再実行しても、同じ入力ファイルであれば同じ内容で出力を上書きするだけであり、
加算・追記による二重計上は発生しない(accounting/data/processed_ledger.json で追跡)。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib import xlsx_report, reconcile, io_utils


def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--store-id", default="moriyama")
    p.add_argument("--year-month", default="2026-08")
    p.add_argument(
        "--source-xlsx",
        default=str(io_utils.DATA_DIR / "moriyama" / "2026-08" / "①8月.xlsx"),
        help="Google Driveミラーから取得済みのローカルExcelファイルパス",
    )
    p.add_argument(
        "--status-snapshot",
        default=str(io_utils.DATA_DIR / "moriyama" / "2026-08" / "sales_status_snapshot.json"),
        help="「2026 売上現状」のローカルスナップショットJSON(MCP経由で事前取得したもの)",
    )
    return p.parse_args()


def main():
    args = parse_args()
    store_id = args.store_id
    year_month = args.year_month
    source_path = Path(args.source_xlsx)
    status_path = Path(args.status_snapshot)

    logger, log_path = io_utils.setup_logger("moriyama_run", store_id, year_month)
    logger.info("=== 守山店 月次集計・照合スクリプト開始 ===")
    logger.info(f"対象: store_id={store_id}, year_month={year_month}")
    logger.info(f"入力ファイル(読み取り専用): {source_path}")

    if not source_path.exists():
        logger.error(f"入力ファイルが見つかりません: {source_path}")
        logger.error("先にGoogle Drive等から日報ミラーを取得し、このパスに配置してください。")
        sys.exit(1)

    # --- 二重計上防止: ソースファイルのハッシュを記録・比較 ---
    source_hash = xlsx_report.file_sha256(source_path)
    logger.info(f"入力ファイルのSHA-256: {source_hash}")

    # --- 1〜31日を読み取り、日別集計 ---
    logger.info("日別シート(1〜31)を読み取り中...")
    day_results = xlsx_report.extract_all_days(source_path)
    logger.info(f"読み取れた日数: {len(day_results)}")

    for day, r in sorted(day_results.items(), key=lambda kv: int(kv[0])):
        if not r.consistency_ok:
            logger.warning(
                f"[{day}日] 内部検算に差異: 再集計合計={r.grand_total} "
                f"vs ファイル自身のAG54={r.file_ag54} (差={r.consistency_diff})"
            )
        if r.unresolved_total != 0:
            logger.warning(
                f"[{day}日] 未確定(数式エラー等で金額が確定できない)取引あり: "
                f"{r.unresolved_total:g}円、{len(r.unresolved_rows)}件"
            )
        if r.category_flagged_rows:
            logger.warning(
                f"[{day}日] 区分(新規/既存)に矛盾のある取引を検知(金額は計上済み、要確認): "
                f"{len(r.category_flagged_rows)}件"
            )

    # --- 日別合計テーブルを出力 ---
    daily_rows = []
    for day, r in sorted(day_results.items(), key=lambda kv: int(kv[0])):
        daily_rows.append({
            "day": day,
            "new_excl_tax": r.new_total,
            "existing_excl_tax": r.existing_total,
            "retail_excl_tax": r.retail_total,
            "unresolved_excl_tax": r.unresolved_total,
            "total_excl_tax": r.grand_total,
            "transaction_count": r.transaction_count,
            "unresolved_count": len(r.unresolved_rows),
            "file_ag54_excl_tax": r.file_ag54,
            "consistency_ok": r.consistency_ok,
        })

    monthly_out_dir = io_utils.OUTPUT_DIR / "monthly" / year_month
    daily_csv_path = monthly_out_dir / f"{store_id}_daily_totals.csv"
    io_utils.write_csv(
        daily_csv_path, daily_rows,
        fieldnames=["day", "new_excl_tax", "existing_excl_tax", "retail_excl_tax",
                    "unresolved_excl_tax", "total_excl_tax", "transaction_count",
                    "unresolved_count", "file_ag54_excl_tax", "consistency_ok"],
    )
    logger.info(f"日別合計を出力: {daily_csv_path}")

    # --- 月合計、月報集計との照合 ---
    month_new = sum(r.new_total for r in day_results.values())
    month_existing = sum(r.existing_total for r in day_results.values())
    month_retail = sum(r.retail_total for r in day_results.values())
    month_unresolved = sum(r.unresolved_total for r in day_results.values())
    month_total = month_new + month_existing + month_retail + month_unresolved

    monthly_report = xlsx_report.extract_monthly_report_summary(source_path)
    month_summary = {
        "store_id": store_id,
        "year_month": year_month,
        "source_file": str(source_path),
        "source_file_sha256": source_hash,
        "days_read": sorted(day_results.keys(), key=lambda x: int(x)),
        "recomputed_from_daily_sheets_excl_tax": {
            "new": round(month_new, 2),
            "existing": round(month_existing, 2),
            "retail": round(month_retail, 2),
            "unresolved": round(month_unresolved, 2),
            "total": round(month_total, 2),
        },
        "monthly_report_sheet_excl_tax": monthly_report,
        "crosscheck": {
            "total_vs_実売実績(I5)": (
                None if monthly_report.get("実売実績_I5") is None
                else round(month_total - monthly_report["実売実績_I5"], 2)
            ),
            "retail_vs_内物販(I7)": (
                None if monthly_report.get("内物販_I7") is None
                else round(month_retail - monthly_report["内物販_I7"], 2)
            ),
        },
    }
    month_summary_path = monthly_out_dir / f"{store_id}_month_summary.json"
    io_utils.write_json(month_summary_path, month_summary)
    logger.info(f"月合計・月報集計との照合結果を出力: {month_summary_path}")

    diff_i5 = month_summary["crosscheck"]["total_vs_実売実績(I5)"]
    if diff_i5 is not None and abs(diff_i5) >= 1.0:
        logger.warning(f"月報集計(実売実績 I5)との差異を検出: {diff_i5}円")
    else:
        logger.info("月報集計(実売実績 I5)と一致(誤差1円未満)")
    diff_i7 = month_summary["crosscheck"]["retail_vs_内物販(I7)"]
    if diff_i7 is not None and abs(diff_i7) >= 1.0:
        logger.warning(f"月報集計(内物販 I7)との差異を検出: {diff_i7}円")
    else:
        logger.info("月報集計(内物販 I7)と一致(誤差1円未満)")

    # --- 「2026 売上現状」との照合 ---
    reconciliation_out_dir = io_utils.OUTPUT_DIR / "reconciliation" / year_month
    if status_path.exists():
        with open(status_path, "r", encoding="utf-8") as f:
            status_snapshot = json.load(f)
        status_daily = status_snapshot.get("daily", {})
        comparisons = reconcile.build_comparison_table(day_results, status_daily)

        comparison_rows = []
        verdict_counts = {}
        for c in comparisons:
            verdict_counts[c.verdict] = verdict_counts.get(c.verdict, 0) + 1
            comparison_rows.append({
                "day": c.date_key,
                "verdict": c.verdict,
                "report_new": c.report_new, "status_new": c.status_new, "diff_new": c.diff_new,
                "report_existing": c.report_existing, "status_existing": c.status_existing,
                "diff_existing": c.diff_existing,
                "report_retail": c.report_retail, "status_retail": c.status_retail,
                "diff_retail": c.diff_retail,
                "report_total": c.report_total, "status_total": c.status_total,
                "diff_total": c.diff_total,
                "report_unresolved": c.report_unresolved,
                "note": c.note,
            })
            if c.verdict in ("分類差異", "金額差異"):
                logger.warning(
                    f"[{c.date_key}日] {c.verdict}: "
                    f"新規 日報={c.report_new}/売上現状={c.status_new}(差{c.diff_new}), "
                    f"既存 日報={c.report_existing}/売上現状={c.status_existing}(差{c.diff_existing}), "
                    f"物販 日報={c.report_retail}/売上現状={c.status_retail}(差{c.diff_retail}), "
                    f"累計 日報={c.report_total}/売上現状={c.status_total}(差{c.diff_total})"
                )
            elif c.verdict == "一致":
                logger.info(f"[{c.date_key}日] 一致")

        comparison_csv_path = reconciliation_out_dir / f"{store_id}_vs_sales_status.csv"
        io_utils.write_csv(
            comparison_csv_path, comparison_rows,
            fieldnames=["day", "verdict", "report_new", "status_new", "diff_new",
                        "report_existing", "status_existing", "diff_existing",
                        "report_retail", "status_retail", "diff_retail",
                        "report_total", "status_total", "diff_total",
                        "report_unresolved", "note"],
        )
        logger.info(f"「2026 売上現状」との照合結果を出力: {comparison_csv_path}")
        logger.info(f"判定件数: {verdict_counts}")
    else:
        logger.warning(f"「2026 売上現状」のスナップショットが見つかりません: {status_path}")
        logger.warning("照合をスキップしました(読み取り専用のスナップショットを別途用意してください)。")

    # --- 未確定取引の一覧化 ---
    unresolved_rows = []
    for day, r in sorted(day_results.items(), key=lambda kv: int(kv[0])):
        for tx in r.unresolved_rows:
            unresolved_rows.append({
                "day": day,
                "row": tx.row,
                "customer_name": tx.customer_name or "(空欄)",
                "staff": tx.staff or "",
                "category_label": tx.category_label or "",
                "course": tx.course or "",
                "product": tx.product or "",
                "unresolved_amount_excl_tax": tx.unresolved_amount,
                "has_error": tx.has_error,
                "reasons": " / ".join(tx.unresolved_reasons),
            })
    unresolved_csv_path = reconciliation_out_dir / f"{store_id}_unresolved_transactions.csv"
    io_utils.write_csv(
        unresolved_csv_path, unresolved_rows,
        fieldnames=["day", "row", "customer_name", "staff", "category_label", "course",
                    "product", "unresolved_amount_excl_tax", "has_error", "reasons"],
    )
    logger.info(f"未確定取引一覧(数式エラー等で金額が確定できないもの)を出力: {unresolved_csv_path}({len(unresolved_rows)}件)")

    # --- 区分(新規/既存)矛盾の一覧化(金額は計上済み。参考・要確認情報) ---
    category_flag_rows = []
    for day, r in sorted(day_results.items(), key=lambda kv: int(kv[0])):
        for tx in r.category_flagged_rows:
            category_flag_rows.append({
                "day": day,
                "row": tx.row,
                "customer_name": tx.customer_name or "(空欄)",
                "staff": tx.staff or "",
                "category_label": tx.category_label or "",
                "course": tx.course or "",
                "new_amount_excl_tax": tx.new_amount,
                "existing_amount_excl_tax": tx.existing_amount,
                "note": tx.category_flag,
            })
    category_flags_csv_path = reconciliation_out_dir / f"{store_id}_category_flags.csv"
    io_utils.write_csv(
        category_flags_csv_path, category_flag_rows,
        fieldnames=["day", "row", "customer_name", "staff", "category_label", "course",
                    "new_amount_excl_tax", "existing_amount_excl_tax", "note"],
    )
    logger.info(
        f"区分矛盾の要確認一覧(金額は集計済み・自動補正なし)を出力: "
        f"{category_flags_csv_path}({len(category_flag_rows)}件)"
    )

    # --- 処理済み台帳の更新(二重計上防止) ---
    output_paths = [
        str(daily_csv_path), str(month_summary_path), str(unresolved_csv_path),
        str(category_flags_csv_path),
    ]
    if status_path.exists():
        output_paths.append(str(comparison_csv_path))
    ledger_info = io_utils.update_ledger(
        store_id=store_id, year_month=year_month,
        source_file_sha256=source_hash, source_file_path=str(source_path),
        output_paths=output_paths,
    )
    if ledger_info["is_rerun_same_source"]:
        logger.info(
            "同一ソースファイル(ハッシュ一致)での再実行です。出力は上書きされ、二重計上は発生していません。"
        )
    else:
        logger.info("処理済み台帳を更新しました。")

    logger.info("=== 完了 ===")
    logger.info(f"ログファイル: {log_path}")


if __name__ == "__main__":
    main()
