# accounting/scripts/ について

日報から売上を集計し、管理用集計表(「2026 売上現状」等)と照合するための読み取り専用スクリプトです。**Googleスプレッドシートへの書き込み機能はまだありません。** 元ファイル(Dropbox原本・Google Driveミラー・Googleスプレッドシート)はどのスクリプトからも一切変更しません。

## 構成

```
scripts/
├── lib/
│   ├── xlsx_report.py   日報Excel(日別シート)の読み取り・集計ロジック
│   ├── reconcile.py     管理用集計表との突き合わせ・差異分類ロジック
│   └── io_utils.py      出力・ログ・処理済み台帳(二重計上防止)の共通処理
└── run_moriyama.py      守山店 月次集計・照合の実行スクリプト(現時点の唯一のエントリポイント)
```

## 前提: 入力データの取得は、このスクリプトの範囲外

このスクリプトはPythonの標準的な実行環境で動作し、Dropbox・Google DriveのAPIを直接呼び出す機能は持っていません(認証情報をこのリポジトリに置かない方針のため)。したがって、以下の2つの入力ファイルは、**事前にClaude(MCP経由の読み取り専用アクセス)によって取得し、ローカルに保存しておく必要があります。**

1. **日報のExcel原本**(店舗のGoogle Driveミラー、またはDropbox原本を都度取得したもの)
   - 保存先の既定値: `accounting/data/<store_id>/<year-month>/<ファイル名>.xlsx`
   - 例: `accounting/data/moriyama/2026-08/①8月.xlsx`
2. **「2026 売上現状」のスナップショットJSON**(MCP経由で読み取った日次の新規・既存・物販・累計をキャッシュしたもの)
   - 保存先の既定値: `accounting/data/<store_id>/<year-month>/sales_status_snapshot.json`
   - フォーマットは `accounting/data/moriyama/2026-08/sales_status_snapshot.json` を参照

どちらも `accounting/data/` 配下(Git管理対象外)に置きます。

## 実行方法

```bash
python3 accounting/scripts/run_moriyama.py \
  --store-id moriyama \
  --year-month 2026-08 \
  --source-xlsx "accounting/data/moriyama/2026-08/①8月.xlsx" \
  --status-snapshot "accounting/data/moriyama/2026-08/sales_status_snapshot.json"
```

引数を省略した場合、上記の守山店・2026年8月のパスが既定値として使われます。

## 出力

- `accounting/output/monthly/<year-month>/<store_id>_daily_totals.csv`:日別の新規・既存・物販・累計(税抜)、未確定金額、内部検算結果
- `accounting/output/monthly/<year-month>/<store_id>_month_summary.json`:月合計と、日報の`月報集計`シート(実売実績・内物販)との突き合わせ結果
- `accounting/output/reconciliation/<year-month>/<store_id>_vs_sales_status.csv`:日ごとの「2026 売上現状」との突き合わせ結果。判定は「一致」「分類差異」「金額差異」「データなし(売上現状未入力)」のいずれか
- `accounting/output/reconciliation/<year-month>/<store_id>_unresolved_transactions.csv`:数式エラー等で金額そのものが確定できない取引(通常は0件)
- `accounting/output/reconciliation/<year-month>/<store_id>_category_flags.csv`:区分(新規/既存)ラベルと実際の計算根拠(コース名等)が食い違う取引。**金額は日報の計算式どおり集計済みで、自動での付け替えは行っていません。** オーナー確認用の参考情報です。

いずれも `accounting/output/`(Git管理対象外)に出力されます。

## ログ

`accounting/logs/<store_id>_<year-month>_<実行日時>.log` に、処理件数・警告(差異検知・区分矛盾など)を記録します。

## 集計・分類のルール

`accounting/docs/accounting-spec.md` 第9章、`accounting/docs/moriyama-sales-mapping.md` を参照してください。要点:

- 日報を正本とし、顧客名の有無では絞り込まない。
- 新規売(AC列)・既存売(AD列)・物販(AF列)は、日報自身の計算式の値をそのまま採用する。1行で新規売と既存売が同時に発生する取引(例:体験当日に回数券を購入)が実在することを確認済みであり、正常な取引として扱う。
- 区分ラベル(D列)と計算根拠(コース名等)が食い違う場合は、金額は計上したうえで「要確認」として別途一覧化する(自動での分類変更はしない)。
- 「2026 売上現状」との比較で、合計は一致するが内訳だけ違う場合は「分類差異」、合計自体が違う場合は「金額差異」として区別する。
- 日報・「2026 売上現状」のいずれも自動で書き換えない。

## 二重計上防止

`accounting/data/processed_ledger.json`(Git管理対象外)に、店舗×年月ごとの入力ファイルのSHA-256ハッシュと実行履歴を記録します。同じ入力ファイルで再実行しても、出力ファイルは毎回上書きされるだけで、加算・追記は行われません。
