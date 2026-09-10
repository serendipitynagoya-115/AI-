"""日報(店舗別 月次Excel)の読み取り専用パーサー。

accounting/docs/moriyama-sales-mapping.md で確定したセル参照(守山店のレイアウト)に基づく。
ファイルへの書き込みは一切行わない(openpyxlはload_workbookのみ使用し、saveは呼ばない)。

分類ルール(accounting/docs/accounting-spec.md 第9章、accounting/CLAUDE.md 準拠):
- 日報を正本とする。顧客名の有無では絞り込まない。
- 新規/既存の区分(D列)と、実際に計算された金額の根拠(コース名等)が矛盾する行は
  「未確定」として新規・既存いずれにも計上せず、別管理する(合計には含める)。
- 金額は推測しない。数式エラー(#N/A等)のセルは0円として扱わず、要確認として分離する。
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

HEADER_ROW = 3
DATA_START_ROW = 4
DATA_END_ROW = 53
TOTAL_ROW = 54
OVERFLOW_START_ROW = 58
OVERFLOW_END_ROW = 62
OVERFLOW_TOTAL_ROW = 63

DAY_SHEETS = [str(d) for d in range(1, 32)]


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _is_error_value(v) -> bool:
    return isinstance(v, str) and v.strip().startswith("#")


def _numeric_or_none(v):
    """セル値を数値として解釈する。数式エラーや空欄は None を返す(0円と断定しない)。"""
    if v is None:
        return None
    if _is_error_value(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    # 想定外の文字列型(通常は発生しない想定。発生した場合は要確認として扱う)
    return None


@dataclass
class TransactionRow:
    sheet: str
    row: int
    customer_name: str | None
    staff: str | None
    category_label: str | None  # D列の値(新規/既存/物販/空欄等)
    course: str | None
    product: str | None
    new_amount: float = 0.0       # 新規売(税別、AC列)。金額が入っていれば区分ラベルに関わらず計上する
    existing_amount: float = 0.0  # 既存売(税別、回数券、AD列)。同上
    retail_amount: float = 0.0    # 物販(税別、AF列)
    unresolved_amount: float = 0.0  # 数式エラー等で「金額そのもの」が確定できないぶんのみ
    unresolved_reasons: list[str] = field(default_factory=list)
    has_error: bool = False
    # 金額は正しく計上した上で、区分ラベル(D列)と実際の計算根拠(コース名等)が
    # 食い違う場合の「参考情報」。集計金額には影響させない(ルール7・12: 推測で補正しない)。
    category_flag: str | None = None


@dataclass
class DayResult:
    sheet: str
    new_total: float
    existing_total: float
    retail_total: float
    unresolved_total: float
    grand_total: float
    transaction_count: int
    unresolved_rows: list[TransactionRow]  # 数式エラー等で金額そのものが確定できない行
    error_rows: list[TransactionRow]
    category_flagged_rows: list[TransactionRow]  # 金額は計上済みだが区分ラベルが食い違う行(参考情報)
    # 検算用: ファイル自身のSUMIFセル(AC54/AD54/AF54/AG54)の値
    file_ac54: float | None
    file_ad54: float | None
    file_af54: float | None
    file_ag54: float | None
    consistency_ok: bool  # (new+existing+retail+unresolved) と file_ag54 が一致するか
    consistency_diff: float | None


def _process_row(ws_v, sheet: str, row: int) -> TransactionRow | None:
    b = ws_v[f"B{row}"].value
    c = ws_v[f"C{row}"].value
    d = ws_v[f"D{row}"].value
    f = ws_v[f"F{row}"].value
    j = ws_v[f"J{row}"].value

    ac_raw = ws_v[f"AC{row}"].value
    ad_raw = ws_v[f"AD{row}"].value
    af_raw = ws_v[f"AF{row}"].value
    ag_raw = ws_v[f"AG{row}"].value

    ac = _numeric_or_none(ac_raw)
    ad = _numeric_or_none(ad_raw)
    af = _numeric_or_none(af_raw)

    is_empty_template_row = (
        b in (None, "") and c in (None, "") and d in (None, "") and
        f in (None, "") and j in (None, "") and
        (ac or 0) == 0 and (ad or 0) == 0 and (af or 0) == 0
    )
    if is_empty_template_row:
        return None

    tx = TransactionRow(
        sheet=sheet, row=row, customer_name=b, staff=c,
        category_label=d, course=f, product=j,
    )

    error_cells = []
    if _is_error_value(ac_raw):
        error_cells.append(("AC", ac_raw))
    if _is_error_value(ad_raw):
        error_cells.append(("AD", ad_raw))
    if _is_error_value(af_raw):
        error_cells.append(("AF", af_raw))
    if _is_error_value(ag_raw):
        error_cells.append(("AG", ag_raw))

    if error_cells:
        tx.has_error = True
        tx.unresolved_reasons.append(
            "数式エラーのため金額を確定できない: "
            + ", ".join(f"{col}={val}" for col, val in error_cells)
        )
        return tx

    ac = ac or 0.0
    ad = ad or 0.0
    af = af or 0.0

    # 金額(AC・AD・AF)は、日報自身の計算式が算出した値をそのまま正本として計上する。
    # 顧客名の有無や区分(D列)ラベルでは絞り込まない(ルール1〜3)。
    # 「体験(トライアル)して当日中に回数券を購入」のように、1行で新規売(AC)と
    # 既存売(AD)が同時に発生する取引が実在することを実データで確認済みであり、
    # これは矛盾ではなく正常な複合取引のため、金額を動かす対象にはしない。
    tx.new_amount = ac
    tx.existing_amount = ad
    tx.retail_amount = af

    flags = []
    if ac > 0 and d not in ("新規", None, ""):
        flags.append(f'区分(D列)="{d}"だが、コース名"{f}"から新規売上(AC={ac:g})が計算されている')
    if ad > 0 and d not in ("既存", None, ""):
        flags.append(f'区分(D列)="{d}"だが、コース名/回数券"{f}"から既存売上(AD={ad:g})が計算されている')
    if flags:
        tx.category_flag = " / ".join(flags)

    return tx


def _process_overflow_row(ws_v, sheet: str, row: int) -> TransactionRow | None:
    """51人目以降の予備行(58〜62行目)。新規/既存/物販の内訳式が無く、AGのみ存在する。"""
    ag_raw = ws_v[f"AG{row}"].value
    ag = _numeric_or_none(ag_raw)
    if ag is None or ag == 0:
        return None
    b = ws_v[f"B{row}"].value
    tx = TransactionRow(
        sheet=sheet, row=row, customer_name=b, staff=None,
        category_label=None, course=None, product=None,
    )
    tx.unresolved_amount = ag
    tx.unresolved_reasons.append(
        "予備行(51人目以降)。新規/既存/物販の内訳計算式が無いため、金額全体を未確定として扱う"
    )
    return tx


def extract_day(ws_v, sheet: str) -> DayResult:
    transactions: list[TransactionRow] = []
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        tx = _process_row(ws_v, sheet, row)
        if tx is not None:
            transactions.append(tx)
    for row in range(OVERFLOW_START_ROW, OVERFLOW_END_ROW + 1):
        tx = _process_overflow_row(ws_v, sheet, row)
        if tx is not None:
            transactions.append(tx)

    new_total = sum(t.new_amount for t in transactions)
    existing_total = sum(t.existing_amount for t in transactions)
    retail_total = sum(t.retail_amount for t in transactions)
    unresolved_total = sum(t.unresolved_amount for t in transactions)
    grand_total = new_total + existing_total + retail_total + unresolved_total

    file_ac54 = _numeric_or_none(ws_v[f"AC{TOTAL_ROW}"].value)
    file_ad54 = _numeric_or_none(ws_v[f"AD{TOTAL_ROW}"].value)
    file_af54 = _numeric_or_none(ws_v[f"AF{TOTAL_ROW}"].value)
    file_ag54 = _numeric_or_none(ws_v[f"AG{TOTAL_ROW}"].value)

    consistency_diff = None
    consistency_ok = True
    if file_ag54 is not None:
        consistency_diff = round(grand_total - file_ag54, 2)
        consistency_ok = abs(consistency_diff) < 1.0

    return DayResult(
        sheet=sheet,
        new_total=round(new_total, 2),
        existing_total=round(existing_total, 2),
        retail_total=round(retail_total, 2),
        unresolved_total=round(unresolved_total, 2),
        grand_total=round(grand_total, 2),
        transaction_count=len(transactions),
        unresolved_rows=[t for t in transactions if t.unresolved_amount != 0 or t.has_error],
        error_rows=[t for t in transactions if t.has_error],
        category_flagged_rows=[t for t in transactions if t.category_flag],
        file_ac54=file_ac54, file_ad54=file_ad54, file_af54=file_af54, file_ag54=file_ag54,
        consistency_ok=consistency_ok, consistency_diff=consistency_diff,
    )


def extract_all_days(path: Path, days: list[str] | None = None) -> dict[str, DayResult]:
    """全日別シート(または指定した日)を読み取り、DayResultの辞書を返す。書き込みは行わない。"""
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb_v = openpyxl.load_workbook(path, data_only=True, read_only=False)
    try:
        target_days = days if days is not None else DAY_SHEETS
        results = {}
        for day in target_days:
            if day not in wb_v.sheetnames:
                continue
            ws_v = wb_v[day]
            results[day] = extract_day(ws_v, day)
        return results
    finally:
        wb_v.close()


def extract_monthly_report_summary(path: Path) -> dict:
    """月報集計シートの主要セル(検算用)を読み取る。"""
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb_v = openpyxl.load_workbook(path, data_only=True, read_only=False)
    try:
        if "月報集計" not in wb_v.sheetnames:
            return {}
        ws = wb_v["月報集計"]
        return {
            "実売目標_I4": _numeric_or_none(ws["I4"].value),
            "実売実績_I5": _numeric_or_none(ws["I5"].value),
            "内回数券等_I6": _numeric_or_none(ws["I6"].value),
            "内物販_I7": _numeric_or_none(ws["I7"].value),
            "施術目標_I9": _numeric_or_none(ws["I9"].value),
            "施術実績_I10": _numeric_or_none(ws["I10"].value),
            "新規客数_I12": _numeric_or_none(ws["I12"].value),
        }
    finally:
        wb_v.close()
