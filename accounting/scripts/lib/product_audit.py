"""物販取引監査ロジック(読み取り専用)。

accounting/docs/product-audit-spec.md の正式ルールに基づく。
- 価格・原価は「取引日時点で有効だった履歴」を使用する(現在値で過去を再計算しない)。
- 価格履歴が無い場合は推測せず、G(価格履歴不足)またはH(商品不明)として扱う。
- 異常があっても「不正」「改ざん」とは判定しない。あくまで規定価格・原価・入力内容との
  不一致として抽出し、現場確認対象とする。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import yaml

# 丸め差として許容する誤差(円)。実際の丸めルールが確認できるまでの暫定値。
# これより大きい差は C(価格不一致)/D(スタッフ価格不一致)として扱う。
ROUNDING_TOLERANCE_YEN = 3.0

# 「異常値引き」として要確認フラグを立てる値引率の暫定しきい値(値引額/値引前金額)。
# 実際の値引運用ルールが確認できるまでの暫定値であり、これを超えたこと自体を
# 不正・改ざんとは判定しない(現場確認対象として抽出するのみ)。
LARGE_DISCOUNT_RATIO = 0.5

# 判定区分(A〜I、および売上金額自体が数式エラーで不明なX)から、4段階の重大度への
# 既定マッピング。個別取引でフラグ(flags)が立った場合は、この既定値より重大度を
# 下げない方向にのみ調整する(見逃しを避けるため)。
SEVERITY_BY_CLASSIFICATION = {
    "A": "正常",
    "B": "注意",
    "C": "要確認",
    "D": "要確認",
    "E": "注意",
    "F": "要確認",
    "G": "要確認",
    "H": "要確認",
    "I": "要確認",
    "X": "重大エラー",  # 物販売上(AF列)自体が数式エラーで金額不明
}

_SEVERITY_RANK = {"正常": 0, "注意": 1, "要確認": 2, "重大エラー": 3}


def _escalate_severity(base: str, *candidates: str) -> str:
    """重大度を候補の中で最も高いものへ引き上げる(下げることはしない)。"""
    best = base
    for c in candidates:
        if _SEVERITY_RANK.get(c, -1) > _SEVERITY_RANK.get(best, -1):
            best = c
    return best


def _normalize_name(name: str | None) -> str:
    if not name:
        return ""
    return str(name).replace("　", "").replace(" ", "").strip()


def load_price_history(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    records_by_product: dict[str, list[dict]] = {}
    for rec in doc.get("products", []):
        records_by_product.setdefault(rec["product_name"], []).append(rec)
    return records_by_product


def load_staff_price_rules(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    rules_by_product: dict[str, list[dict]] = {}
    for rec in doc.get("staff_price_rules", []) or []:
        rules_by_product.setdefault(rec["product_name"], []).append(rec)
    return rules_by_product


def _find_effective_record(records: list[dict], date_key: str) -> dict | None:
    target = dt.date.fromisoformat(date_key)
    for rec in records:
        start = dt.date.fromisoformat(rec["effective_start_date"])
        end_raw = rec.get("effective_end_date")
        end = dt.date.fromisoformat(end_raw) if end_raw else dt.date.max
        if start <= target <= end:
            return rec
    return None


def classify_purchaser(customer_name: str | None, staff_names: list[str]) -> str:
    """来店者区分(一般顧客/スタッフ/不明)を判定する。顧客名がスタッフ名で始まる場合を
    自己購入とみなす(フリガナが氏名の直後に続く表記のため前方一致で判定)。"""
    norm = _normalize_name(customer_name)
    if not norm:
        return "不明"
    for staff in staff_names:
        staff_norm = _normalize_name(staff)
        if staff_norm and norm.startswith(staff_norm):
            return "スタッフ"
    return "一般顧客"


UNKNOWN_PRODUCT_LABELS = (None, "", "無")


def determine_primary_bucket(purchaser_type: str, product_known: bool) -> str:
    """全取引を排他的に分類する集計バケット。

    2026-09-11の調査で判明した通り、「購入者区分」と「商品が価格履歴にあるか」は
    別々の軸であり、両方を無条件に足し合わせると同じ取引が二重に数えられたり
    (または漏れたり)する。監査サマリでは必ずこの5分類のいずれか1つにだけ
    振り分け、5分類の合計が物販総売上(AF列合計)と一致するようにする。
    商品不明を購入者区分より優先する(商品が特定できない取引は、まず商品不明として
    扱い、一般顧客/スタッフ内訳には二重計上しない)。
    """
    if not product_known:
        return "商品不明"
    if purchaser_type == "一般顧客":
        return "一般顧客"
    if purchaser_type == "スタッフ":
        return "スタッフ"
    if purchaser_type == "不明":
        return "購入者区分不明"
    return "その他"  # 到達しない想定のセーフティネット


@dataclass
class ProductTransaction:
    date: str
    sheet_row: int
    customer_name: str | None
    purchaser_type: str
    primary_bucket: str  # 一般顧客/スタッフ/購入者区分不明/商品不明/その他(排他的・合計が総売上と一致)
    product_name: str
    quantity: float | None
    regular_price_incl_tax_expected: float | None
    staff_price_incl_tax_expected: float | None
    actual_price_incl_tax: float  # P列(値引前の税込金額)
    discount_incl_tax: float      # T列(値引額)
    net_price_incl_tax: float     # W列(値引後の税込金額。社割等の値引はここで反映されている)
    tax_excl_revenue: float       # AF列(税抜売上。日報の実売をそのまま使用。売上の正本)
    cost_excl_tax_unit: float | None
    cost_excl_tax_total: float | None
    gross_profit: float | None
    gross_margin: float | None
    classification: str  # A〜I(またはX:売上金額自体が数式エラーで不明)
    classification_detail: str
    profit_confirmed: bool
    severity: str = "正常"          # 正常/注意/要確認/重大エラー(4段階)
    flags: list[str] = field(default_factory=list)  # 追加の要確認事由(複数可)


def _large_discount_flag(gross_incl_tax, discount_incl_tax) -> str | None:
    """値引額が値引前金額に対して暫定しきい値を超える場合、要確認フラグ文字列を返す。"""
    if not discount_incl_tax or not gross_incl_tax or gross_incl_tax <= 0:
        return None
    ratio = discount_incl_tax / gross_incl_tax
    if ratio > LARGE_DISCOUNT_RATIO:
        return f"異常値引き(値引率{ratio:.0%}、暫定しきい値{LARGE_DISCOUNT_RATIO:.0%}超)"
    return None


def _staff_price_for_general_flag(
    *, purchaser_type: str, product_name: str, date: str, qty: float,
    gross_incl_tax: float, net_incl_tax: float, staff_price_rules: dict,
) -> str | None:
    """一般顧客・購入者区分不明の取引が、実は社割(スタッフ)価格と一致していないかを確認する。

    一致していても自動的に不正とはせず、「要現場確認」の追加フラグとして抽出するのみ。
    """
    if purchaser_type not in ("一般顧客", "不明"):
        return None
    staff_records = staff_price_rules.get(product_name)
    if not staff_records:
        return None
    staff_rec = _find_effective_record(staff_records, date)
    if staff_rec is None or staff_rec.get("rule_type") != "explicit_staff_price_incl_tax":
        return None
    expected_staff_total = staff_rec["explicit_staff_price_incl_tax"] * qty
    if abs(gross_incl_tax - expected_staff_total) < 0.01 or abs(net_incl_tax - expected_staff_total) < 0.01:
        return f"一般顧客/区分不明だが実売価格が社割(スタッフ)価格({staff_rec['explicit_staff_price_incl_tax']}円)と一致"
    return None


def audit_transaction(
    *, date: str, row: int, customer_name, staff_col, category_label,
    product_name, quantity, gross_incl_tax, discount_incl_tax, net_incl_tax,
    tax_excl_revenue, price_history: dict, staff_price_rules: dict, staff_names: list[str],
    revenue_error: bool = False,
) -> ProductTransaction:
    purchaser_type = classify_purchaser(customer_name, staff_names)
    qty_for_checks = quantity if quantity else 1

    if revenue_error:
        # 物販売上(AF列)自体が数式エラー(#N/A等)で、金額を読み取れない場合。
        # 売上を0円と断定せず、金額不明の重大エラーとして分離する(売上を落とさないため、
        # 集計スクリプト側で「金額不明」取引として別掲し、正常売上合計には含めない)。
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, primary_bucket=determine_primary_bucket(purchaser_type, False),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="X",
            classification_detail="物販売上(AF列)が数式エラーのため金額不明。金額を0円と断定せず要確認とする。",
            profit_confirmed=False,
            severity=SEVERITY_BY_CLASSIFICATION["X"],
            flags=["売上金額が数式エラーで不明"],
        )

    is_unknown_label = product_name in UNKNOWN_PRODUCT_LABELS

    records = None if is_unknown_label else price_history.get(product_name)
    if records is None:
        detail = (
            "商品名が「無」のまま(未選択)で物販売上が計上されている"
            if is_unknown_label else
            "商品マスター(価格履歴)に存在しない商品名(表記ゆれの可能性も含め要確認)"
        )
        flags = []
        ld = _large_discount_flag(gross_incl_tax, discount_incl_tax)
        if ld:
            flags.append(ld)
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, primary_bucket=determine_primary_bucket(purchaser_type, False),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="H", classification_detail=detail,
            profit_confirmed=False,
            severity=_escalate_severity(SEVERITY_BY_CLASSIFICATION["H"], "要確認" if flags else "正常"),
            flags=flags,
        )

    rec = _find_effective_record(records, date)
    if rec is None:
        flags = []
        ld = _large_discount_flag(gross_incl_tax, discount_incl_tax)
        if ld:
            flags.append(ld)
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, primary_bucket=determine_primary_bucket(purchaser_type, True),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="G",
            classification_detail=f"取引日({date})を含む価格履歴レコードが無い",
            profit_confirmed=False,
            severity=_escalate_severity(SEVERITY_BY_CLASSIFICATION["G"], "要確認" if flags else "正常"),
            flags=flags,
        )

    regular_price = rec["regular_price_incl_tax"]
    cost_unit = rec.get("cost_price_excl_tax")

    staff_price_expected = None
    staff_rule_detail = None
    if purchaser_type == "スタッフ":
        staff_records = staff_price_rules.get(product_name)
        staff_rec = _find_effective_record(staff_records, date) if staff_records else None
        if staff_rec is None:
            staff_rule_detail = "missing"
        elif staff_rec["rule_type"] == "explicit_staff_price_incl_tax":
            staff_price_expected = staff_rec["explicit_staff_price_incl_tax"]
        elif staff_rec["rule_type"] == "discount_percent":
            staff_rule_detail = "discount_percent_unrounded"
            staff_price_expected = round(regular_price * (1 - staff_rec["discount_percent"]))

    qty = quantity if quantity else 1
    expected_total_regular = regular_price * qty if regular_price is not None else None
    expected_total_staff = staff_price_expected * qty if staff_price_expected is not None else None

    # 比較は「値引前(gross=P列)」「値引後(net=P列−T列)」の両方を試し、
    # どちらか一方が期待価格と一致すればそれを採用する。
    # 実データで2つの異なる運用が確認できたため:
    #   - 一般顧客の多くは、通常価格をそのままP列に入力し、T列は別の調整
    #     (本監査の対象外の値引・端数調整等)に使われている → gross側が一致する。
    #   - 一部のスタッフ購入は、通常価格をP列に入力したうえで、社割差額を
    #     T列に値引として記録している → net側が一致する。
    # 一方だけで比較すると、どちらかの正常な運用を誤って不一致判定してしまう
    # (2026-09-11に判明。最初はnet固定で比較し、一般顧客の正常取引を大量に
    # 誤検知したため、この「両方試す」方式に修正した)。
    def _best_match(expected):
        if expected is None:
            return None, None, None
        diff_gross = gross_incl_tax - expected
        diff_net = net_incl_tax - expected
        if abs(diff_gross) <= abs(diff_net):
            return diff_gross, "値引前(gross)", diff_net
        return diff_net, "値引後(net)", diff_gross

    if cost_unit is None:
        classification = "F"
        detail = "価格履歴レコードはあるが仕入原価が未登録"
        cost_total = None
        profit = None
        margin = None
        profit_confirmed = False
    else:
        cost_total = cost_unit * qty
        profit = round(tax_excl_revenue - cost_total, 2)
        margin = round(profit / tax_excl_revenue, 4) if tax_excl_revenue else None

        if purchaser_type == "スタッフ" and staff_rule_detail == "missing":
            classification = "G"
            detail = "スタッフ購入だが、該当商品・取引日の社割価格履歴が無い"
            profit_confirmed = False
        elif purchaser_type == "スタッフ" and staff_rule_detail == "discount_percent_unrounded":
            classification = "I"
            detail = "割合ベースの社割価格(端数処理ルール未確定)のため現場確認が必要"
            profit_confirmed = False
        elif purchaser_type == "スタッフ" and expected_total_staff is not None:
            diff, which, _ = _best_match(expected_total_staff)
            if abs(diff) < 0.01:
                classification, detail = "A", f"正常({which}の実売価格がスタッフ価格と一致)"
                profit_confirmed = True
            elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
                classification, detail = "B", f"丸め差の可能性({which}との差額{diff:+.2f}円、暫定許容範囲内)"
                profit_confirmed = True
            else:
                classification, detail = "D", f"スタッフ価格と不一致({which}との差額{diff:+.2f}円)"
                profit_confirmed = True
        else:
            # 一般顧客、または購入者区分「不明」(通常価格で判定)
            diff, which, _ = _best_match(expected_total_regular)
            if diff is None:
                classification, detail = "G", "通常価格が価格履歴に無い"
                profit_confirmed = False
            elif abs(diff) < 0.01:
                classification, detail = "A", f"正常({which}の実売価格が通常価格と一致)"
                profit_confirmed = True
            elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
                classification, detail = "B", f"丸め差の可能性({which}との差額{diff:+.2f}円、暫定許容範囲内)"
                profit_confirmed = True
            else:
                classification, detail = "C", f"通常価格と不一致({which}との差額{diff:+.2f}円)"
                profit_confirmed = True

    if classification in ("F", "G", "H"):
        profit, margin, profit_confirmed = None, None, False

    # 追加の要確認フラグ(判定区分A〜Iとは別軸。いずれも「不正確定」ではなく現場確認対象)。
    flags: list[str] = []
    if cost_unit is not None and cost_unit == 0:
        flags.append("原価0円(価格履歴上0円で登録されている)")
    ld = _large_discount_flag(gross_incl_tax, discount_incl_tax)
    if ld:
        flags.append(ld)
    staff_flag = _staff_price_for_general_flag(
        purchaser_type=purchaser_type, product_name=product_name, date=date, qty=qty,
        gross_incl_tax=gross_incl_tax, net_incl_tax=net_incl_tax, staff_price_rules=staff_price_rules,
    )
    if staff_flag:
        flags.append(staff_flag)

    severity = _escalate_severity(
        SEVERITY_BY_CLASSIFICATION[classification], "要確認" if flags else "正常",
    )

    return ProductTransaction(
        date=date, sheet_row=row, customer_name=customer_name,
        purchaser_type=purchaser_type, primary_bucket=determine_primary_bucket(purchaser_type, True),
        product_name=product_name, quantity=quantity,
        regular_price_incl_tax_expected=regular_price,
        staff_price_incl_tax_expected=staff_price_expected,
        actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
        net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
        cost_excl_tax_unit=cost_unit, cost_excl_tax_total=cost_total,
        gross_profit=profit, gross_margin=margin,
        classification=classification, classification_detail=detail,
        profit_confirmed=profit_confirmed,
        severity=severity, flags=flags,
    )


def parse_inventory_ledger(wb) -> list[dict]:
    """在庫帳シートを読み取り専用で解析し、商品ごとの当月販売数量・原価ベース販売金額を
    抽出する(数量と在庫減少の不整合、販売数量×原価と販売原価の不整合の突合に使用)。

    在庫帳は商品ごとに2行1組(1行目=入荷数行、2行目=販売数行)で構成されている。
    - C列:商品名(基本情報シートを参照するVLOOKUP式。キャッシュ済みの値を読む)
    - AM列(販売数行):当月の日別販売数量の小計(店舗スタッフによる手入力の日別列の合計)
    - AU列(入荷数行):原価(基本情報シートを参照するVLOOKUP式)
    - BA列(入荷数行):当月販売金額(原価ベース。=AM(販売数行)×AU)
    在庫帳自体はこのシート内で完結しており、外部ブック参照は期首在庫(F列)のみで、
    ここで使う列(C・AM・AU・BA)はいずれも外部参照を経由しない(product-audit-spec.md §7参照)。
    """
    ws = wb["在庫帳"]
    blocks = []
    row = 4
    while row + 1 <= ws.max_row:
        label1 = ws.cell(row=row, column=7).value      # G列
        label2 = ws.cell(row=row + 1, column=7).value
        if label1 != "入荷数" or label2 != "販売数":
            row += 1
            continue
        name = ws.cell(row=row, column=3).value         # C列
        qty_sold = ws.cell(row=row + 1, column=39).value  # AM列(販売数行)
        unit_cost = ws.cell(row=row, column=47).value      # AU列(入荷数行)
        cost_value = ws.cell(row=row, column=53).value      # BA列(入荷数行)
        blocks.append({
            "row": row,
            "product_name": str(name).strip() if name not in (None, "") else None,
            "monthly_qty_sold": qty_sold if isinstance(qty_sold, (int, float)) else None,
            "unit_cost_excl_tax": unit_cost if isinstance(unit_cost, (int, float)) else None,
            "monthly_cost_value": cost_value if isinstance(cost_value, (int, float)) else None,
        })
        row += 2
    return blocks
