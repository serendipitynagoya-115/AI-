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


@dataclass
class ProductTransaction:
    date: str
    sheet_row: int
    customer_name: str | None
    purchaser_type: str
    product_name: str
    quantity: float | None
    regular_price_incl_tax_expected: float | None
    staff_price_incl_tax_expected: float | None
    actual_price_incl_tax: float  # P列(値引前の税込金額)
    discount_incl_tax: float      # T列(値引額)
    net_price_incl_tax: float     # W列(値引後の税込金額)
    tax_excl_revenue: float       # AF列(税抜売上。日報の実売をそのまま使用)
    cost_excl_tax_unit: float | None
    cost_excl_tax_total: float | None
    gross_profit: float | None
    gross_margin: float | None
    classification: str  # A〜I
    classification_detail: str
    profit_confirmed: bool


def audit_transaction(
    *, date: str, row: int, customer_name, staff_col, category_label,
    product_name, quantity, gross_incl_tax, discount_incl_tax, net_incl_tax,
    tax_excl_revenue, price_history: dict, staff_price_rules: dict, staff_names: list[str],
) -> ProductTransaction:
    purchaser_type = classify_purchaser(customer_name, staff_names)

    records = price_history.get(product_name)
    if records is None:
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="H", classification_detail="商品マスター(価格履歴)に存在しない商品名",
            profit_confirmed=False,
        )

    rec = _find_effective_record(records, date)
    if rec is None:
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="G",
            classification_detail=f"取引日({date})を含む価格履歴レコードが無い",
            profit_confirmed=False,
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
            diff = gross_incl_tax - expected_total_staff
            if abs(diff) < 0.01:
                classification, detail = "A", "正常(スタッフ価格と一致)"
                profit_confirmed = True
            elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
                classification, detail = "B", f"丸め差の可能性(差額{diff:+.2f}円、暫定許容範囲内)"
                profit_confirmed = True
            else:
                classification, detail = "D", f"スタッフ価格と不一致(差額{diff:+.2f}円)"
                profit_confirmed = True
        else:
            # 一般顧客、または購入者区分「不明」(通常価格で判定)
            diff = gross_incl_tax - expected_total_regular if expected_total_regular is not None else None
            if diff is None:
                classification, detail = "G", "通常価格が価格履歴に無い"
                profit_confirmed = False
            elif abs(diff) < 0.01:
                classification, detail = "A", "正常(通常価格と一致)"
                profit_confirmed = True
            elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
                classification, detail = "B", f"丸め差の可能性(差額{diff:+.2f}円、暫定許容範囲内)"
                profit_confirmed = True
            else:
                classification, detail = "C", f"通常価格と不一致(差額{diff:+.2f}円)"
                profit_confirmed = True

    if classification in ("F", "G", "H"):
        profit, margin, profit_confirmed = None, None, False

    return ProductTransaction(
        date=date, sheet_row=row, customer_name=customer_name,
        purchaser_type=purchaser_type, product_name=product_name, quantity=quantity,
        regular_price_incl_tax_expected=regular_price,
        staff_price_incl_tax_expected=staff_price_expected,
        actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
        net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
        cost_excl_tax_unit=cost_unit, cost_excl_tax_total=cost_total,
        gross_profit=profit, gross_margin=margin,
        classification=classification, classification_detail=detail,
        profit_confirmed=profit_confirmed,
    )
