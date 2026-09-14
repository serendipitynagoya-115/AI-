"""物販取引監査ロジック(読み取り専用)。

accounting/docs/product-audit-spec.md の正式ルールに基づく。
- 価格・原価は「取引日時点で有効だった履歴」を使用する(現在値で過去を再計算しない)。
- 価格履歴が無い場合は推測せず、G(価格履歴不足)またはH(商品不明)として扱う。
- 異常があっても「不正」「改ざん」とは判定しない。あくまで規定価格・原価・入力内容との
  不一致として抽出し、現場確認対象とする。
- 「価格監査」(実売価格が規定通りか)と「原価監査」(仕入原価が判明しているか)は
  別軸として管理する。スタッフ社割価格が未登録でも、商品の原価自体が価格履歴に
  登録されていれば、原価は確定として扱い、売上原価・粗利益の集計に含める
  (2026-09-13確定。product-audit-spec.md §9参照)。
"""
from __future__ import annotations

import datetime as dt
import itertools
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

# 判定区分(A〜I、既知差異のK、売上金額自体が数式エラーで不明なX)から、
# 5段階の重大度への既定マッピング。個別取引でフラグ(flags)が立った場合は、
# この既定値より重大度を下げない方向にのみ調整する(見逃しを避けるため)。
# 「既知差異(K)」は原因が特定済みの表示価格変動であり、要確認・重大エラーには含めない。
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
    "K": "既知差異",  # 別期間の価格履歴と一致する既知の表示価格変動(product-audit-spec.md §8)
    "S": "正常",      # 売上シェア(確定)。1商品を複数スタッフで分担入力(product-audit-spec.md §11)
    "N": "正常",      # 社内例外取引(通常非販売の備品等を原価のまま社内購入。product-audit-spec.md §18)
    "V": "注意",      # 価格可変・非定番商品(セット料金等)。原価未確認の場合の既定値(product-audit-spec.md §24)
    "X": "重大エラー",  # 物販売上(AF列)自体が数式エラーで金額不明
}

# 価格監査(A軸)の状態。原価監査(B軸)とは独立して管理する(2026-09-13確定)。
PRICE_STATUS_BY_CLASSIFICATION = {
    "A": "価格確定",
    "B": "価格確定",
    "C": "価格要確認",
    "D": "価格要確認",
    "E": "価格要確認",
    "F": "価格要確認",
    "G": "価格要確認",
    "H": "価格要確認",
    "I": "価格要確認",
    "K": "既知差異",
    "S": "価格確定",
    "N": "価格確定",
    "V": "価格確定",  # 価格可変・非定番商品は「実際に記録された金額」自体を確定値として扱う
    "X": "価格要確認",
}

# 売上シェア(1商品を複数スタッフで分担入力)の金額一致を判定する許容誤差(円)。
SHARE_AMOUNT_TOLERANCE_YEN = 1.0

# 日報ExcelのI列に記載される、売上シェア入力であることの明示マーカー文字列。
# I列は単純な「備考」欄ではなく、「既存 単発・回数券」(施術チケットの種別)を記録する欄だが、
# 特殊運用として「売上シェア」という文字列が入力される場合がある(2026-09-15オーナー確認、
# product-audit-spec.md §13)。2026-09-15、守山8月の実データで、シェア分担側の行(「無」行)に
# I列「売上シェア」の記載があることを確認した(product-audit-spec.md §11)。これが存在する場合は、
# 金額の一致・近接行等の推測より優先する一次証跡として扱う(I列自体の意味を「売上シェア専用欄」と
# 再定義するものではない)。
SHARE_MARKER_LABEL = "売上シェア"

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


def load_staff_aliases(path: Path) -> dict:
    """スタッフ・社内購入者の別名(alias)マスターを読み込む。

    戻り値は正規化済みの別名(alias)をキーに、{"formal_name":..., "note":...} を
    値に持つ辞書。表記ゆれ(ひらがな表記・オーナーの通称等)を、購入者区分の
    判定(classify_purchaser)で吸収するために使う。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    aliases: dict[str, dict] = {}
    for rec in doc.get("aliases", []) or []:
        key = _normalize_name(rec["alias"])
        if key:
            aliases[key] = {"formal_name": rec["formal_name"], "note": rec.get("note") or ""}
    return aliases


def load_confirmed_purchaser_blocks(path: Path) -> list[dict]:
    """確定済み「同一購入者・連続購入ブロック」マスターを読み込む(2026-09-15確定、
    product-audit-spec.md §12参照)。オーナーが日報原本を確認して確定した範囲だけを
    保持し、推測で範囲を広げない。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    return doc.get("blocks", []) or []


def apply_confirmed_purchaser_blocks(
    rows: list[dict], *, store_id: str, year_month: str, date_key: str, blocks: list[dict],
) -> None:
    """確定済みブロック内で顧客名(customer_name)が空欄の行に、起点行の購入者名を適用する
    (in-place)。既に顧客名が入力されている行は上書きしない。範囲外の空欄行には一切影響しない。
    """
    matching = [
        b for b in blocks
        if b["store_id"] == store_id and b["year_month"] == year_month and b["date"] == date_key
    ]
    if not matching:
        return
    for r in rows:
        for b in matching:
            if b["start_row"] <= r["row"] <= b["end_row"] and not r.get("customer_name"):
                r["customer_name"] = b["purchaser_name"]
                r["purchaser_name_note"] = (
                    f"顧客名は空欄のため、確定済み連続購入ブロック({date_key} {b['start_row']}〜"
                    f"{b['end_row']}行、起点の購入者「{b['purchaser_name']}」)により顧客名を適用"
                    "(product-audit-spec.md §12)。"
                )
                break


def load_confirmed_quantity_corrections(path: Path) -> dict[tuple, dict]:
    """確定済み数量修正マスターを読み込む(2026-09-15確定、product-audit-spec.md §19参照)。
    キーは(store_id, year_month, date, row)。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    result = {}
    for rec in doc.get("corrections", []) or []:
        key = (rec["store_id"], rec["year_month"], rec["date"], rec["row"])
        result[key] = rec
    return result


def apply_confirmed_quantity_corrections(
    rows: list[dict], *, store_id: str, year_month: str, date_key: str, corrections: dict[tuple, dict],
) -> None:
    """確定済み数量修正を、監査前の生データ行(rows)に適用する(in-place)。

    日報K列の数量が、現場確認により実際の販売数量と異なると確定した場合に使う
    (例:売上シェア運用によりK列が実数量と異なる値になっているケース)。数量を
    修正することで、原価・価格一致判定を正しい数量ベースで計算し直せる
    (元Excelは変更しない。product-audit-spec.md §19参照)。
    """
    for r in rows:
        rec = corrections.get((store_id, year_month, date_key, r["row"]))
        if rec is None:
            continue
        r["quantity_correction_note"] = (
            f"日報K列は数量{r['quantity']}だが、現場確認により実際の販売数量は"
            f"{rec['corrected_quantity']}と確定({rec['note'].strip()})"
        )
        r["quantity"] = rec["corrected_quantity"]


def load_confirmed_category_reclassifications(path: Path) -> dict[tuple, dict]:
    """物販監査の対象から除外する、確定済みの区分振替マスターを読み込む(2026-09-15確定、
    product-audit-spec.md §14参照)。キーは(store_id, year_month, date, row)。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    result = {}
    for rec in doc.get("reclassifications", []) or []:
        key = (rec["store_id"], rec["year_month"], rec["date"], rec["row"])
        result[key] = rec
    return result


def load_confirmed_status_overrides(path: Path) -> dict[tuple, dict]:
    """確定済みステータス上書きマスターを読み込む(2026-09-15確定、
    product-audit-spec.md §15参照)。キーは(store_id, year_month, date, row)。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    result = {}
    for rec in doc.get("overrides", []) or []:
        key = (rec["store_id"], rec["year_month"], rec["date"], rec["row"])
        result[key] = rec
    return result


def apply_confirmed_status_overrides(
    transactions: list[ProductTransaction], *, store_id: str, year_month: str, overrides: dict[tuple, dict],
) -> int:
    """確定済みステータス上書きを取引に反映する(in-place)。適用件数を返す。

    `new_severity`は必須。`new_classification`・`new_price_status`・`new_transaction_type`は
    任意で、現場確認により価格異常自体が解消したと確定した場合にのみ指定する
    (例:「無」ではなく実際には売上シェアだったと現場確認できたが、対応する行が
    ファイル内に見当たらず自動検出できないケース。product-audit-spec.md §19)。
    指定が無ければ従来通りseverityのみを変更する(classification・原価判定は変更しない)。
    `clear_flags: true`が指定されていれば、追加の要確認フラグ(flags)もクリアする。

    `mark_cost_unconfirmed: true`が指定されていれば、原価監査(B軸)をcost_confirmed=False
    へ上書きする(2026-09-16確定、product-audit-spec.md §26参照)。商品マスターに
    cost_price_excl_tax=0が登録されているだけで、実際の原価がまだ特定できていない
    非定番商品(§24のvariable_price_non_standardと同種だが、店舗×月固有のconfirmed設定
    としてオーナー確認した個別取引に使う)について、売上確定と原価確定が独立である
    という原則(§25)を守るため、原価0円を確定原価として扱わないようにする。
    売上(tax_excl_revenue)は変更しない。
    """
    applied = 0
    for t in transactions:
        key = (store_id, year_month, t.date, t.sheet_row)
        rec = overrides.get(key)
        if rec is None:
            continue
        t.severity = rec["new_severity"]
        if rec.get("new_classification"):
            t.classification = rec["new_classification"]
        if rec.get("new_price_status"):
            t.price_status = rec["new_price_status"]
        if rec.get("new_transaction_type"):
            t.transaction_type = rec["new_transaction_type"]
        if rec.get("clear_flags"):
            t.flags = []
        if rec.get("mark_cost_unconfirmed"):
            t.cost_confirmed = False
            t.cost_excl_tax_unit = None
            t.cost_excl_tax_total = None
            t.gross_profit = None
            t.gross_margin = None
        t.classification_detail = t.classification_detail + " / " + rec["note"].strip()
        applied += 1
    return applied


def load_confirmed_exception_transactions(path: Path) -> dict[tuple, dict]:
    """確定済み社内例外取引マスターを読み込む(2026-09-15確定、product-audit-spec.md §18参照)。
    キーは(store_id, year_month, date, row)。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    result = {}
    for rec in doc.get("exceptions", []) or []:
        key = (rec["store_id"], rec["year_month"], rec["date"], rec["row"])
        result[key] = rec
    return result


def apply_confirmed_exception_transactions(
    transactions: list[ProductTransaction], *, store_id: str, year_month: str, exceptions: dict[tuple, dict],
) -> int:
    """確定済み社内例外取引を反映する(in-place)。適用件数を返す。

    通常は販売しない備品等を、原価のまま社内購入した例外取引を扱う(product-audit-spec.md
    §18)。商品マスター・原価マスターに存在しないことを異常として扱わず、
    classification="N"(社内例外取引)・severity="正常"とし、原価=税抜売上(粗利益0円)
    として確定する。通常の商品マスターには登録しない。
    """
    applied = 0
    for t in transactions:
        key = (store_id, year_month, t.date, t.sheet_row)
        rec = exceptions.get(key)
        if rec is None:
            continue
        t.product_name = rec["product_name"]
        t.product_status = "商品特定済み"
        t.purchaser_type = rec["purchaser_type_override"]
        t.cost_excl_tax_unit = round(t.tax_excl_revenue / t.quantity, 2) if t.quantity else t.tax_excl_revenue
        t.cost_excl_tax_total = t.tax_excl_revenue
        t.cost_confirmed = True
        t.gross_profit = 0.0
        t.gross_margin = 0.0
        t.classification = "N"
        t.classification_detail = f"社内例外取引({rec['exception_type']}): {rec['note'].strip()}"
        t.price_status = PRICE_STATUS_BY_CLASSIFICATION["N"]
        t.severity = SEVERITY_BY_CLASSIFICATION["N"]
        t.flags = []
        t.transaction_type = "internal_exception_sale"
        applied += 1
    return applied


def _find_effective_record(records: list[dict], date_key: str) -> dict | None:
    target = dt.date.fromisoformat(date_key)
    for rec in records:
        start = dt.date.fromisoformat(rec["effective_start_date"])
        end_raw = rec.get("effective_end_date")
        end = dt.date.fromisoformat(end_raw) if end_raw else dt.date.max
        if start <= target <= end:
            return rec
    return None


def classify_purchaser(customer_name: str | None, staff_names: list[str], aliases: dict | None = None) -> str:
    """来店者区分(一般顧客/スタッフ/不明)を判定する。

    区分は「スタッフ」「社内購入」「一般顧客」「購入者不明」の4種類(2026-09-13確定)。
    - 顧客名が空欄の場合は「購入者不明」とする。
    - 顧客名がスタッフ名簿(基本情報シート)の前方一致に該当する場合は「スタッフ」とする
      (フリガナが氏名の直後に続く表記のため前方一致で判定)。別名(alias)経由で
      解決した正式氏名がスタッフ名簿に該当する場合も「スタッフ」とする
      (表記ゆれ吸収のため。例:こたぎりけいこ→小田切敬子)。
    - 顧客名が別名(alias)マスターに一致するが、解決した正式氏名がスタッフ名簿には
      無い場合は「社内購入」とする(例:オーナー→辻正裕)。
    - いずれにも該当しない場合は「一般顧客」とする。
    """
    norm = _normalize_name(customer_name)
    if not norm:
        return "購入者不明"
    if aliases and norm in aliases:
        formal_norm = _normalize_name(aliases[norm]["formal_name"])
        for staff in staff_names:
            staff_norm = _normalize_name(staff)
            if staff_norm and formal_norm.startswith(staff_norm):
                return "スタッフ"
        return "社内購入"
    for staff in staff_names:
        staff_norm = _normalize_name(staff)
        if staff_norm and norm.startswith(staff_norm):
            return "スタッフ"
    return "一般顧客"


def resolve_purchaser_alias(customer_name: str | None, aliases: dict | None) -> dict | None:
    """顧客名が別名(alias)マスターに一致する場合、その正式氏名・備考を返す(無ければNone)。"""
    if not aliases:
        return None
    norm = _normalize_name(customer_name)
    return aliases.get(norm)


UNKNOWN_PRODUCT_LABELS = (None, "", "無")


def determine_product_status(product_known: bool) -> str:
    """商品特定状態(商品特定済み/商品不明)。購入者区分とは独立した別軸(2026-09-14確定)。

    2026-09-11の調査時点では「購入者区分」と「商品が価格履歴にあるか」を1つの
    排他的な集計バケット(primary_bucket)にまとめていたが、これだと商品不明の
    取引が購入者区分別の集計から機械的に除外されてしまい、「一般顧客なのに
    商品不明」のような実態を表せなかった。2026-09-14、購入者区分(4分類:
    スタッフ/社内購入/一般顧客/購入者不明)と商品特定状態(2分類:商品特定済み/
    商品不明)を、それぞれ独立に「合計が物販総売上と一致する」軸として管理する
    方式に改めた。1つの取引が両方の属性(例:購入者=一般顧客、商品=商品不明)を
    同時に持てる。
    """
    return "商品特定済み" if product_known else "商品不明"


@dataclass
class ProductTransaction:
    date: str
    sheet_row: int
    customer_name: str | None
    purchaser_type: str  # スタッフ/社内購入/一般顧客/購入者不明(排他的・合計が総売上と一致)
    product_status: str  # 商品特定済み/商品不明(購入者区分とは独立した軸。排他的・合計が総売上と一致)
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
    classification: str  # A〜I、K(既知差異)、X(売上金額が数式エラーで不明)
    classification_detail: str
    cost_confirmed: bool   # 原価監査(B軸): 原価(仕入原価)自体が判明しているか。価格の一致とは独立。
    price_status: str = "価格要確認"  # 価格監査(A軸): 価格確定/価格要確認/既知差異
    severity: str = "正常"          # 正常/注意/既知差異/要確認/重大エラー
    flags: list[str] = field(default_factory=list)  # 追加の要確認事由(複数可)
    purchaser_alias_note: str = ""  # 別名(alias)経由でスタッフ/社内購入と判定した場合の備考
    purchaser_name_note: str = ""  # 確定済み連続購入ブロックにより顧客名を適用した場合の注記(§12)
    quantity_correction_note: str = ""  # 確定済み数量修正を適用した場合の注記(§19)
    # 売上シェア(1商品を複数スタッフで分担入力)関連(2026-09-15確定。product-audit-spec.md §11)。
    transaction_type: str = "normal_sale"  # normal_sale/shared_sale/shared_sale_candidate
    share_group_id: str | None = None
    share_count: int | None = None      # グループ内の行数(確定は2、候補は3以上)
    share_total: float | None = None    # グループ全体の合計金額(規定価格と一致する額)
    linked_product: str | None = None   # グループが表す実際の商品名
    linked_rows: list[int] = field(default_factory=list)  # グループを構成する全行番号
    share_marker_raw: str | None = None  # 日報I列(既存 単発・回数券欄)の生値。「売上シェア」の明示記載を保持する。
    # 監査調整額(audit_adjustment、税抜)。K判定のうち、取引日時点より後の期間の
    # 価格履歴と一致した(=商品マスターの後日更新により過去の表示が歴史的事実と
    # 異なっている)取引についてのみ算出する。単なる価格ルール違反・スタッフ価格差異
    # には適用しない(2026-09-16確定。product-audit-spec.md §28参照)。
    audit_adjustment_excl_tax: float = 0.0
    audit_adjustment_applicable: bool = False
    audit_adjustment_note: str = ""


def _has_share_marker(t: ProductTransaction) -> bool:
    """日報I列(既存 単発・回数券欄)に「売上シェア」の明示記載があるかどうかを返す。"""
    return bool(t.share_marker_raw) and str(t.share_marker_raw).strip() == SHARE_MARKER_LABEL


def _qty_note(qty: float, gross_incl_tax: float, net_incl_tax: float, which: str) -> str:
    """数量が2以上の価格不一致(C・D)について、単純な合計差額だけでは「1個分の売上が
    抜けている」ように誤読されやすいため、実質単価も併記する(2026-09-15確定。
    例:数量2・合計5,940円を「通常価格(1個分)との差額-5,940円」とだけ表示すると、
    1個分の売上が丸ごと無いように読めるが、実際は2個×2,970円という単価の問題である)。
    """
    if not qty or qty == 1:
        return ""
    amount = gross_incl_tax if which == "値引前(gross)" else net_incl_tax
    unit_price = round(amount / qty, 2)
    return f"、数量{qty:g}個(実質単価{unit_price:g}円/個)"


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
    スタッフ・社内購入(別名マスターで判定済み)には適用しない(既に社内購入者と
    確認済みのため、このフラグの対象外)。
    """
    if purchaser_type not in ("一般顧客", "購入者不明"):
        return None
    staff_records = staff_price_rules.get(product_name)
    if not staff_records:
        return None
    staff_rec = _find_effective_record(staff_records, date)
    if staff_rec is None or staff_rec.get("rule_type") != "explicit_staff_price_incl_tax":
        return None
    expected_staff_total = staff_rec["explicit_staff_price_incl_tax"] * qty
    if abs(gross_incl_tax - expected_staff_total) < 0.01 or abs(net_incl_tax - expected_staff_total) < 0.01:
        return f"一般顧客/購入者不明だが実売価格が社割(スタッフ)価格({staff_rec['explicit_staff_price_incl_tax']}円)と一致"
    return None


def _find_known_discrepancy(records: list[dict], current_rec: dict, qty: float, gross_incl_tax: float) -> dict | None:
    """実売価格(値引前)が、取引日時点以外の期間の通常価格履歴レコードと一致するかを確認する。

    2026-09-12、守山8月のマグネシウム(ドクターセレン)監査で、日報Excelの
    基本情報シートの商品マスターが、取得タイミングによっては後の期間の価格へ
    既に更新されており、過去の取引行の実売価格の入力(P列)にもその新価格が
    反映されていた事象が確認された(product-audit-spec.md §8参照)。
    これは価格自体の異常ではなく、日報Excelの商品マスター更新に起因する
    既知の表示価格変動であるため、通常の「C:価格不一致」等とは区別する。
    """
    for rec in records:
        if rec is current_rec:
            continue
        other_regular = rec.get("regular_price_incl_tax")
        if other_regular is None:
            continue
        if abs(gross_incl_tax - other_regular * qty) < 0.01:
            return rec
    return None


def audit_transaction(
    *, date: str, row: int, customer_name, staff_col, category_label,
    product_name, quantity, gross_incl_tax, discount_incl_tax, net_incl_tax,
    tax_excl_revenue, price_history: dict, staff_price_rules: dict, staff_names: list[str],
    staff_aliases: dict | None = None, revenue_error: bool = False, note_raw: str | None = None,
    purchaser_name_note: str = "", quantity_correction_note: str = "",
) -> ProductTransaction:
    purchaser_type = classify_purchaser(customer_name, staff_names, staff_aliases)
    alias_rec = resolve_purchaser_alias(customer_name, staff_aliases)
    alias_note = (
        f"別名マスターにより「{alias_rec['formal_name']}」と判定" + (f"({alias_rec['note']})" if alias_rec.get("note") else "")
        if alias_rec else ""
    )

    if revenue_error:
        # 物販売上(AF列)自体が数式エラー(#N/A等)で、金額を読み取れない場合。
        # 売上を0円と断定せず、金額不明の重大エラーとして分離する(売上を落とさないため、
        # 集計スクリプト側で「金額不明」取引として別掲し、正常売上合計には含めない)。
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, product_status=determine_product_status(False),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="X",
            classification_detail="物販売上(AF列)が数式エラーのため金額不明。金額を0円と断定せず要確認とする。",
            cost_confirmed=False,
            price_status=PRICE_STATUS_BY_CLASSIFICATION["X"],
            severity=SEVERITY_BY_CLASSIFICATION["X"],
            flags=["売上金額が数式エラーで不明"],
            purchaser_alias_note=alias_note,
            purchaser_name_note=purchaser_name_note,
            quantity_correction_note=quantity_correction_note,
            share_marker_raw=note_raw,
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
            purchaser_type=purchaser_type, product_status=determine_product_status(False),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="H", classification_detail=detail,
            cost_confirmed=False,
            price_status=PRICE_STATUS_BY_CLASSIFICATION["H"],
            severity=_escalate_severity(SEVERITY_BY_CLASSIFICATION["H"], "要確認" if flags else "正常"),
            flags=flags,
            purchaser_alias_note=alias_note,
            purchaser_name_note=purchaser_name_note,
            quantity_correction_note=quantity_correction_note,
            share_marker_raw=note_raw,
        )

    rec = _find_effective_record(records, date)
    if rec is None:
        flags = []
        ld = _large_discount_flag(gross_incl_tax, discount_incl_tax)
        if ld:
            flags.append(ld)
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, product_status=determine_product_status(True),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=None, staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=None, cost_excl_tax_total=None,
            gross_profit=None, gross_margin=None,
            classification="G",
            classification_detail=f"取引日({date})を含む価格履歴レコードが無い",
            cost_confirmed=False,
            price_status=PRICE_STATUS_BY_CLASSIFICATION["G"],
            severity=_escalate_severity(SEVERITY_BY_CLASSIFICATION["G"], "要確認" if flags else "正常"),
            flags=flags,
            purchaser_alias_note=alias_note,
            purchaser_name_note=purchaser_name_note,
            quantity_correction_note=quantity_correction_note,
            share_marker_raw=note_raw,
        )

    if rec.get("variable_price_non_standard"):
        # 価格可変・非定番商品(セット料金等)。商品種類・組み合わせが多数あり、頻繁に
        # 販売しない非定番商品(補正下着等)を通常の商品マスターへ全商品登録せず、
        # 日報に実際の販売金額を直接入力する運用(2026-09-15オーナー確認、
        # product-audit-spec.md §24参照)。商品マスターの価格(regular_price_incl_tax、
        # 多くは0円のプレースホルダー)との比較による価格異常判定(C・D)は行わず、
        # 実際に記録された金額をそのまま実績売上として100%計上する。
        # 「variable_price_non_standard」フラグを明示登録した商品名だけが対象で、
        # 他の0円プレースホルダー商品(その他店販等)には一切適用しない。
        cost_unit = rec.get("cost_price_excl_tax")
        qty = quantity if quantity else 1
        cost_confirmed = cost_unit is not None
        if cost_confirmed:
            cost_total = cost_unit * qty
            profit = round(tax_excl_revenue - cost_total, 2)
            margin = round(profit / tax_excl_revenue, 4) if tax_excl_revenue else None
        else:
            cost_total = None
            profit = None
            margin = None
        detail = (
            "価格可変・非定番商品(セット料金等)。商品マスターの価格との比較による"
            "価格異常判定は行わず、実際に記録された金額をそのまま実績売上として計上する"
            "(product-audit-spec.md §24参照)。"
        )
        detail += (
            "原価は商品マスター(在庫帳含む)からは特定できないため「原価未確認」とする"
            "(推測原価は適用しない)。" if not cost_confirmed else
            f"原価は商品マスターの登録値({cost_unit}円/個)を使用する。"
        )
        flags = []
        ld = _large_discount_flag(gross_incl_tax, discount_incl_tax)
        if ld:
            flags.append(ld)
        if quantity_correction_note:
            detail = f"{detail} / {quantity_correction_note}"
        return ProductTransaction(
            date=date, sheet_row=row, customer_name=customer_name,
            purchaser_type=purchaser_type, product_status=determine_product_status(True),
            product_name=product_name, quantity=quantity,
            regular_price_incl_tax_expected=rec.get("regular_price_incl_tax"),
            staff_price_incl_tax_expected=None,
            actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
            net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
            cost_excl_tax_unit=cost_unit, cost_excl_tax_total=cost_total,
            gross_profit=profit, gross_margin=margin,
            classification="V", classification_detail=detail,
            cost_confirmed=cost_confirmed,
            price_status=PRICE_STATUS_BY_CLASSIFICATION["V"],
            severity="正常" if cost_confirmed else SEVERITY_BY_CLASSIFICATION["V"],
            flags=flags,
            purchaser_alias_note=alias_note,
            purchaser_name_note=purchaser_name_note,
            quantity_correction_note=quantity_correction_note,
            share_marker_raw=note_raw,
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

    # 原価監査(B軸): 価格履歴レコードに仕入原価が登録されているかどうかだけで判定する。
    # 価格(実売価格)が規定通りかどうかとは独立に扱う(2026-09-13確定。
    # 従来はG判定(社割価格履歴不足)等の場合に原価まで一律未確定扱いにしていたが、
    # 原価自体は判明しているため、売上原価・粗利益の集計に含めるよう改めた)。
    cost_confirmed = cost_unit is not None
    if cost_confirmed:
        cost_total = cost_unit * qty
        profit = round(tax_excl_revenue - cost_total, 2)
        margin = round(profit / tax_excl_revenue, 4) if tax_excl_revenue else None
    else:
        cost_total = None
        profit = None
        margin = None

    # 価格監査(A軸): 実売価格が取引日時点の規定価格と一致するか。原価の有無とは独立。
    if not cost_confirmed:
        classification = "F"
        detail = "価格履歴レコードはあるが仕入原価が未登録"
    elif purchaser_type == "スタッフ" and staff_rule_detail == "missing":
        classification = "G"
        detail = "スタッフ購入だが、該当商品・取引日の社割価格履歴が無い(原価は判明しているため確定売上原価には含める)"
    elif purchaser_type == "スタッフ" and staff_rule_detail == "discount_percent_unrounded":
        classification = "I"
        detail = "割合ベースの社割価格(端数処理ルール未確定)のため現場確認が必要"
    elif purchaser_type == "スタッフ" and expected_total_staff is not None:
        diff, which, _ = _best_match(expected_total_staff)
        confirmed_rounding_yen = staff_rec.get("confirmed_rounding_tolerance_yen") if staff_rec else None
        if abs(diff) < 0.01:
            classification, detail = "A", f"正常({which}の実売価格がスタッフ価格と一致)"
        elif confirmed_rounding_yen is not None and abs(diff) <= confirmed_rounding_yen:
            # 税込価格×割引率(社割)の計算結果が小数になる商品について、「値引額を丸める」
            # 「最終価格を丸める」のどちらで丸めるかにより、正式なスタッフ価格(登録済みの
            # explicit_staff_price_incl_tax)との間に生じる円単位の差。オーナーが商品ごとに
            # 確認・確定した場合のみ、staff_price_rules.yamlのconfirmed_rounding_tolerance_yen
            # に登録し、価格異常(B・D)として検出しない(2026-09-15確定。推測では適用しない)。
            classification, detail = "A", (
                f"正常(登録済み社割価格{expected_total_staff}円との差額{diff:+.2f}円は、"
                "税込価格×割引率の円未満端数処理の違いによる確定済みの丸め差。"
                "正式な社割価格を優先して使用。product-audit-spec.md §16参照)"
            )
        elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
            classification, detail = "B", f"丸め差の可能性({which}との差額{diff:+.2f}円、暫定許容範囲内)"
        else:
            classification = "D"
            detail = f"スタッフ価格と不一致({which}との差額{diff:+.2f}円{_qty_note(qty, gross_incl_tax, net_incl_tax, which)})"
    else:
        # 一般顧客、または購入者区分「不明」(通常価格で判定)
        diff, which, _ = _best_match(expected_total_regular)
        if diff is None:
            classification, detail = "G", "通常価格が価格履歴に無い"
        elif abs(diff) < 0.01:
            classification, detail = "A", f"正常({which}の実売価格が通常価格と一致)"
        elif abs(diff) <= ROUNDING_TOLERANCE_YEN:
            classification, detail = "B", f"丸め差の可能性({which}との差額{diff:+.2f}円、暫定許容範囲内)"
        else:
            classification = "C"
            detail = f"通常価格と不一致({which}との差額{diff:+.2f}円{_qty_note(qty, gross_incl_tax, net_incl_tax, which)})"

    # 既知差異(K)の判定: C・D(価格不一致)になった場合のみ、取引日以外の期間の
    # 価格履歴レコードと実売価格(値引前)が一致しないか確認する。一致すれば、
    # 日報Excelの商品マスター更新による既知の表示価格変動として区別する
    # (product-audit-spec.md §8・2026-09-12確定)。
    known_flags: list[str] = []
    audit_adjustment_excl_tax = 0.0
    audit_adjustment_applicable = False
    audit_adjustment_note = ""
    if classification in ("C", "D"):
        other_rec = _find_known_discrepancy(records, rec, qty, gross_incl_tax)
        if other_rec is not None:
            other_end = other_rec.get("effective_end_date") or "現在も有効"
            classification = "K"
            detail = (
                f"取引日時点の価格履歴({rec['effective_start_date']}〜{rec.get('effective_end_date') or '現在'}、"
                f"{regular_price}円)とは不一致だが、別期間の価格履歴({other_rec['effective_start_date']}〜{other_end}、"
                f"{other_rec['regular_price_incl_tax']}円)と一致。日報Excelの商品マスターが後日更新され、"
                "過去の表示価格が変わったことによる既知の差異(product-audit-spec.md §8参照)。"
            )

            # 監査調整(audit_adjustment)の判定: 一致した別期間の価格履歴の開始日が、
            # 取引日時点の価格履歴の開始日より後(=将来)である場合のみ、「商品マスターが
            # 後日更新され、過去の取引の表示が歴史的事実と異なっている」パターンと判断する。
            # 一致した別期間が取引日時点より前(過去)の場合は、将来のマスター更新による
            # 表示汚染とは異なる可能性があるため、自動調整の対象にはしない
            # (2026-09-16確定。product-audit-spec.md §28参照。§23の「監査調整売上」と
            # 「ルール基準参考額」を混同しないこと)。
            own_start = dt.date.fromisoformat(rec["effective_start_date"])
            other_start = dt.date.fromisoformat(other_rec["effective_start_date"])
            if other_start > own_start:
                own_excl = rec.get("regular_price_excl_tax")
                other_excl = other_rec.get("regular_price_excl_tax")
                if own_excl is not None and other_excl is not None:
                    audit_adjustment_applicable = True
                    audit_adjustment_excl_tax = round((own_excl - other_excl) * qty, 2)
                    audit_adjustment_note = (
                        f"監査調整対象: 取引日時点({rec['effective_start_date']}〜)の税抜正規価格"
                        f"{own_excl}円/個に対し、表示は後日({other_rec['effective_start_date']}〜)の"
                        f"税抜正規価格{other_excl}円/個に基づいており、歴史的事実と異なる。"
                        f"管理会計PL採用売上への調整額{audit_adjustment_excl_tax:+.2f}円"
                        "(product-audit-spec.md §28参照)。"
                    )
                else:
                    audit_adjustment_note = (
                        "既知差異(K)だが、税抜正規価格が価格履歴に未登録のため監査調整額を"
                        "推測せず算出しない(調整額0円のまま保持)。"
                    )
            else:
                audit_adjustment_note = (
                    f"既知差異(K)だが、一致した価格履歴({other_rec['effective_start_date']}〜)が"
                    f"取引日時点の価格履歴({rec['effective_start_date']}〜)より前の期間のため、"
                    "商品マスターの後日更新による表示汚染パターンとは異なる。自動での監査調整は"
                    "行わない(個別確認が必要な場合はオーナー確認のうえ別途対応)。"
                )
            detail = f"{detail} {audit_adjustment_note}"

    if classification == "K":
        # 既知差異は原因特定済みのため、追加の要確認フラグ(異常値引き等)は付与せず、
        # 重大度も固定する(要確認・重大エラーには含めない)。
        flags: list[str] = []
        severity = SEVERITY_BY_CLASSIFICATION["K"]
    else:
        # 追加の要確認フラグ(判定区分A〜Iとは別軸。いずれも「不正確定」ではなく現場確認対象)。
        flags = []
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

    if quantity_correction_note:
        detail = f"{detail} / {quantity_correction_note}"

    return ProductTransaction(
        date=date, sheet_row=row, customer_name=customer_name,
        purchaser_type=purchaser_type, product_status=determine_product_status(True),
        product_name=product_name, quantity=quantity,
        regular_price_incl_tax_expected=regular_price,
        staff_price_incl_tax_expected=staff_price_expected,
        actual_price_incl_tax=gross_incl_tax, discount_incl_tax=discount_incl_tax,
        net_price_incl_tax=net_incl_tax, tax_excl_revenue=tax_excl_revenue,
        cost_excl_tax_unit=cost_unit, cost_excl_tax_total=cost_total,
        gross_profit=profit, gross_margin=margin,
        classification=classification, classification_detail=detail,
        cost_confirmed=cost_confirmed,
        price_status=PRICE_STATUS_BY_CLASSIFICATION[classification],
        severity=severity, flags=flags,
        purchaser_alias_note=alias_note,
        purchaser_name_note=purchaser_name_note,
        quantity_correction_note=quantity_correction_note,
        share_marker_raw=note_raw,
        audit_adjustment_excl_tax=audit_adjustment_excl_tax,
        audit_adjustment_applicable=audit_adjustment_applicable,
        audit_adjustment_note=audit_adjustment_note,
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


def parse_monthly_summary(wb) -> dict:
    """月報集計シートから、店舗全体の実売・内訳(回数券等・物販)の当月累計を読み取る
    (読み取り専用、外部参照を経由しない自己完結型の値。product-audit-spec.md §17参照)。

    「772,881.85円」等の物販監査上の金額を、誤って「店舗全体売上」と読み違えないよう、
    店舗全体の実売合計(店舗全体売上)・内回数券等(施術・回数券)・内物販の3つを
    明確に区別して返す。ラベル(B列)を検索して該当行を特定し、ヘッダー行で
    「累計」列を探すことで、行位置の変更にある程度頑健にする。値が見つからない
    場合はNoneのまま返し、推測で埋めない。
    """
    ws = wb["月報集計"]
    total_col = None
    for row in range(1, 10):
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=row, column=col).value == "累計":
                total_col = col
                break
        if total_col:
            break

    labels = {
        "store_total_actual_sales": "実売（実績）",
        "ticket_treatment_sales_reported": "内回数券等",
        "retail_sales_reported": "内物販",
    }
    result: dict = {k: None for k in labels}
    if total_col is None:
        return result
    for row in range(1, ws.max_row + 1):
        label = ws.cell(row=row, column=2).value
        for key, target_label in labels.items():
            if label == target_label:
                v = ws.cell(row=row, column=total_col).value
                result[key] = v if isinstance(v, (int, float)) else None
    return result


def _apply_confirmed_share(group: list[ProductTransaction], share_total: float, linked_product: str) -> dict:
    """売上シェア(確定済みの業務ルール)をグループに適用する(in-place)。

    2名・3名いずれも、日報I列に「売上シェア」の明示記載があり、かつグループ合計金額が
    規定価格(×数量)と一致する場合は確定した業務ルールとして扱う(2026-09-15確定。
    I列記載が無く金額の一致のみで推測した場合は_apply_candidate_shareで候補扱いとし、
    ここでは確定しない)。
    「価格異常にしない・商品不明にしない・重複売上にしない」ため、classification・
    price_status・severityを正常化する。売上(tax_excl_revenue)・購入者区分別の売上実績は
    変更しない(各スタッフの取り分をそのまま保持)。
    原価は「商品あり」側の1個分だけを実際の原価として扱い、「無」側は0円・原価確定済みとする
    (原価の二重計上を避ける)。数量も「無」側は0として二重計上を避ける。
    """
    group_id = f"{group[0].date}-share-{min(t.sheet_row for t in group)}"
    rows = sorted(t.sheet_row for t in group)
    for t in group:
        t.transaction_type = "shared_sale"
        t.share_group_id = group_id
        t.share_count = len(group)
        t.share_total = share_total
        t.linked_product = linked_product
        t.linked_rows = rows
        t.classification = "S"
        t.classification_detail = (
            f"売上シェア(確定):「{linked_product}」を{len(group)}名で分担入力(行{rows}の合計{share_total}円が"
            "規定価格と一致)。価格異常・商品不明・重複売上のいずれとしても扱わない(product-audit-spec.md §11)。"
        )
        t.price_status = PRICE_STATUS_BY_CLASSIFICATION["S"]
        t.severity = SEVERITY_BY_CLASSIFICATION["S"]
        t.flags = []
        if t.product_name in UNKNOWN_PRODUCT_LABELS:
            # 「無」側: 商品は特定済み(シェア分担入力と判明)、原価・数量は二重計上しない。
            t.product_status = "商品特定済み"
            t.cost_excl_tax_total = 0.0
            t.cost_confirmed = True
            t.quantity = 0
            # 原価0円のため、このスタッフの分担売上(税抜)がそのまま粗利益になる。
            # 更新しないとNoneのまま残り、集計側でgetattr(...) or 0により0円と
            # 誤集計されてしまう(2026-09-15判明・修正)。
            t.gross_profit = round(t.tax_excl_revenue, 2)
            t.gross_margin = 1.0 if t.tax_excl_revenue else None
    return {
        "share_group_id": group_id, "transaction_type": "shared_sale", "date": group[0].date,
        "linked_product": linked_product, "share_count": len(group), "share_total": share_total,
        "linked_rows": rows,
    }


def _apply_candidate_share(group: list[ProductTransaction], share_total: float, linked_product: str) -> dict:
    """売上シェア候補(未確定)をグループに注記する(in-place)。

    日報I列に「売上シェア」の明示記載が無く、金額の一致(近接行・同日等)のみから
    推測したグループに使う(2026-09-15確定)。I列記載という一次証跡が無いまま
    推測だけで自動確定しないため、classification・severity・原価・数量は一切
    変更しない(要確認のまま残す)。参考情報(transaction_type・share_*)のみ付与する。
    """
    group_id = f"{group[0].date}-sharecandidate-{min(t.sheet_row for t in group)}"
    rows = sorted(t.sheet_row for t in group)
    for t in group:
        t.transaction_type = "shared_sale_candidate"
        t.share_group_id = group_id
        t.share_count = len(group)
        t.share_total = share_total
        t.linked_product = linked_product
        t.linked_rows = rows
    return {
        "share_group_id": group_id, "transaction_type": "shared_sale_candidate", "date": group[0].date,
        "linked_product": linked_product, "share_count": len(group), "share_total": share_total,
        "linked_rows": rows,
    }


def detect_and_apply_shared_sales(transactions: list[ProductTransaction]) -> list[dict]:
    """売上シェア(1商品を複数スタッフで分担入力する運用)を検出し、対象取引に反映する(in-place)。

    正式ルール(2026-09-15確定、product-audit-spec.md §11)。判定は次の優先順位で行う。

    1. 日報I列(既存 単発・回数券欄)に「売上シェア」の明示記載があるかを最優先の根拠として確認する
       (SHARE_MARKER_LABEL・_has_share_marker)。明示記載は、店舗が自らその行を
       シェア分担入力だと記録した一次証跡であり、金額の一致や近接行からの推測より
       優先する。
    2. 明示記載がある「無」行(1行または2行)について、商品名が記録された行(実売価格が
       規定価格に届かない、classification=C)と、「商品あり行の金額 + 無行の金額(1〜2行) =
       取引日時点の規定価格×数量」が一致するかを確認する。一致すれば、2名・3名いずれも
       確定した業務ルールとして扱い、classification等を正常化する(_apply_confirmed_share)。
       候補が複数ある(曖昧)場合は確定しない。
    3. I列に明示記載が無い「無」行については、金額の一致(同日・近接行)のみを根拠に
       「シェア候補」として情報を付与するのみで、classification・severity・原価・数量は
       変更しない(推測だけでは自動確定しない。_apply_candidate_share)。
    4. I列に明示記載があるのに対応する商品あり行を特定できなかった行(例: 商品あり行が
       日報上に見当たらないケース)は、transaction_typeを"unknown"とし、要確認フラグを
       追加するのみで、classification・severityは変更しない(自動正常化しない)。

    戻り値は確定・候補として検出したグループのサマリ一覧(出力・ログ用)。
    """
    groups_summary: list[dict] = []
    by_date: dict[str, list[ProductTransaction]] = {}
    for t in transactions:
        by_date.setdefault(t.date, []).append(t)

    def _unique_combo(pool, deficit, used_ids):
        """poolから、deficitに一致する組み合わせ(サイズ1→2の順)を一意に探す。

        該当サイズで複数の組み合わせが見つかった(曖昧)場合は、そのサイズで探索を止め、
        Noneを返す(より大きいサイズへは進まない。誤確定より見逃しを優先する)。
        """
        for size in (1, 2):
            candidates = [u for u in pool if id(u) not in used_ids]
            combos = [
                c for c in itertools.combinations(candidates, size)
                if abs(sum(u.actual_price_incl_tax for u in c) - deficit) < SHARE_AMOUNT_TOLERANCE_YEN
            ]
            if len(combos) == 1:
                return combos[0]
            if len(combos) > 1:
                return None
        return None

    for day_txs in by_date.values():
        # 「商品あり」候補: 通常価格不一致(C)で、規定価格に対して不足額(deficit)がある行のみ対象。
        # K(既知差異)・D(スタッフ価格不一致)・G/F等は対象外とし、影響範囲を限定する。
        known_candidates = []
        for t in day_txs:
            if t.classification != "C":
                continue
            if t.regular_price_incl_tax_expected is None:
                continue
            qty = t.quantity if t.quantity else 1
            expected_total = t.regular_price_incl_tax_expected * qty
            deficit = round(expected_total - t.actual_price_incl_tax, 2)
            if deficit <= SHARE_AMOUNT_TOLERANCE_YEN:
                continue
            known_candidates.append((t, expected_total, deficit))

        marked_pool = [
            t for t in day_txs
            if t.product_name in UNKNOWN_PRODUCT_LABELS and t.classification == "H" and _has_share_marker(t)
        ]
        unmarked_pool = [
            t for t in day_txs
            if t.product_name in UNKNOWN_PRODUCT_LABELS and t.classification == "H" and not _has_share_marker(t)
        ]
        used_ids: set[int] = set()

        # パス1: I列に「売上シェア」の明示記載がある行を最優先で確定する(2名・3名とも)。
        remaining = []
        for t, expected_total, deficit in known_candidates:
            combo = _unique_combo(marked_pool, deficit, used_ids)
            if combo:
                group = [t] + list(combo)
                summary = _apply_confirmed_share(group, expected_total, t.product_name)
                groups_summary.append(summary)
                for u in combo:
                    used_ids.add(id(u))
            else:
                remaining.append((t, expected_total, deficit))

        # パス2: I列に明示記載が無い場合のみ、金額の一致からの「候補」判定(未確定のまま注記のみ)。
        for t, expected_total, deficit in remaining:
            combo = _unique_combo(unmarked_pool, deficit, used_ids)
            if combo:
                group = [t] + list(combo)
                summary = _apply_candidate_share(group, expected_total, t.product_name)
                groups_summary.append(summary)
                for u in combo:
                    used_ids.add(id(u))

        # I列に「売上シェア」の明示記載があるのに、対応する商品あり行が特定できなかった行。
        # 推測で正常化せず、要確認のまま残しつつ、記載があった事実だけ記録する。
        for u in marked_pool:
            if id(u) in used_ids:
                continue
            u.transaction_type = "unknown"
            u.flags = u.flags + [
                "日報I列に「売上シェア」の記載があるが、対応する商品あり行を特定できず、"
                "自動では正常化しない(要確認のまま)"
            ]
            u.classification_detail = (
                u.classification_detail
                + " / I列に「売上シェア」の記載あり(対応する商品あり行が未特定のため要確認のまま)"
            )

    return groups_summary


def load_confirmed_structural_discrepancies(path: Path) -> list[dict]:
    """確定済み「既知構造差異」マスターを読み込む(2026-09-16確定、
    product-audit-spec.md §26参照)。

    原因が既に特定済みだが、日報・月報集計の構造上の理由により解消できない
    (または意図的に解消しない)差異を登録する。店舗全体売上照合等で検出された
    差額が、ここに登録された既知の差額と一致する場合、「原因不明の店舗全体要確認」
    ではなく「既知構造差異」として区別する。差異そのものは消さず、日報実績・
    月報集計・差額・原因はすべて保持したまま扱う。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    return doc.get("discrepancies", []) or []


def load_confirmed_manual_share_groups(path: Path) -> list[dict]:
    """現場確認に基づく手動確定・売上シェアグループマスターを読み込む(2026-09-15確定、
    product-audit-spec.md §21参照)。
    """
    with open(path, "r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    return doc.get("groups", []) or []


def apply_confirmed_manual_share_groups(
    transactions: list[ProductTransaction], *, store_id: str, year_month: str, groups: list[dict],
) -> list[dict]:
    """現場確認により確定した売上シェアグループを適用する(in-place)。

    detect_and_apply_shared_sales(§11)の自動検出は「1商品の商品あり行+その商品の
    無行」という単純な組み合わせしか扱えない。1つの「無」行が複数の異なる商品の
    売上シェアを同時に受け止めているケース等、自動検出の対象外だが現場確認により
    確定した組み合わせをここで明示的に処理する。_apply_confirmed_share と同じ
    ロジックを再利用するため、確定済み売上シェア(§11)と同様にclassification="S"・
    severity="正常"となり、「無」側の行のみ原価・数量を0にする(商品あり側の行は
    日報記載どおりの数量・原価をそのまま使うため、二重計上は発生しない)。

    グループを構成する行が(存在しない行番号の指定などにより)全て見つからない
    場合は、安全側に倒して何も適用しない。戻り値は適用したグループのサマリ一覧。
    """
    by_key = {(t.date, t.sheet_row): t for t in transactions}
    summaries = []
    for g in groups:
        if g["store_id"] != store_id or g["year_month"] != year_month:
            continue
        date_key = g["date"]
        group_members = [by_key[(date_key, r)] for r in g["rows"] if (date_key, r) in by_key]
        if len(group_members) != len(g["rows"]):
            continue
        share_total = round(sum(t.actual_price_incl_tax for t in group_members), 2)
        summary = _apply_confirmed_share(group_members, share_total, g["linked_product"])
        note = (g.get("note") or "").strip()
        if note:
            for t in group_members:
                t.classification_detail = t.classification_detail + " / " + note
        summaries.append(summary)
    return summaries
