"""日報からの再集計値と、「2026 売上現状」等の管理用集計表を突き合わせるロジック。

accounting/docs/accounting-spec.md 第9章の正式ルールに基づく。
- 日報が正本。管理用集計表は照合対象であり、自動で書き換えない。
- 合計は一致するが内訳(新規/既存)だけが違う場合は「分類差異」。
- 合計自体が一致しない場合は「金額差異」。
- 未確定(区分矛盾)の取引は、新規・既存いずれにも計上しないため、
  「分類差異」の一因として別途参考情報に出す。
"""
from __future__ import annotations

from dataclasses import dataclass

TOLERANCE_YEN = 1.0  # 浮動小数点の丸め誤差を吸収するための許容差(円)


@dataclass
class DayComparison:
    date_key: str
    verdict: str  # "一致" / "分類差異" / "金額差異" / "データなし(売上現状未入力)"
    report_new: float | None
    report_existing: float | None
    report_retail: float | None
    report_unresolved: float
    report_total: float | None
    status_new: float | None
    status_existing: float | None
    status_retail: float | None
    status_total: float | None
    diff_new: float | None
    diff_existing: float | None
    diff_retail: float | None
    diff_total: float | None
    note: str = ""


def _close(a: float | None, b: float | None, tol: float = TOLERANCE_YEN) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) < tol


def compare_day(date_key: str, day_result, status_entry: dict | None) -> DayComparison:
    """day_result: xlsx_report.DayResult。status_entry: {"new","existing","retail","total"} または None。"""
    report_new = day_result.new_total
    report_existing = day_result.existing_total
    report_retail = day_result.retail_total
    report_total = day_result.grand_total

    if status_entry is None:
        return DayComparison(
            date_key=date_key, verdict="データなし(売上現状未入力)",
            report_new=report_new, report_existing=report_existing,
            report_retail=report_retail, report_unresolved=day_result.unresolved_total,
            report_total=report_total,
            status_new=None, status_existing=None, status_retail=None, status_total=None,
            diff_new=None, diff_existing=None, diff_retail=None, diff_total=None,
            note="「2026 売上現状」にこの日の入力がない、またはスナップショットに含まれていない",
        )

    status_new = float(status_entry["new"])
    status_existing = float(status_entry["existing"])
    status_retail = float(status_entry["retail"])
    status_total = float(status_entry["total"])

    diff_new = round(report_new - status_new, 2)
    diff_existing = round(report_existing - status_existing, 2)
    diff_retail = round(report_retail - status_retail, 2)
    diff_total = round(report_total - status_total, 2)

    total_matches = _close(report_total, status_total)
    all_categories_match = (
        _close(report_new, status_new)
        and _close(report_existing, status_existing)
        and _close(report_retail, status_retail)
    )

    note = ""
    if not total_matches:
        verdict = "金額差異"
        if day_result.unresolved_total != 0:
            note = (
                f"日報側に数式エラー等で金額が確定できない取引が{day_result.unresolved_total:g}円分あり、"
                "合計に反映できていない可能性がある"
            )
    elif not all_categories_match:
        verdict = "分類差異"
        if day_result.category_flagged_rows:
            note = (
                f"区分(新規/既存)ラベルと計算根拠が食い違う取引が{len(day_result.category_flagged_rows)}件あり、"
                "「2026 売上現状」の分類と異なる可能性がある(金額自体は日報の計算式どおり集計済み)"
            )
    else:
        verdict = "一致"

    return DayComparison(
        date_key=date_key, verdict=verdict,
        report_new=report_new, report_existing=report_existing,
        report_retail=report_retail, report_unresolved=day_result.unresolved_total,
        report_total=report_total,
        status_new=status_new, status_existing=status_existing,
        status_retail=status_retail, status_total=status_total,
        diff_new=diff_new, diff_existing=diff_existing,
        diff_retail=diff_retail, diff_total=diff_total,
        note=note,
    )


def build_comparison_table(day_results: dict, status_daily: dict) -> list[DayComparison]:
    comparisons = []
    for day in sorted(day_results.keys(), key=lambda x: int(x)):
        status_entry = status_daily.get(day)
        comparisons.append(compare_day(day, day_results[day], status_entry))
    return comparisons
