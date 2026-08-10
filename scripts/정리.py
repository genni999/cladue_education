# -*- coding: utf-8 -*-
"""
부서별_경비_2601.xlsx를 읽어 부서별 합계와 전체 합계를 계산한다.
(01_read_table.py + 02_sum.py를 하나로 정리한 버전)
- 원본 파일은 읽기만 하고 수정하지 않는다.
- 결과와 검증 결과를 출력/ 에 쓴다.
"""
import sys
from pathlib import Path

import pandas as pd

# 콘솔에 한글이 깨져 보이지 않도록 표준출력을 UTF-8로 고정
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "부서별_경비_2601.xlsx"
OUTPUT_DIR = BASE_DIR / "출력"
OUTPUT_FILE = OUTPUT_DIR / "정리_부서별_합계.xlsx"
REPORT_FILE = OUTPUT_DIR / "정리_검증결과.txt"

# 원본에서 소계성 행을 걸러내기 위한 항목 라벨(공백 변형 포함)
SUBTOTAL_LABELS = {"소계", "소 계", "계", "합계"}


def load_raw(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="경비", dtype={"금액": "Int64"})
    df.columns = [str(c).strip() for c in df.columns]
    return df


def split_rows(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict, int | None]:
    """원본을 항목 상세 행 / 부서별 소계 / 총계로 분리한다."""
    is_subtotal_row = raw["부서"].isna() & raw["항목"].isin(SUBTOTAL_LABELS)
    is_grandtotal_row = raw["부서"] == "총계"

    detail = raw[~is_subtotal_row & ~is_grandtotal_row].copy()
    detail = detail.dropna(subset=["부서", "항목"])
    detail["금액"] = detail["금액"].astype("Int64")
    detail = detail.reset_index(drop=True)

    raw["부서_ffill"] = raw["부서"].ffill()
    subtotal_rows = raw[is_subtotal_row]
    ref_dept_totals = dict(zip(subtotal_rows["부서_ffill"], subtotal_rows["금액"]))

    grandtotal_rows = raw[is_grandtotal_row]
    ref_grand_total = int(grandtotal_rows["금액"].iloc[0]) if len(grandtotal_rows) else None

    return detail, ref_dept_totals, ref_grand_total


def build_totals(detail: pd.DataFrame) -> pd.DataFrame:
    totals = (
        detail.groupby("부서", as_index=False)["금액"]
        .sum()
        .rename(columns={"금액": "합계금액"})
        .sort_values("부서")
        .reset_index(drop=True)
    )
    return totals


def check(label: str, value_a: int, value_b: int) -> tuple[bool, str]:
    diff = value_a - value_b
    passed = diff == 0
    tag = "[PASS]" if passed else "[FAIL]"
    line = f"{tag} {label} : {value_a:,} vs {value_b:,} (차이 {diff:,})"
    return passed, line


def validate(totals: pd.DataFrame, ref_dept_totals: dict, ref_grand_total: int | None) -> list[str]:
    lines = []
    ok = True

    # (1) 부서별: 항목 금액의 합 = 그 부서 소계
    for _, row in totals.iterrows():
        dept, amt = row["부서"], int(row["합계금액"])
        ref_amt = ref_dept_totals.get(dept)
        if ref_amt is None:
            ok = False
            lines.append(f"[FAIL] {dept} 항목합계 vs 부서소계 : {amt:,} vs (원본 소계 없음)")
            continue
        passed, line = check(f"{dept} 항목합계 vs 부서소계", amt, int(ref_amt))
        ok &= passed
        lines.append(line)

    # (2) 부서 소계 4개의 합 = 파일 맨 아래 총계
    dept_subtotal_sum = int(sum(ref_dept_totals.values())) if ref_dept_totals else None
    if dept_subtotal_sum is not None and ref_grand_total is not None:
        passed, line = check("부서 소계 합 vs 원본 총계", dept_subtotal_sum, ref_grand_total)
        ok &= passed
        lines.append(line)
    else:
        ok = False
        lines.append("[FAIL] 부서 소계 합 vs 원본 총계 : 원본 소계/총계 행을 찾지 못함")

    # (3) 우리가 낸 합계 = 총계
    computed_grand_total = int(totals["합계금액"].sum())
    if ref_grand_total is not None:
        passed, line = check("계산 합계 vs 원본 총계", computed_grand_total, ref_grand_total)
        ok &= passed
        lines.append(line)
    else:
        ok = False
        lines.append("[FAIL] 계산 합계 vs 원본 총계 : 원본 총계 행을 찾지 못함")

    lines.append(f"검증 결과: {'통과' if ok else '실패'}")
    return lines


def main() -> None:
    if not INPUT_FILE.exists():
        print(f"원본 파일을 찾을 수 없습니다: {INPUT_FILE}", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    raw = load_raw(INPUT_FILE)
    detail, ref_dept_totals, ref_grand_total = split_rows(raw)
    totals = build_totals(detail)

    totals.to_excel(OUTPUT_FILE, index=False)

    report_lines = validate(totals, ref_dept_totals, ref_grand_total)
    REPORT_FILE.write_text("\n".join(report_lines), encoding="utf-8")

    print(f"[출력] 부서별 합계: {OUTPUT_FILE}")
    print(f"[출력] 검증 결과: {REPORT_FILE}")
    print()
    print("=== 부서별 합계 ===")
    for _, row in totals.iterrows():
        print(f"  {row['부서']}: {int(row['합계금액']):,}")
    print(f"  전체 합계: {int(totals['합계금액'].sum()):,}")
    print()
    print("=== 검증 결과 ===")
    for line in report_lines:
        print(line)


if __name__ == "__main__":
    main()
