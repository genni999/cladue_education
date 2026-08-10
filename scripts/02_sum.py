# -*- coding: utf-8 -*-
"""
2단계: 1단계 출력(부서별 경비 상세 표)을 읽어 부서별 합계를 계산한다.
- 원본 파일은 검증용으로만 읽고 수정하지 않는다.
- 결과와 검증 결과를 출력/ 에 쓴다.
"""
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "출력" / "01_부서별_경비_표.xlsx"
RAW_FILE = BASE_DIR / "부서별_경비_2601.xlsx"  # 검증 대조용 원본, 읽기만 함
OUTPUT_DIR = BASE_DIR / "출력"
OUTPUT_FILE = OUTPUT_DIR / "02_부서별_합계.xlsx"
REPORT_FILE = OUTPUT_DIR / "02_검증결과.txt"

SUBTOTAL_LABELS = {"소계", "소 계", "계", "합계"}


def load_detail(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype={"금액": "Int64"})
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_totals(detail: pd.DataFrame) -> pd.DataFrame:
    totals = (
        detail.groupby("부서", as_index=False)["금액"]
        .sum()
        .rename(columns={"금액": "합계금액"})
        .sort_values("부서")
        .reset_index(drop=True)
    )
    return totals


def load_reference_totals(path: Path) -> tuple[dict, int | None]:
    """원본 파일의 소계/총계 행에서 부서별·전체 합계를 읽어 대조용으로 반환한다."""
    if not path.exists():
        return {}, None
    raw = pd.read_excel(path, sheet_name="경비", dtype={"금액": "Int64"})
    raw.columns = [str(c).strip() for c in raw.columns]

    # 소계 행은 부서가 비어 있음 -> 바로 위 부서 행의 값을 ffill로 채워 매핑
    raw["부서_ffill"] = raw["부서"].ffill()
    subtotal_rows = raw[raw["부서"].isna() & raw["항목"].isin(SUBTOTAL_LABELS)]
    dept_totals = dict(zip(subtotal_rows["부서_ffill"], subtotal_rows["금액"]))

    grandtotal_rows = raw[raw["부서"] == "총계"]
    grand_total = int(grandtotal_rows["금액"].iloc[0]) if len(grandtotal_rows) else None

    return dept_totals, grand_total


def validate(totals: pd.DataFrame, ref_dept_totals: dict, ref_grand_total: int | None) -> list[str]:
    lines = []
    computed_grand_total = int(totals["합계금액"].sum())
    lines.append(f"부서 수: {len(totals)}")
    lines.append(f"계산된 전체 합계: {computed_grand_total:,}")

    ok = True

    if ref_grand_total is not None:
        match = computed_grand_total == ref_grand_total
        ok &= match
        lines.append(
            f"원본 총계와 대조: 원본={ref_grand_total:,} / 계산={computed_grand_total:,} -> {'일치' if match else '불일치'}"
        )
    else:
        lines.append("원본 총계 행을 찾지 못해 대조를 건너뜀")

    if ref_dept_totals:
        lines.append("부서별 대조:")
        for _, row in totals.iterrows():
            dept, amt = row["부서"], int(row["합계금액"])
            ref_amt = ref_dept_totals.get(dept)
            if ref_amt is None:
                ok = False
                lines.append(f"  - {dept}: 계산={amt:,} / 원본 소계 없음 -> 불일치")
            else:
                match = amt == ref_amt
                ok &= match
                lines.append(
                    f"  - {dept}: 계산={amt:,} / 원본={int(ref_amt):,} -> {'일치' if match else '불일치'}"
                )
    else:
        lines.append("원본 소계 행을 찾지 못해 부서별 대조를 건너뜀")

    lines.append(f"검증 결과: {'통과' if ok else '실패'}")
    return lines


def main() -> None:
    if not INPUT_FILE.exists():
        print(f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}", file=sys.stderr)
        print("먼저 01_read_table.py를 실행하세요.", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    detail = load_detail(INPUT_FILE)
    totals = build_totals(detail)
    totals.to_excel(OUTPUT_FILE, index=False)

    ref_dept_totals, ref_grand_total = load_reference_totals(RAW_FILE)
    report_lines = validate(totals, ref_dept_totals, ref_grand_total)
    REPORT_FILE.write_text("\n".join(report_lines), encoding="utf-8")

    print(f"[출력] 부서별 합계: {OUTPUT_FILE}")
    print(f"[출력] 검증 결과: {REPORT_FILE}")
    print()
    print("=== 검증 결과 ===")
    for line in report_lines:
        print(line)


if __name__ == "__main__":
    main()
