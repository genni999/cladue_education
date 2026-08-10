# -*- coding: utf-8 -*-
"""
부서별_경비_YYMM.xlsx를 읽어 부서별 합계와 전체 합계를 계산한다.
(01_read_table.py + 02_sum.py를 하나로 정리한 버전)
- 원본 파일은 읽기만 하고 수정하지 않는다.
- 상세내역/부서별합계/검증결과를 시트로 묶어 출력/{연월}_정리.xlsx 하나로 낸다.

사용법:
    py scripts/정리.py            # 폴더에서 부서별_경비_*.xlsx 중 가장 최근 파일을 자동 사용
    py scripts/정리.py 2602       # 부서별_경비_2602.xlsx 지정
    py scripts/정리.py 부서별_경비_2602.xlsx   # 파일명/경로 직접 지정
"""
import sys
from pathlib import Path

import pandas as pd

# 콘솔에 한글이 깨져 보이지 않도록 표준출력을 UTF-8로 고정
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "출력"
INPUT_PATTERN = "부서별_경비_*.xlsx"

# 원본에서 소계성 행을 걸러내기 위한 항목 라벨(공백 변형 포함)
SUBTOTAL_LABELS = {"소계", "소 계", "계", "합계"}


def resolve_input_file(arg: str | None) -> Path:
    """인자로 받은 월 코드/파일명을 실제 원본 파일 경로로 바꾼다.
    인자가 없으면 폴더에서 부서별_경비_*.xlsx 중 이름순 가장 나중 파일을 사용한다."""
    if arg:
        candidate = Path(arg)
        if candidate.suffix.lower() != ".xlsx":
            candidate = BASE_DIR / f"부서별_경비_{arg}.xlsx"
        elif not candidate.is_absolute():
            candidate = BASE_DIR / candidate
        return candidate

    candidates = sorted(BASE_DIR.glob(INPUT_PATTERN))
    if not candidates:
        print(f"{BASE_DIR}에서 '{INPUT_PATTERN}' 파일을 찾을 수 없습니다.", file=sys.stderr)
        sys.exit(1)
    return candidates[-1]


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
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    input_file = resolve_input_file(arg)
    if not input_file.exists():
        print(f"원본 파일을 찾을 수 없습니다: {input_file}", file=sys.stderr)
        sys.exit(1)

    # 파일명에서 월 코드(예: 2602)를 뽑아 산출물 파일명에 사용한다.
    tag = input_file.stem.replace("부서별_경비_", "")
    output_file = OUTPUT_DIR / f"{tag}_정리.xlsx"

    OUTPUT_DIR.mkdir(exist_ok=True)

    raw = load_raw(input_file)
    detail, ref_dept_totals, ref_grand_total = split_rows(raw)
    totals = build_totals(detail)
    report_lines = validate(totals, ref_dept_totals, ref_grand_total)
    report_df = pd.DataFrame({"검증 결과": report_lines})

    with pd.ExcelWriter(output_file) as writer:
        detail.to_excel(writer, sheet_name="상세내역", index=False)
        totals.to_excel(writer, sheet_name="부서별합계", index=False)
        report_df.to_excel(writer, sheet_name="검증결과", index=False)

    print(f"[출력] {output_file}")


if __name__ == "__main__":
    main()
