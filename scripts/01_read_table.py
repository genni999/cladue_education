# -*- coding: utf-8 -*-
"""
1단계: 엑셀을 읽어서 표(부서별 경비 상세 내역)로 만든다.
- 원본 파일은 읽기만 하고 수정하지 않는다.
- 원본에 섞여 있는 소계/계/합계/총계 행은 제외하고, 항목 단위 상세 데이터만 표로 만든다.
- 결과와 검증 결과를 출력/ 에 쓴다.
"""
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "부서별_경비_2601.xlsx"
OUTPUT_DIR = BASE_DIR / "출력"
OUTPUT_FILE = OUTPUT_DIR / "01_부서별_경비_표.xlsx"

# 원본에서 소계성 행을 걸러내기 위한 항목 라벨(공백 변형 포함)
SUBTOTAL_LABELS = {"소계", "소 계", "계", "합계"}


def load_raw(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="경비", dtype={"금액": "Int64"})
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_table(raw: pd.DataFrame) -> pd.DataFrame:
    # 소계/합계 행: 부서가 비어 있고 항목이 소계성 라벨인 행
    is_subtotal_row = raw["부서"].isna() & raw["항목"].isin(SUBTOTAL_LABELS)
    # 총계 행: 부서가 "총계"이고 항목이 비어 있는 행
    is_grandtotal_row = raw["부서"] == "총계"

    detail = raw[~is_subtotal_row & ~is_grandtotal_row].copy()
    detail = detail.dropna(subset=["부서", "항목"])
    detail["금액"] = detail["금액"].astype("Int64")
    detail = detail.reset_index(drop=True)
    return detail


def validate(raw: pd.DataFrame, detail: pd.DataFrame) -> list[str]:
    lines = []
    lines.append(f"원본 행 수: {len(raw)}")
    lines.append(f"상세(항목) 행 수: {len(detail)}")

    null_dept = detail["부서"].isna().sum()
    null_item = detail["항목"].isna().sum()
    null_amt = detail["금액"].isna().sum()
    lines.append(f"부서 결측치: {null_dept}건")
    lines.append(f"항목 결측치: {null_item}건")
    lines.append(f"금액 결측치: {null_amt}건")

    neg_amt = (detail["금액"] < 0).sum()
    lines.append(f"음수 금액: {neg_amt}건")

    depts = sorted(detail["부서"].unique())
    lines.append(f"부서 목록({len(depts)}개): {', '.join(depts)}")

    ok = (null_dept == 0) and (null_item == 0) and (null_amt == 0) and (neg_amt == 0)
    lines.append(f"검증 결과: {'통과' if ok else '실패'}")
    return lines


def main() -> None:
    if not INPUT_FILE.exists():
        print(f"원본 파일을 찾을 수 없습니다: {INPUT_FILE}", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    raw = load_raw(INPUT_FILE)
    detail = build_table(raw)

    detail.to_excel(OUTPUT_FILE, index=False)

    report_lines = validate(raw, detail)
    report_path = OUTPUT_DIR / "01_검증결과.txt"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    print(f"[출력] 표: {OUTPUT_FILE}")
    print(f"[출력] 검증 결과: {report_path}")
    print()
    print("=== 검증 결과 ===")
    for line in report_lines:
        print(line)


if __name__ == "__main__":
    main()
