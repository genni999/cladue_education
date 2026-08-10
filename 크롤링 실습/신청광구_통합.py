# -*- coding: utf-8 -*-
"""
MGB(필리핀 광업지질국) ArcGIS 저장소에서 '신청 광구' 두 서비스를 받아 하나로 병합한다.

  구: Tenement_Application            (1,077건 / 2023-11-06 / 전국 / 속성 빈약)
  신: Application_Mining_Tenement     (  181건 / 2026-07-23 / 4개 지역만 / 속성 상세)

두 서비스는 서로의 신·구 버전이 아니다. 광구번호가 겹치는 건 61개뿐이라
어느 한쪽만 쓰면 정보가 손실된다. 그래서 둘 다 남기고 출처를 컬럼으로 구분한다.
어느 값이 맞는지는 판단하지 않는다 (보는 사람이 정할 몫).

컬럼명은 승인광구_전체.csv(신 스키마)에 맞춰 두 서비스를 같은 틀로 정렬한다.

사용법:
    python 신청광구_통합.py

주의: 이 환경에서는 Python urllib의 HTTPS 인증서 검증이 막혀 있어(사내 SSL 검사)
      curl에 --ssl-no-revoke를 붙여 내려받는다. curl이 PATH에 있어야 한다.
"""
import csv
import json
import re
import subprocess
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
OUTPUT = BASE_DIR / "신청광구_통합.csv"

REST = "https://services7.arcgis.com/Z0dvtKpPYjB1vNXq/arcgis/rest/services"
SERVICE_OLD = "Tenement_Application"
SERVICE_NEW = "Application_Mining_Tenement"
PAGE_SIZE = 1000

# 필리핀 지역코드. 구 서비스는 지역 컬럼이 없어 광구번호 접미사에서 유추하는데,
# 접미사에 일련번호·'AMENDED' 같은 값도 섞여 있어 이 목록에 있는 것만 인정한다.
PH_REGIONS = {
    "NCR", "CAR", "I", "II", "III", "IVA", "IVB", "MIMAROPA", "V", "VI", "VII",
    "VIII", "IX", "X", "XI", "XII", "XIII", "ARMM", "BARMM",
}

# 출력 컬럼: 신 스키마를 기준으로 하고, 구 서비스 전용 필드를 뒤에 붙인다.
COLUMNS = [
    "출처", "중복여부",
    "docTenementNum", "tenementType", "tenementName",
    "docRegOfc", "docProvName", "docMuniName",
    "totalAreaHas",
    "miningStage", "tenementStatus", "miningPhase", "operationStatus",
    "DateFiled", "StatusDate",
    "companyRepresentative", "companyAddress", "companyNumber",
    "genType", "tenementMineral", "Remarks",
    # 구 서비스 전용 (필지 단위 정보)
    "ParcelNumber", "TotalParcel", "AreaHasParcel", "MTMD_ID",
    # 공통 메타
    "원본OBJECTID", "last_edited_date",
]

NEW_PASSTHROUGH = (
    "docTenementNum", "tenementType", "tenementName", "docRegOfc",
    "docProvName", "docMuniName", "totalAreaHas", "miningStage",
    "tenementStatus", "miningPhase", "operationStatus",
    "companyRepresentative", "companyAddress", "companyNumber",
    "genType", "tenementMineral", "Remarks",
)


def fetch_json(url: str) -> dict:
    result = subprocess.run(
        ["curl", "-sL", "--ssl-no-revoke", url],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        print(f"내려받기 실패(curl rc={result.returncode}): {url}", file=sys.stderr)
        sys.exit(1)
    try:
        data = json.loads(result.stdout)
    except json.JSONDecodeError:
        print(f"JSON 파싱 실패: {url}\n{result.stdout[:300]}", file=sys.stderr)
        sys.exit(1)
    if "error" in data:
        print(f"서비스 오류 응답: {data['error']}", file=sys.stderr)
        sys.exit(1)
    return data


def fetch_all_features(service: str) -> list[dict]:
    """maxRecordCount를 넘으면 잘려 오므로 offset을 넘겨가며 전량 받는다."""
    features, offset = [], 0
    while True:
        query = urlencode({
            "where": "1=1", "outFields": "*", "returnGeometry": "false",
            "resultOffset": offset, "resultRecordCount": PAGE_SIZE, "f": "json",
        })
        data = fetch_json(f"{REST}/{service}/FeatureServer/0/query?{query}")
        page = data.get("features", [])
        features.extend(page)
        if len(page) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return [ft["attributes"] for ft in features]


def fetch_count(service: str) -> int:
    query = urlencode({"where": "1=1", "returnCountOnly": "true", "f": "json"})
    return fetch_json(f"{REST}/{service}/FeatureServer/0/query?{query}").get("count")


def to_date(val):
    """epoch ms -> YYYY-MM-DD"""
    if val is None or val == "":
        return ""
    try:
        return datetime.fromtimestamp(val / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OSError):
        return val


def norm_num(s):
    return (s or "").strip().upper().replace(" ", "")


def region_from_number(tenement_number):
    m = re.search(r"-([A-Za-z0-9]+)$", (tenement_number or "").strip())
    if not m:
        return ""
    code = m.group(1).upper()
    return code if code in PH_REGIONS else ""


def row_from_new(a, dup):
    row = {c: "" for c in COLUMNS}
    row.update({
        "출처": "신", "중복여부": dup,
        "DateFiled": to_date(a.get("DateFiled")),
        "StatusDate": to_date(a.get("StatusDate")),
        "원본OBJECTID": a.get("OBJECTID"),
        "last_edited_date": to_date(a.get("last_edited_date")),
    })
    for f in NEW_PASSTHROUGH:
        row[f] = a.get(f) if a.get(f) is not None else ""
    return row


def row_from_old(a, dup):
    """구 서비스 필드를 신 스키마 이름으로 옮긴다.
    주의: AreaHasParcel은 '필지 면적'이라 신 서비스의 totalAreaHas(광구 전체 면적)와
    의미가 다르다. 같은 칸에 넣으면 합계가 틀어지므로 별도 컬럼으로 유지한다."""
    num = a.get("TenementNumber") or ""
    row = {c: "" for c in COLUMNS}
    row.update({
        "출처": "구", "중복여부": dup,
        "docTenementNum": num,
        "tenementType": a.get("TenementType") or "",
        "tenementName": a.get("TenementHolder") or "",
        "docRegOfc": region_from_number(num),
        "ParcelNumber": a.get("ParcelNumber") if a.get("ParcelNumber") is not None else "",
        "TotalParcel": a.get("TotalParcel") if a.get("TotalParcel") is not None else "",
        "AreaHasParcel": a.get("AreaHasParcel") if a.get("AreaHasParcel") is not None else "",
        "MTMD_ID": a.get("MTMD_ID") or "",
        "원본OBJECTID": a.get("OBJECTID"),
        "last_edited_date": to_date(a.get("last_edited_date")),
    })
    return row


def main():
    print("[내려받기] 서비스 2개 조회 중...")
    expected_old, expected_new = fetch_count(SERVICE_OLD), fetch_count(SERVICE_NEW)
    old = fetch_all_features(SERVICE_OLD)
    new = fetch_all_features(SERVICE_NEW)

    old_nums = {norm_num(a.get("TenementNumber")) for a in old}
    new_nums = {norm_num(a.get("docTenementNum")) for a in new}
    both = old_nums & new_nums

    rows = [row_from_old(a, "양쪽" if norm_num(a.get("TenementNumber")) in both else "단독")
            for a in old]
    rows += [row_from_new(a, "양쪽" if norm_num(a.get("docTenementNum")) in both else "단독")
             for a in new]

    with open(OUTPUT, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    src = Counter(r["출처"] for r in rows)
    dup = Counter(r["중복여부"] for r in rows)
    region_ok = sum(1 for r in rows if r["출처"] == "구" and r["docRegOfc"])

    print(f"[출력] {OUTPUT}")
    print()
    print("=== 건수 검증 (서비스가 알려준 건수 vs 실제 받은 행수) ===")
    for label, got, expected in (
        (f"구 {SERVICE_OLD}", src["구"], expected_old),
        (f"신 {SERVICE_NEW}", src["신"], expected_new),
    ):
        tag = "[PASS]" if got == expected else "[FAIL]"
        print(f"  {tag} {label} : {got:,} vs {expected:,}")
    print()
    print(f"총 행수: {len(rows):,}행")
    print(f"겹치는 광구번호: {len(both)}개 -> 양쪽 {dup['양쪽']}행 / 단독 {dup['단독']}행")
    print(f"구 데이터 지역코드 추출: {region_ok}/{src['구']}행 "
          f"(나머지는 광구번호 접미사가 지역코드가 아니라 공란)")


if __name__ == "__main__":
    main()
