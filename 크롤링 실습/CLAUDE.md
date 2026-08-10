# ArcGIS 대시보드 데이터 추출

필리핀 광업지질국(MGB)의 ArcGIS Experience 대시보드에서 공개 데이터를 REST API로 받아 CSV로 만드는 작업. 다른 폴더와 달리 입력이 로컬 엑셀이 아니라 원격 서비스다.

- 대시보드: https://experience.arcgis.com/experience/00a0cfe6bcf94fa4a3c36a0b742b2b45
- 데이터 저장소: `https://services7.arcgis.com/Z0dvtKpPYjB1vNXq/arcgis/rest/services` (공개 서비스 55개)

## 명령어

```
python 신청광구_통합.py     # 신청 광구 2개 서비스를 받아 병합 -> 신청광구_통합.csv
```

승인 광구(`승인광구_전체.csv`)는 아직 스크립트가 없다 — 수동 조회로 만든 결과물이다.

## 네트워크

이 PC에서는 Python `urllib`의 HTTPS 인증서 검증이 실패하고(`Missing Authority Key Identifier`, 사내 SSL 검사로 추정), 일반 `curl`도 `CRYPT_E_NO_REVOCATION_CHECK`로 막힌다. **`curl -sL --ssl-no-revoke`로 shell out 하는 방식만 작동한다.** 재진단하느라 시간 쓰지 말 것.

WebFetch로 대시보드 URL을 직접 열면 JS 렌더링 전의 빈 껍데기만 나온다. 데이터는 아래 경로로 앱 정의를 거슬러 올라가 찾는다:

1. 앱 정의: `arcgis.com/sharing/rest/content/items/{앱ID}/data?f=json` → `dataSources`의 웹맵 itemId
2. 웹맵 정의: `.../items/{웹맵ID}/data?f=json` → `operationalLayers`의 url (**GroupLayer는 안쪽 `layers` 배열을 재귀로 펼쳐야 한다** — 4단계까지 중첩됨)
3. 레이어 쿼리: `{레이어url}/query?where=1=1&outFields=*&returnGeometry=false&f=json`

`maxRecordCount`(보통 1000~2000)를 넘으면 잘려 오므로 `resultOffset`으로 페이지네이션할 것. 받은 행수를 `returnCountOnly=true` 결과와 반드시 대조한다.

## 광구(Tenement) 서비스 4개의 관계 — 중요

이름이 비슷한 서비스가 4개 있는데 **신·구 두 계열이고, 서로의 최신판이 아니다.**

| 서비스 | 건수 | 최종수정 | 지역범위 | 스키마 |
|---|---|---|---|---|
| `Tenement_Application` (구 신청) | 1,077 | 2023-11-06 | 전국 ~14개 | 필지 기반, 빈약 |
| `Application_Mining_Tenement` (신 신청) | 181 | 2026-07-23 | **4개 지역뿐** | 상세 |
| `Mining_Tenement_Approved` (구 승인) | 431 | 2023-11-06 | — | 필지 기반 |
| `Approved_Tenements` (신 승인) | 621 | 2026-07-23 | 전국 16개 | 상세 |

대시보드는 **신 계열만** 참조한다. 그런데 신 신청 서비스는 16개 지역 중 4개(I·III·CAR·II)만 들어와 있어 **마이그레이션이 진행 중**인 것으로 보인다. 승인(621건)보다 신청(181건)이 적은 건 실제 현황이 아니라 이 누락 때문이다.

두 신청 서비스는 광구번호가 **61개만 겹친다**(구 단독 1,016개 / 신 단독 68개). 어느 한쪽만 쓰면 정보가 손실되므로, `신청광구_통합.py`는 둘 다 남기고 `출처`(구/신)·`중복여부`(양쪽/단독) 컬럼으로 구분한다. **어느 값이 맞는지는 판단하지 않는다.**

## 데이터 다룰 때 주의

- **면적 필드를 합치지 말 것.** 구 서비스 `AreaHasParcel`은 필지별 면적, 신 서비스 `totalAreaHas`는 광구 전체 면적이라 의미가 다르다. 같은 칸에 넣으면 합계가 틀어진다.
- **광구번호는 유일하지 않다.** 필지 분할로 한 광구가 여러 행에 걸친다 (승인 621행 = 고유번호 457개, 신 신청 181행 = 고유번호 129개). 광구 단위로 세려면 중복 제거가 필요하다.
- 구 서비스에는 지역 컬럼이 없어 광구번호 접미사에서 유추하는데, 접미사에 일련번호·`AMENDED` 등이 섞여 있어 1,077행 중 907행만 추출된다. 나머지는 공란.
- 날짜 필드는 epoch ms(UTC)다. CSV로 낼 때 `YYYY-MM-DD`로 변환한다.
- CSV는 UTF-8 BOM으로 저장한다 (엑셀에서 한글이 깨지지 않도록).

## 진행 현황

- [x] 대시보드에서 숨은 데이터 API 역추적
- [x] 승인 광구 621건 → `승인광구_전체.csv` (수동 조회)
- [x] 신청 광구 신·구 병합 1,258건 → `신청광구_통합.csv` (`신청광구_통합.py`)
- [ ] 승인 광구도 스크립트화 (자료 분석 후 결정하기로 함)
