# -*- coding: utf-8 -*-
"""
MGB(필리핀 광업지질국) 공개 지리정보 다운로더 — Streamlit 웹앱.

ArcGIS Experience 대시보드는 화면을 자바스크립트로 그리기 때문에 페이지를 그대로
긁으면 빈 껍데기만 나온다. 대신 앱 정의 -> 웹맵 정의 -> 피처 레이어 순으로 거슬러
올라가 실제 데이터 API를 찾아 받는다.

실행:  streamlit run app.py
"""
from __future__ import annotations

import csv
import io
import json
import re
import subprocess
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from urllib.parse import urlencode, urlparse

import streamlit as st

# ──────────────────────────────────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────────────────────────────────

APP_ID = "00a0cfe6bcf94fa4a3c36a0b742b2b45"
PORTAL = "https://www.arcgis.com"
DIRECTORY = "https://services7.arcgis.com/Z0dvtKpPYjB1vNXq/arcgis/rest/services"

# [안전] 이 호스트 외에는 요청하지 않는다. 앱 정의에 낯선 URL이 섞여 들어와도 차단된다.
ALLOWED_HOSTS = {"www.arcgis.com", "services7.arcgis.com", "controlmap.mgb.gov.ph"}

# [안전] 수집 상한. 실수로 대용량을 끌어와 서버에 부담을 주는 것을 막는다.
MAX_PER_LAYER = 50_000
MAX_PER_RUN = 100_000

# [견고성] 재시도 정책 (429 / 5xx 대상)
MAX_ATTEMPTS = 4
BACKOFF_BASE = 1.5          # 초. 시도마다 2배씩 늘어난다.
HTTP_TIMEOUT = 60

DEFAULT_INTERVAL = 0.3      # 초. 요청 사이 최소 간격.
FALLBACK_PAGE_SIZE = 200    # 서비스가 maxRecordCount를 안 알려줄 때만 쓰는 보수적 기본값.

# 엑셀 수식 실행을 유발할 수 있는 시작 문자 (CSV 인젝션 방어)
FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

# 피처 조회가 가능한 레이어 타입. 타일 지도 서비스는 그림만 주므로 제외한다.
QUERYABLE_TYPES = {"ArcGISFeatureLayer", "ArcGISMapServiceLayer", None, ""}


# ──────────────────────────────────────────────────────────────────────────
# HTTP — 호스트 검증 · 재시도 · requests 우선 / curl 폴백
# ──────────────────────────────────────────────────────────────────────────

class FetchError(Exception):
    pass


def _check_host(url: str) -> None:
    host = (urlparse(url).hostname or "").lower()
    if host not in ALLOWED_HOSTS:
        raise FetchError(f"허용되지 않은 호스트라 요청하지 않았습니다: {host or url}")


def _get_requests():
    """requests가 있으면 반환. 없으면 None (curl로 폴백)."""
    try:
        import requests  # noqa: PLC0415
        return requests
    except ImportError:
        return None


def _http_once(url: str) -> tuple[int, str]:
    """한 번 요청하고 (상태코드, 본문)을 돌려준다. 리다이렉트는 따라가지 않는다.

    사내망처럼 SSL 검사 장비가 낀 환경에서는 파이썬의 인증서 검증이 실패한다.
    그런 경우에만 curl(--ssl-no-revoke)로 넘어간다.
    """
    requests = _get_requests()
    if requests is not None:
        try:
            resp = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=False)
            return resp.status_code, resp.text
        except requests.exceptions.SSLError:
            pass  # 아래 curl 폴백으로
        except requests.exceptions.RequestException as exc:
            raise FetchError(f"요청 실패: {exc}") from exc

    # curl 폴백 (-L 없음 = 리다이렉트 미추적)
    try:
        proc = subprocess.run(
            ["curl", "-s", "--ssl-no-revoke", "--max-time", str(HTTP_TIMEOUT),
             "-w", "\n%{http_code}", url],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
    except FileNotFoundError as exc:
        raise FetchError(
            "requests의 SSL 검증이 실패했고 curl도 찾을 수 없습니다. "
            "requests를 설치하거나 curl을 PATH에 추가해 주세요."
        ) from exc
    if proc.returncode != 0:
        raise FetchError(f"curl 실패(rc={proc.returncode}): {proc.stderr.strip()[:200]}")
    body, _, code = proc.stdout.rpartition("\n")
    return int(code or 0), body


def fetch_json(url: str, params: dict | None = None, interval: float = DEFAULT_INTERVAL) -> dict:
    """JSON을 받아 온다. 429/5xx는 백오프하며 재시도하고, 4xx는 즉시 포기한다."""
    if params:
        url = f"{url}?{urlencode(params)}"
    _check_host(url)

    last = ""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        time.sleep(interval)
        status, body = _http_once(url)

        if status == 200:
            try:
                data = json.loads(body)
            except json.JSONDecodeError as exc:
                raise FetchError(f"JSON이 아닌 응답입니다: {body[:200]}") from exc
            # ArcGIS는 오류도 200으로 내려보내는 경우가 있다.
            # message가 비어 있고 정작 쓸모 있는 내용은 details에 들어오는 일이 잦다
            # (예: WHERE에 없는 필드를 쓰면 "'Invalid field: xxx' parameter is invalid").
            if isinstance(data, dict) and "error" in data:
                err = data["error"]
                detail = "; ".join(err.get("details") or [])
                reason = err.get("message") or detail or str(err)
                if detail and err.get("message"):
                    reason = f"{err['message']} ({detail})"
                raise FetchError(f"서비스 오류 {err.get('code', '')}: {reason}")
            return data

        if status in (429, 500, 502, 503, 504):
            last = f"HTTP {status}"
            if attempt < MAX_ATTEMPTS:
                time.sleep(BACKOFF_BASE * (2 ** (attempt - 1)))
                continue
        elif 300 <= status < 400:
            raise FetchError(f"리다이렉트(HTTP {status})는 따라가지 않습니다.")
        else:
            raise FetchError(f"HTTP {status}: {body[:200]}")

    raise FetchError(f"{MAX_ATTEMPTS}회 재시도 후에도 실패했습니다 ({last}).")


# ──────────────────────────────────────────────────────────────────────────
# 레이어 탐색
# ──────────────────────────────────────────────────────────────────────────

@dataclass
class Layer:
    """다운로드 대상 하나. 대시보드에서 같은 URL을 여러 이름으로 쓰는 경우가 많아
    URL 기준으로 합치고, 화면에서 색을 나누던 기준은 subcats로 남겨 둔다."""
    url: str
    title: str
    source: str                      # "대시보드" 또는 "저장소"
    layer_type: str = ""
    aliases: list[str] = field(default_factory=list)
    # (필드명, 값) — 예: ("Commodity", "Tuff"). 세부 항목만 골라 받을 때 WHERE로 바뀐다.
    subcats: list[tuple[str, str]] = field(default_factory=list)

    # st.cache_data는 반환값을 pickle로 저장하는데, 스크립트가 다시 실행될 때마다
    # 이 클래스가 새로 정의되어 "not the same object as __main__.Layer" 오류가 난다.
    # 그래서 캐시를 거치는 구간에서는 평범한 dict로 주고받는다.
    def to_dict(self) -> dict:
        return {
            "url": self.url, "title": self.title, "source": self.source,
            "layer_type": self.layer_type, "aliases": list(self.aliases),
            "subcats": [tuple(s) for s in self.subcats],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Layer":
        return cls(**d)

    @property
    def key(self) -> str:
        return self.url

    @property
    def queryable(self) -> bool:
        return self.layer_type in QUERYABLE_TYPES

    @property
    def slug(self) -> str:
        """파일명으로 쓸 안전한 이름. 서비스명/레이어번호에서 만든다."""
        m = re.search(r"/services/(.+?)/(?:Feature|Map)Server/(\d+)", self.url)
        base = f"{m.group(1)}_{m.group(2)}" if m else self.title
        return re.sub(r"[^0-9A-Za-z가-힣_.-]+", "_", base).strip("_") or "layer"


def _renderer_subcats(entry: dict) -> list[tuple[str, str]]:
    """대시보드가 uniqueValue 렌더러로 색을 나눈 기준을 (필드, 값)으로 뽑는다.
    서버측 definitionExpression이 아니라 표시용 구분이므로, 이걸 WHERE로 바꿔야
    화면에서 보던 것과 같은 부분집합을 받을 수 있다."""
    renderer = (entry.get("layerDefinition") or {}).get("drawingInfo", {}).get("renderer", {})
    field_name = renderer.get("field1")
    if renderer.get("type") != "uniqueValue" or not field_name:
        return []
    out = []
    for info in renderer.get("uniqueValueInfos") or []:
        value = info.get("value")
        if value not in (None, ""):
            out.append((field_name, str(value)))
    return out


@st.cache_data(show_spinner=False)
def discover_dashboard_layers(app_id: str, interval: float) -> tuple[list[dict], list[str]]:
    """앱 정의 -> 웹맵 -> operationalLayers(그룹 재귀) 순으로 실제 레이어를 찾는다.
    캐시에 담기려면 pickle이 되어야 하므로 Layer가 아니라 dict 목록을 돌려준다."""
    notes: list[str] = []
    app = fetch_json(f"{PORTAL}/sharing/rest/content/items/{app_id}/data",
                     {"f": "json"}, interval)

    webmap_ids: list[str] = []
    for src in (app.get("dataSources") or {}).values():
        item_id = src.get("itemId")
        if src.get("type") == "WEB_MAP" and item_id and item_id not in webmap_ids:
            webmap_ids.append(item_id)
    if not webmap_ids:
        notes.append("앱 정의에서 웹맵을 찾지 못했습니다.")

    by_url: dict[str, Layer] = {}

    def walk(entries: list[dict]) -> None:
        for entry in entries:
            children = entry.get("layers")
            if children:                       # 그룹 레이어 -> 안쪽까지 재귀
                walk(children)
                continue
            url = entry.get("url")
            if not url:
                continue
            title = entry.get("title") or entry.get("id") or url
            layer = by_url.get(url)
            if layer is None:
                layer = Layer(url=url, title=title, source="대시보드",
                              layer_type=entry.get("layerType") or "")
                by_url[url] = layer
            elif title not in layer.aliases and title != layer.title:
                layer.aliases.append(title)
            for sub in _renderer_subcats(entry):
                if sub not in layer.subcats:
                    layer.subcats.append(sub)

    for wid in webmap_ids:
        try:
            webmap = fetch_json(f"{PORTAL}/sharing/rest/content/items/{wid}/data",
                                {"f": "json"}, interval)
            walk(webmap.get("operationalLayers") or [])
        except FetchError as exc:
            notes.append(f"웹맵 {wid} 를 읽지 못했습니다: {exc}")

    # 이름이 여러 개로 합쳐진 레이어는 대표 이름을 서비스명으로 바꿔 준다.
    for layer in by_url.values():
        if layer.aliases:
            m = re.search(r"/services/(.+?)/(?:Feature|Map)Server/(\d+)", layer.url)
            if m:
                layer.title = f"{m.group(1)} (레이어 {m.group(2)})"

    layers = sorted(by_url.values(), key=lambda x: x.title.lower())
    return [x.to_dict() for x in layers], notes


@st.cache_data(show_spinner=False)
def discover_directory_layers(interval: float, _report=None) -> tuple[list[dict], list[str]]:
    """저장소(services 디렉터리)에 공개된 서비스를 모두 열거한다.
    대시보드에 걸리지 않은 서비스도 여기서 보인다.

    _report는 진행률 콜백. 앞의 밑줄은 st.cache_data가 해시 대상에서 빼라는 표시다."""
    notes: list[str] = []
    listing = fetch_json(DIRECTORY, {"f": "json"}, interval)
    services = [s for s in listing.get("services", []) if s.get("type") == "FeatureServer"]

    layers: list[Layer] = []
    for i, svc in enumerate(services, start=1):
        name = svc["name"].split("/")[-1]
        if _report:
            _report(i / len(services), f"저장소 탐색 중… ({i}/{len(services)}) {name}")
        try:
            root = fetch_json(f"{DIRECTORY}/{name}/FeatureServer", {"f": "json"}, interval)
        except FetchError as exc:
            notes.append(f"{name}: {exc}")
            continue
        for sub in (root.get("layers") or []) + (root.get("tables") or []):
            layers.append(Layer(
                url=f"{DIRECTORY}/{name}/FeatureServer/{sub['id']}",
                title=f"{name} / {sub.get('name')} (레이어 {sub['id']})",
                source="저장소",
                layer_type="ArcGISFeatureLayer",
            ))
    return [x.to_dict() for x in sorted(layers, key=lambda x: x.title.lower())], notes


# ──────────────────────────────────────────────────────────────────────────
# 데이터 수집
# ──────────────────────────────────────────────────────────────────────────

def combine_where(base: str, subcat: tuple[str, str] | None) -> str:
    base = (base or "").strip() or "1=1"
    if not subcat:
        return base
    field_name, value = subcat
    escaped = value.replace("'", "''")          # SQL 작은따옴표 이스케이프
    clause = f"{field_name} = '{escaped}'"
    return clause if base == "1=1" else f"({base}) AND {clause}"


def layer_metadata(url: str, interval: float) -> dict:
    return fetch_json(url, {"f": "json"}, interval)


def count_features(url: str, where: str, interval: float) -> int:
    data = fetch_json(f"{url}/query",
                      {"where": where, "returnCountOnly": "true", "f": "json"}, interval)
    return int(data.get("count", 0))


def fetch_features(url: str, where: str, page_size: int, want_geometry: bool,
                   interval: float, budget: int, on_page=None) -> tuple[list[dict], bool]:
    """페이지를 넘겨가며 전량 수집한다. budget(잔여 허용 건수)에 걸리면 중단한다.

    반환: (features, truncated) — truncated는 상한 때문에 잘렸는지 여부.
    want_geometry면 f=geojson(도형+속성), 아니면 f=json(속성만)으로 받는다.
    """
    fmt = "geojson" if want_geometry else "json"
    features: list[dict] = []
    offset = 0
    while True:
        remaining = budget - len(features)
        if remaining <= 0:
            return features, True
        params = {
            "where": where,
            "outFields": "*",
            "returnGeometry": "true" if want_geometry else "false",
            "resultOffset": offset,
            "resultRecordCount": min(page_size, remaining),
            "f": fmt,
        }
        data = fetch_json(f"{url}/query", params, interval)
        page = data.get("features") or []
        features.extend(page)
        if on_page:
            on_page(len(features))
        if len(page) < params["resultRecordCount"]:
            return features, False
        offset += len(page)


# ──────────────────────────────────────────────────────────────────────────
# 내보내기
# ──────────────────────────────────────────────────────────────────────────

def sanitize_cell(value):
    """[안전] 엑셀/스프레드시트 수식 인젝션 차단.

    =, +, -, @ 등으로 시작하는 '문자열'은 앞에 작은따옴표를 붙여 수식이 아닌
    텍스트로 읽히게 한다. 숫자형(-45000000 같은 음수)은 그대로 두어야 계산이
    깨지지 않으므로 문자열일 때만 적용한다.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if not isinstance(value, str):
        return value
    return "'" + value if value.startswith(FORMULA_PREFIXES) else value


def epoch_to_date(value):
    if value in (None, ""):
        return ""
    try:
        return datetime.fromtimestamp(value / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OSError):
        return value


def build_csv(features: list[dict], fields: list[dict], convert_dates: bool) -> bytes:
    """원본 필드 순서를 그대로 유지한 CSV. 엑셀용으로 UTF-8 BOM을 붙인다."""
    names = [f["name"] for f in fields]
    date_fields = {f["name"] for f in fields if f.get("type") == "esriFieldTypeDate"}

    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=names, extrasaction="ignore")
    writer.writeheader()
    for feat in features:
        attrs = feat.get("attributes") or feat.get("properties") or {}
        row = {}
        for name in names:
            val = attrs.get(name)
            if convert_dates and name in date_fields:
                val = epoch_to_date(val)
            row[name] = sanitize_cell(val)
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")


def build_geojson(features: list[dict]) -> bytes:
    fc = {"type": "FeatureCollection", "features": features}
    return json.dumps(fc, ensure_ascii=False, indent=1).encode("utf-8")


def build_zip(files: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, blob in files.items():
            zf.writestr(name, blob)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────
# 화면
# ──────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="MGB 공개데이터 다운로더", page_icon="⛏️", layout="wide")
st.title("⛏️ MGB 공개 지리정보 다운로더")
st.caption(
    "필리핀 광업지질국(MGB) ArcGIS 대시보드의 공개 레이어를 찾아 GeoJSON/CSV로 내려받습니다. "
    "공개된 자료만 다루며, 서버에 부담이 가지 않도록 요청 간격과 수집 상한을 둡니다."
)

with st.sidebar:
    st.header("설정")
    interval = st.slider("요청 간격 (초)", 0.0, 2.0, DEFAULT_INTERVAL, 0.1,
                         help="값이 클수록 서버에 순하지만 느려집니다.")
    convert_dates = st.checkbox("날짜를 YYYY-MM-DD로 변환", value=True,
                                help="원본은 epoch 밀리초입니다. 끄면 원본 숫자를 그대로 씁니다.")
    st.divider()
    st.subheader("안전 한도")
    st.markdown(
        f"- 레이어당 **{MAX_PER_LAYER:,}건**\n"
        f"- 실행당 **{MAX_PER_RUN:,}건**\n"
        f"- 재시도 최대 {MAX_ATTEMPTS}회 (429·5xx)\n"
        "- 리다이렉트 미추적\n"
        f"- 허용 호스트 {len(ALLOWED_HOSTS)}곳"
    )
    st.caption("허용 호스트: " + ", ".join(sorted(ALLOWED_HOSTS)))

tab_dash, tab_dir = st.tabs(["대시보드 레이어", "저장소 전체"])
catalog: list[Layer] = []

with tab_dash:
    st.write(f"앱 ID `{APP_ID}` 의 웹맵을 따라가 레이어를 찾습니다.")
    try:
        dash_raw, dash_notes = discover_dashboard_layers(APP_ID, interval)
        dash_layers = [Layer.from_dict(d) for d in dash_raw]
        catalog.extend(dash_layers)
        usable = [x for x in dash_layers if x.queryable]
        skipped = [x for x in dash_layers if not x.queryable]
        st.success(f"조회 가능한 레이어 {len(usable)}개를 찾았습니다.")
        for note in dash_notes:
            st.warning(note)
        if skipped:
            with st.expander(f"조회 불가 {len(skipped)}개 (타일 지도 서비스 — 지도 그림만 제공)"):
                for x in skipped:
                    st.write(f"• {x.title} — `{x.layer_type}`")
    except FetchError as exc:
        st.error(f"레이어 탐색 실패: {exc}")
        dash_layers = []

with tab_dir:
    st.write(
        "대시보드에 걸리지 않은 서비스까지 저장소 전체를 훑습니다. "
        "**주의:** 이름이 비슷해도 대시보드 것과 내용이 다른 구버전 서비스가 섞여 있습니다."
    )
    if st.button("저장소 탐색하기", help="서비스 수만큼 요청이 나가므로 시간이 걸립니다."):
        st.session_state["dir_done"] = True
    if st.session_state.get("dir_done"):
        try:
            bar = st.progress(0.0, text="저장소 서비스 탐색 중…")
            dir_raw, dir_notes = discover_directory_layers(
                interval, _report=lambda frac, msg: bar.progress(frac, text=msg)
            )
            bar.empty()
            dir_layers = [Layer.from_dict(d) for d in dir_raw]
            catalog.extend(dir_layers)
            st.success(f"저장소 레이어 {len(dir_layers)}개를 찾았습니다.")
            for note in dir_notes[:10]:
                st.warning(note)
        except FetchError as exc:
            st.error(f"저장소 탐색 실패: {exc}")

if not catalog:
    st.stop()

# ── 레이어 선택 ────────────────────────────────────────────────────────────
st.divider()
st.subheader("1. 레이어 선택")

selectable = [x for x in catalog if x.queryable]
by_key = {x.key: x for x in selectable}

select_all = st.checkbox("전체 선택", value=False)
default = [x.key for x in selectable] if select_all else []
chosen_keys = st.multiselect(
    "받을 레이어",
    options=[x.key for x in selectable],
    default=default,
    format_func=lambda k: f"[{by_key[k].source}] {by_key[k].title}",
)

# 같은 URL을 여러 이름으로 쓰던 레이어는 세부 항목을 따로 고를 수 있게 한다.
subcat_choice: dict[str, list[tuple[str, str]]] = {}
splittable = [by_key[k] for k in chosen_keys if by_key[k].subcats]
if splittable:
    with st.expander(f"세부 항목으로 나눠 받기 ({len(splittable)}개 레이어에서 가능)"):
        st.caption(
            "대시보드는 아래 항목들을 한 레이어에서 색만 다르게 칠해 보여줍니다. "
            "고르지 않으면 레이어 전체를 한 파일로 받습니다."
        )
        for layer in splittable:
            picked = st.multiselect(
                layer.title,
                options=[f"{f} = {v}" for f, v in layer.subcats],
                key=f"sub_{layer.key}",
            )
            if picked:
                lookup = {f"{f} = {v}": (f, v) for f, v in layer.subcats}
                subcat_choice[layer.key] = [lookup[p] for p in picked]

# ── 조건과 형식 ────────────────────────────────────────────────────────────
st.subheader("2. 조건과 형식")
col1, col2 = st.columns([3, 2])
with col1:
    where = st.text_input("WHERE 조건", value="1=1",
                          help="SQL WHERE 절. 예: docRegOfc = 'CAR'")
    st.caption("예시 — 전체: `1=1` · 지역: `docRegOfc = 'III'` · "
               "면적: `totalAreaHas > 1000` · 조합: `tenementType = 'MPSA' AND totalAreaHas > 500`")
with col2:
    fmt_label = st.radio("형식", ["CSV", "GeoJSON", "둘 다"], horizontal=True)
want_csv = fmt_label in ("CSV", "둘 다")
want_geojson = fmt_label in ("GeoJSON", "둘 다")

# ── 실행 ──────────────────────────────────────────────────────────────────
st.subheader("3. 내려받기")

if not chosen_keys:
    st.info("레이어를 하나 이상 골라 주세요.")
    st.stop()

# 실제 내려받을 작업 단위(레이어 또는 레이어의 세부 항목)를 펼친다.
jobs: list[tuple[Layer, tuple[str, str] | None]] = []
for key in chosen_keys:
    layer = by_key[key]
    subs = subcat_choice.get(key)
    if subs:
        jobs.extend((layer, s) for s in subs)
    else:
        jobs.append((layer, None))

st.write(f"작업 **{len(jobs)}건** · 형식 **{fmt_label}**")

if st.button("수집 시작", type="primary"):
    outputs: dict[str, bytes] = {}
    failures: list[tuple[str, str]] = []
    warnings: list[str] = []
    used = 0

    progress = st.progress(0.0, text="시작하는 중…")
    log = st.container()

    for idx, (layer, subcat) in enumerate(jobs, start=1):
        label = layer.title + (f" [{subcat[0]}={subcat[1]}]" if subcat else "")
        progress.progress((idx - 1) / len(jobs), text=f"({idx}/{len(jobs)}) {label}")

        if used >= MAX_PER_RUN:
            failures.append((label, f"실행당 상한 {MAX_PER_RUN:,}건에 도달해 건너뜀"))
            continue

        try:
            job_where = combine_where(where, subcat)
            meta = layer_metadata(layer.url, interval)
            page_size = int(meta.get("maxRecordCount") or FALLBACK_PAGE_SIZE)
            fields = meta.get("fields") or []

            # [데이터 품질] 받기 전에 서버가 말하는 건수를 먼저 확인한다.
            expected = count_features(layer.url, job_where, interval)
            if expected == 0:
                warnings.append(f"{label}: 조건에 맞는 자료가 0건입니다.")
                progress.progress(idx / len(jobs))
                continue

            budget = min(MAX_PER_LAYER, MAX_PER_RUN - used, expected)
            if expected > budget:
                warnings.append(
                    f"{label}: {expected:,}건 중 상한에 걸려 {budget:,}건만 받았습니다."
                )

            def tick(done: int, _label=label, _exp=expected, _i=idx):
                progress.progress(
                    min((_i - 1 + done / max(_exp, 1)) / len(jobs), 1.0),
                    text=f"({_i}/{len(jobs)}) {_label} — {done:,}/{min(_exp, budget):,}건",
                )

            features, truncated = fetch_features(
                layer.url, job_where, page_size, want_geojson, interval, budget, tick
            )
            used += len(features)

            # [데이터 품질] 상한 때문이 아닌데도 건수가 어긋나면 알린다.
            if not truncated and len(features) != expected:
                warnings.append(
                    f"{label}: 예상 {expected:,}건과 받은 {len(features):,}건이 다릅니다. "
                    "수집 중 원본이 바뀌었을 수 있습니다."
                )

            stem = layer.slug + (f"_{re.sub(r'[^0-9A-Za-z가-힣]+', '_', subcat[1])}" if subcat else "")
            if want_csv:
                outputs[f"{stem}.csv"] = build_csv(features, fields, convert_dates)
            if want_geojson:
                outputs[f"{stem}.geojson"] = build_geojson(features)

            log.write(f"✅ {label} — {len(features):,}건")

        except FetchError as exc:
            failures.append((label, str(exc)))
            log.write(f"❌ {label} — {exc}")
        except Exception as exc:  # 레이어 하나가 터져도 나머지는 계속 간다
            failures.append((label, f"예상치 못한 오류: {exc}"))
            log.write(f"❌ {label} — {exc}")

        progress.progress(idx / len(jobs))

    progress.progress(1.0, text="완료")
    st.divider()

    ok = len(jobs) - len(failures)
    st.metric("성공 / 전체", f"{ok} / {len(jobs)}")
    st.write(f"수집한 총 건수: **{used:,}건**")

    for msg in warnings:
        st.warning(msg)
    if failures:
        st.error(f"실패 {len(failures)}건")
        st.table({"레이어": [f[0] for f in failures], "사유": [f[1] for f in failures]})

    if not outputs:
        st.info("내려받을 결과가 없습니다.")
    elif len(outputs) == 1:
        name, blob = next(iter(outputs.items()))
        st.download_button(f"⬇️ {name}", blob, file_name=name, type="primary")
    else:
        stamp = datetime.now().strftime("%y%m%d_%H%M")
        st.download_button(
            f"⬇️ 전체 {len(outputs)}개 파일 ZIP으로 받기",
            build_zip(outputs), file_name=f"mgb_{stamp}.zip",
            mime="application/zip", type="primary",
        )
        with st.expander("개별 파일로 받기"):
            for name, blob in outputs.items():
                st.download_button(name, blob, file_name=name, key=f"dl_{name}")
