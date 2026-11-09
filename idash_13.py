"""iCReaT 코드북 데이터 자동 기탁 시스템 (Streamlit)

SOP 흐름 (헤더 단계 버튼과 동일)
  ① 파일 업로드          : iCReaT 과제정보 PDF(Ctrl+P 저장) + eCRF 코드북(xlsx) drag & drop
  ② 변수 매핑       : eCRF 변수명 → CODA 변수명 자동 변환 미리보기, PDF → 과제정보 자동 추출
  ③ 과제·데이터 입력 : PDF 자동 입력값 확인·수정, 임상 변수 공개/비공개, 기탁 정보 입력
  ④ 결과 확인            : 기탁용 임상정보 코드북(xlsx) 다운로드

파일
  pdf_extract.py                              : 과제정보 PDF 텍스트 추출 (OCR 없이 텍스트 레이어 사용)
  style.css                                   : 화면 디자인
  data/CODA_clinical_codebook_template.xlsx   : CODA 기탁_임상정보 코드북 양식 (결과 Excel의 틀)
  data/C170005_ecrf_codebook.xlsx             : 샘플 eCRF 코드북
  data/C170005_project_info_sample.pdf        : 샘플 과제정보 PDF
  table_translate.py                          : 테이블명(영문) → 국문 자동 번역 (오프라인 용어 사전)

큐레이션 규칙
  No. ID·FAMILY_ID·PNTTM = 0, 이후 1부터 연번 / 코드리스트 없으면 빈 셀 유지 /
  통합 설문지 = 변수 설명 + 변수값(코드) / 공개여부 = 3단계 체크박스
"""
import base64
import datetime as dt
import html
import io
import re
from collections import Counter
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment

from pdf_extract import FIELD_LABELS, extract_project_info
from table_translate import translate_table_name

st.set_page_config(page_title="iCReaT 코드북 데이터 자동 기탁 시스템", page_icon="🗄️", layout="wide",
                   initial_sidebar_state="collapsed")
E = html.escape
BASE = Path(__file__).parent
DATA = BASE / "data"
TEMPLATE_FILE = DATA / "CODA_clinical_codebook_template.xlsx"
LOGO_FILE = DATA / "knih_logo.png"  # 질병관리청 국립보건연구원 로고

# ══════════════════════════════════════════════════════════════
# 1. 규칙 / 설정  (CODA 규칙 확정 시 여기만 수정)
# ══════════════════════════════════════════════════════════════


def coda_type(dtype, codelist):
    """iCReaT 데이터 유형 → CODA 변수 타입"""
    if codelist:
        return "범주형"
    return {"Integer": "연속형", "Float": "연속형", "String": "문자형", "Date": "날짜형"}.get(dtype, "검토 필요")


def coda_codes(codelist):
    """'1=Normal|2=Abnormal' → '1=Normal,\\n2=Abnormal' (CODA 양식 표기). 코드리스트가 없으면 빈 셀"""
    parts = [p.strip() for p in (codelist or "").replace(";", "|").split("|") if p.strip()]
    return ",\n".join(parts)


def coda_survey(desc, codes):
    """통합 설문지 = 변수 설명 + 변수값(코드) concat"""
    lines = [f"■ {desc or '-'}"] + [f"   □ {c.rstrip(',')}" for c in codes.split("\n") if c]
    return "\n".join(lines)


KEY_TABLE = ("00_EXAMINEE", "기본 정보")  # CODA 큐레이션 키 변수 (No.0)
KEY_VARS = [("ID", "아이디", "아이디"), ("FAMILY_ID", "가족 아이디", "아이디"), ("PNTTM", "시점", "연속형")]

META_KEYS = ["pid", "title", "org", "pi", "period", "goal", "content", "disease"]  # PDF 자동 입력 대상
INFO_ROWS = [  # CODA '과제정보' 시트 (행, 항목명, key, 출처)
    (4, "과제명", "title", "pdf"), (5, "과제 번호", "pid", "pdf"),
    (6, "과제 수행기관(데이터 기탁기관)", "org", "pdf"), (7, "과제 책임자", "pi", "pdf"),
    (8, "과제 수행기간", "period", "pdf"), (9, "과제 목표", "goal", "pdf"),
    (10, "과제 주요내용", "content", "pdf"), (11, "연구 질환명", "disease", "pdf"),
    (12, "임상변수 리스트", "var_list", "user"), (13, "임상변수 개수", "var_count", "calc"),
    (14, "대상자수(샘플수)", "n_subjects", "user"), (15, "데이터 기탁 담당자/연락처", "mgr", "user"),
    (16, "데이터 기탁일", "date", "user"), (17, "데이터 설명(특이사항)", "desc", "user"),
    (18, "CODA 데이터 기탁 신청자/아이디/연락처(기탁담당자와 다른 경우)", "coda", "user"),
]
STEPS = [("upload", "파일 업로드"), ("mapping", "변수 매핑"), ("form", "과제 및 데이터 입력 폼"), ("result", "결과 확인")]

# ══════════════════════════════════════════════════════════════
# 2. eCRF 코드북 파싱
# ══════════════════════════════════════════════════════════════
ECRF_COLS = ["eCRF", "항목명", "변수명", "항목 유형", "데이터 유형", "길이", "단위", "코드리스트"]
REQ_COLS = [  # (내부 컬럼, 인식 별칭, CODA 대상 필드)
    ("eCRF", ("ecrf명", "ecrf", "crf명"), "테이블명(영문·국문)"),
    ("항목명", ("항목명", "itemname"), "변수 설명"),
    ("변수명", ("변수명", "variable"), "변수명(영문)"),
    ("항목 유형", ("항목유형", "itemtype"), "— (참고)"),
    ("데이터 유형", ("데이터유형", "datatype"), "변수 타입"),
    ("길이", ("길이", "length"), "— (참고)"),
    ("단위", ("단위", "unit"), "— (참고)"),
    ("코드리스트", ("코드리스트", "codelist"), "변수값(코드) 설명"),
]


def _match(h):
    n = re.sub(r"\s+", "", str(h or "")).lower()
    for key, alias, _ in REQ_COLS:
        if n and any(n.startswith(a) for a in alias):
            return key
    return None


@st.cache_data(show_spinner=False)
def parse_ecrf_file(data, name):
    """업로드 파일(xlsx/csv) → eCRF 행 목록 + 컬럼 인식 결과"""
    tables = []
    try:
        if name.lower().endswith(".csv"):
            df = None
            for enc in ("utf-8-sig", "cp949"):
                try:
                    df = pd.read_csv(io.BytesIO(data), dtype=str, encoding=enc, header=None).fillna("")
                    break
                except UnicodeDecodeError:
                    continue
            tables.append(("CSV", df.values.tolist() if df is not None else []))
        else:
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for ws in wb.worksheets:
                tables.append((ws.title, [list(r) for r in ws.iter_rows(max_row=10000, values_only=True)]))
    except Exception as e:
        return dict(ok=False, rows=[], error=f"파일을 읽지 못했습니다: {e}", name=name)
    best = None
    for sheet, rows in tables:  # 헤더 행 자동 탐지
        for hi, row in enumerate(rows[:10]):
            hit = {}
            for ci, h in enumerate(row or []):
                k = _match(h)
                if k and k not in hit:
                    hit[k] = (ci, str(h).replace("\n", " ").strip())
            if not best or len(hit) > len(best["hit"]):
                best = dict(sheet=sheet, hi=hi, hit=hit, rows=rows)
    if not best or len(best["hit"]) < 3:
        return dict(ok=False, rows=[], name=name, error="eCRF 코드북 헤더(eCRF명·항목명·변수명 등)를 찾지 못했습니다.")
    out = []
    for r in best["rows"][best["hi"] + 1:]:
        if not r or not any(v not in (None, "") for v in r):
            continue
        rec = {c: "" for c in ECRF_COLS}
        for k, (ci, _) in best["hit"].items():
            v = r[ci] if ci < len(r) else None
            rec[k] = "" if v is None else str(v).strip()
        out.append(rec)
    missing = [k for k, _, _ in REQ_COLS if k not in best["hit"]]
    m = re.search(r"([A-Z]\d{6})", name)
    ok = bool(out) and not ({"eCRF", "항목명", "변수명"} & set(missing))
    return dict(ok=ok, rows=out, sheet=best["sheet"], hit=best["hit"], missing=missing, name=name,
                pid=m.group(1) if m else "", error="" if out else "데이터 행이 없습니다.")


@st.cache_data(show_spinner=False)
def parse_pdf(data, name):
    r = extract_project_info(data)
    r["name"] = name
    return r


# ══════════════════════════════════════════════════════════════
# 3. 상태
# ══════════════════════════════════════════════════════════════
ss = st.session_state
for k, v in dict(stage="upload", page="mapping", forms={}, saved={}, ever={}, dels={}, use_sample=False,
                 upload=None).items():
    ss.setdefault(k, v)


def auto_table_ko(ecrf):
    return {n: translate_table_name(n) for n in dict.fromkeys(r["eCRF"] for r in ecrf)}


def new_form(meta, ecrf, sample=False):
    f = dict(meta={k: str(meta.get(k, "") or "") for k in META_KEYS}, pdf_meta={k: str(meta.get(k, "") or "") for k in META_KEYS},
             vars={i: False for i in range(len(ecrf))}, table_ko=auto_table_ko(ecrf), n_subjects=int(meta["n_subjects"]) if str(meta.get("n_subjects", "")).isdigit() else 0,
             mgr="", mgr_tel="", date=dt.date.today() + dt.timedelta(days=30), desc="", same=True,
             coda_name="", coda_id="", coda_tel="")
    if sample:  # 샘플 체험용 예시값 (생년월일·등록번호 등 식별 가능 변수는 비공개)
        f["vars"] = {i: not (r["변수명"].startswith(("BIRTHD", "CRCRA")) or r["항목명"] == "생년월일" or r["eCRF"] == "AE")
                     for i, r in enumerate(ecrf)}
        f.update(mgr="김연구", mgr_tel="02-1234-5678", desc="신장 이식 후 1년 추적 자료 포함. 결측은 99999로 코딩.",
                 coda_id="kimresearch01")
    return f


def start_mapping(ep, pp, sample=False):
    meta = dict(pp["values"]) if pp else {}
    pid = meta.get("pid") or ep.get("pid") or "UPLOAD-" + dt.date.today().strftime("%y%m%d")
    meta["pid"] = pid
    ss.upload = dict(pid=pid, rows=ep["rows"], ecrf_name=ep["name"], pdf_name=pp["name"] if pp else "",
                     pdf_found=pp.get("found", {}) if pp else {}, pdf_scanned=bool(pp and pp.get("scanned")))
    ss.forms[pid] = new_form(meta, ep["rows"], sample)
    ss.ever[pid] = ss.ever.get(pid, 0) + 1
    ss.stage, ss.page = "main", "mapping"


# ══════════════════════════════════════════════════════════════
# 4. eCRF → CODA 변환 (매핑 + 큐레이션)
# ══════════════════════════════════════════════════════════════
KEY_NAMES = {k for k, _, _ in KEY_VARS}
HANGUL = re.compile(r"[ㄱ-ㅎㅏ-ㅣ가-힣]")


def lang_issues(en_table, ko_table, en_var):
    """CODA 컬럼명 언어 규칙: (영문) 컬럼 = 영문만, (국문) 컬럼 = 한글 포함"""
    out = []
    if en_table and HANGUL.search(en_table):
        out.append("테이블명(영문)에 한글 포함")
    if ko_table and not HANGUL.search(ko_table):
        out.append("테이블명(국문)이 한글이 아님")
    if en_var and HANGUL.search(en_var):
        out.append("변수명(영문)에 한글 포함")
    return out


def build_coda(ecrf, public_map, table_ko):
    """eCRF → CODA 임상정보 코드북 행
    - No.: ID·FAMILY_ID·PNTTM만 0, 나머지 변수는 1부터 연번 (eCRF에 같은 이름이 있으면 키 변수로 흡수)
    - 변수값(코드) 설명: 코드리스트 변환 (없으면 빈 셀 그대로 — 해당없음 여부는 자동으로 알 수 없음)
    - 통합 설문지: 변수 설명 + 변수값(코드) concat
    - 테이블명(국문): 테이블명(영문) 자동 번역 (table_ko, 2단계에서 수정 가능)
    """
    dup = {k for k, v in Counter(r["변수명"] for r in ecrf if r["변수명"]).items() if v > 1}
    rows = []
    for k, d, t in KEY_VARS:
        codes = ""
        rows.append({"No.": 0, "공개여부": "공개", "테이블명(영문)": KEY_TABLE[0], "테이블명(국문)": KEY_TABLE[1],
                     "변수명(영문)": k, "변수 설명": d, "변수값(코드) 설명": codes, "변수 타입": t,
                     "통합 설문지": coda_survey(d, codes), "_src": "큐레이션", "_status": "key", "_issues": [], "_cur": []})
    no = 0
    for i, r in enumerate(ecrf):
        if r["변수명"].upper() in KEY_NAMES:  # eCRF에 키 변수가 있으면 No.0 키 변수로 흡수
            continue
        no += 1
        issues, cur = [], []
        if not r["변수명"]:
            issues.append("변수명 누락")
        elif r["변수명"] in dup:
            issues.append("변수명 중복")
        typ = coda_type(r["데이터 유형"], r["코드리스트"])
        if typ == "검토 필요":
            issues.append("타입 미정")
        tk = table_ko.get(r["eCRF"], {})
        if tk.get("status") == "검토":
            issues.append("테이블 국문명 확인")
        issues += lang_issues(tk.get("en") or r["eCRF"], tk.get("ko", ""), r["변수명"])
        codes = coda_codes(r["코드리스트"])
        if codes:
            cur.append("코드 표기 변환")
        pub = public_map.get(i, False)
        rows.append({"No.": no, "공개여부": "공개" if pub else "비공개",
                     "테이블명(영문)": tk.get("en") or r["eCRF"], "테이블명(국문)": tk.get("ko", ""), "_ecrf": r["eCRF"],
                     "변수명(영문)": r["변수명"], "변수 설명": r["항목명"], "변수값(코드) 설명": codes, "변수 타입": typ,
                     "통합 설문지": coda_survey(r["항목명"], codes), "_src": f'{r["eCRF"]}.{r["변수명"] or "?"}',
                     "_status": "rev" if issues else ("cln" if cur else "map"), "_issues": issues, "_cur": cur})
    return rows


def dist(rows):
    """상태 집계 (코드북 표의 '상태'와 같은 기준): 큐레이션(key) / 매핑 완료(map) / 누락(rev)"""
    n = max(len(rows), 1)
    c = Counter("map" if r["_status"] in ("map", "cln") else r["_status"] for r in rows)
    pct = {k: round(c[k] * 100 / n) for k in ("key", "map", "rev")}
    return pct, round((c["key"] + c["map"]) * 100 / n), c


CODA_COLS = ["No.", "공개여부", "테이블명(영문)", "테이블명(국문)", "변수명(영문)", "변수 설명", "변수값(코드) 설명", "변수 타입", "통합 설문지"]
CODA_SRC = {"No.": "큐레이션", "공개여부": "← 연구자 입력", "테이블명(영문)": "← eCRF명", "테이블명(국문)": "← 영문명 자동 번역",
            "변수명(영문)": "← 변수명", "변수 설명": "← 항목명", "변수값(코드) 설명": "← 코드리스트",
            "변수 타입": "← 데이터 유형", "통합 설문지": "← 설명 + 코드 concat"}


# ══════════════════════════════════════════════════════════════
# 5. 공통 UI
# ══════════════════════════════════════════════════════════════
def icon(name, size=20, color="currentColor", sw=1.8):
    p = {
        "db": '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.66 3.58 3 8 3s8-1.34 8-3V5"/><path d="M4 12c0 1.66 3.58 3 8 3s8-1.34 8-3"/>',
        "server": '<rect x="3" y="3" width="18" height="8" rx="2"/><rect x="3" y="13" width="18" height="8" rx="2"/><path d="M7 7h.01M7 17h.01"/>',
        "layers": '<path d="m12 2 9 5-9 5-9-5 9-5z"/><path d="m3 12 9 5 9-5"/><path d="m3 17 9 5 9-5"/>',
        "list": '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
        "info": '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
        "check": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
        "alert": '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
        "pie": '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
        "file": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/>',
        "pdf": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 15h1.5a1.5 1.5 0 0 0 0-3H8v5M13 12v5h1a2 2 0 0 0 2-2v-1a2 2 0 0 0-2-2z"/>',
        "lock": '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
        "edit": '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
        "arrow": '<path d="M5 12h14M13 6l6 6-6 6"/>',
        "table": '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
        "send": '<path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4z"/>',
        "user": '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
        "clip": '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>',
        "upload": '<path d="M12 16V4M7 9l5-5 5 5"/><path d="M20 16v3a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-3"/>',
        "download": '<path d="M12 4v12M7 11l5 5 5-5"/><path d="M20 16v3a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-3"/>',
    }[name]
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="{color}" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round" '
            f'style="flex-shrink:0;vertical-align:middle">{p}</svg>')


def md(s):
    st.markdown(s, unsafe_allow_html=True)


md("<style>" + (BASE / "style.css").read_text(encoding="utf-8") + "</style>")
md("""<style>
/* 로고 */
.brand-logo{height:52px;width:auto;display:block;flex-shrink:0;padding-right:16px;border-right:1px solid #e5e7eb;}
/* 3. 입력 폼 섹션 헤더 색 통일 */
.sect.auto,.sect.user{background:#eef4fb !important;border:1px solid #c7d8ee !important;color:#0b1f3a !important;}
/* 4. 변수 개수 박스 안에 전체 공개/비공개 버튼 */
.st-key-varbulk{background:#eef4fb;border:1px solid #c7d8ee;border-radius:12px;padding:8px 14px 8px 4px;}
.st-key-varbulk .cnt{background:transparent !important;border:none !important;padding:4px 14px !important;}
.st-key-varbulk button{min-height:38px !important;}
.st-key-varbulk [data-testid="stHorizontalBlock"]{align-items:center !important;}
/* 5. KPI 카드 반투명 */
.k2.hero{background:linear-gradient(135deg,rgba(0,76,151,.80) 0%,rgba(0,51,102,.80) 100%) !important;
  backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);border:1px solid rgba(255,255,255,.35) !important;}
/* 1. 파일이 모두 올라오면 업로드 영역 숨김 */
.st-key-upzone_done [data-testid="stFileUploader"]{display:none !important;}
.st-key-reupbtn button{border-radius:10px !important;border:1px dashed #9dbbe3 !important;color:#004c97 !important;background:#fff !important;}
.legend-row.sub{border-top:1px dashed #e5e7eb;padding-top:9px;margin-top:12px;}
/* 헤더 단계 표시: 파랑 → 빨강 (로고 태극 색) */
.st-key-stepnav{background:linear-gradient(#ffffff,#ffffff) padding-box,
  linear-gradient(90deg,#004c97 0%,#5b6fb5 55%,#e6002e 100%) border-box !important;
  border:1.5px solid transparent !important;border-radius:14px !important;padding:6px 10px !important;
  box-shadow:0 6px 18px -10px rgba(0,76,151,.45),0 4px 14px -10px rgba(230,0,46,.35) !important;}
.st-key-stepnav [data-testid="stColumn"]:not(:last-child)::after,.st-key-stepnav [data-testid="column"]:not(:last-child)::after{
  background:linear-gradient(90deg,#004c97,#e6002e);-webkit-background-clip:text;background-clip:text;color:transparent !important;font-weight:700;}
.st-key-stepnav .stButton > button{border-radius:10px !important;transition:all .15s ease !important;}
.st-key-stepnav .stButton > button[kind="primary"]{
  background:linear-gradient(120deg,#004c97 0%,#003a75 58%,#e6002e 135%) !important;color:#fff !important;border:none !important;
  box-shadow:0 6px 14px -6px rgba(0,76,151,.55),inset 0 -2px 0 rgba(230,0,46,.9) !important;}
.st-key-stepnav .stButton > button[kind="primary"] p{color:#fff !important;font-weight:700 !important;}
.st-key-stepnav .stButton > button[kind="secondary"]:hover:not(:disabled){background:#fdf2f4 !important;color:#e6002e !important;}
/* SOP 왼쪽 정렬 + 하위 번호 */
.sop-step{text-align:left !important;align-items:flex-start !important;}
.sop-t{justify-content:flex-start !important;}
.sop-body li{display:flex;gap:8px;align-items:flex-start;}
.sop-n{flex-shrink:0;min-width:28px;font-size:11.5px;font-weight:800;color:#004c97;background:#eef4fb;border-radius:5px;
  padding:1px 5px;margin-top:2px;text-align:center;}
.sop-body li.note{display:block;}
/* 표 색 구분: iCReaT eCRF = 빨강, CODA 코드북 = 파랑 */
.ecrf-tbl{border-color:#f6c9d2 !important;}
.ecrf-tbl table.pv th{background:#fdf2f4 !important;color:#8a0019 !important;border-bottom:2px solid #e6002e !important;}
.ecrf-tbl table.pv td.v{color:#e6002e !important;}
.coda-tbl{border-color:#c7d8ee !important;}
.coda-tbl table.pv th{background:#eef4fb !important;color:#003a75 !important;border-bottom:2px solid #004c97 !important;}
.coda-tbl table.pv td.v{color:#004c97 !important;}
</style>""")


def go(page):
    if page == "upload":
        ss.stage = "upload"
    else:
        ss.stage, ss.page = "main", page


@st.cache_data(show_spinner=False)
def _logo_b64(path):
    return base64.b64encode(Path(path).read_bytes()).decode()


def logo_html():
    if LOGO_FILE.exists():
        return f'<img class="brand-logo" src="data:image/png;base64,{_logo_b64(str(LOGO_FILE))}" alt="질병관리청 국립보건연구원">'
    return f'<div class="brand-ic">{icon("db", 28, "#ffffff", 2)}</div>'


def header():
    """헤더: 브랜드 + SOP 단계 버튼(① 업로드 ② 매핑 ③ 입력 폼 ④ 결과)"""
    cur = "upload" if ss.stage == "upload" else ss.page
    ready = bool(ss.get("upload"))
    with st.container(key="topbar"):
        h1, h2 = st.columns([1, 1.25], vertical_alignment="center")
        with h1:
            md(f'<div class="brand">{logo_html()}<div class="brand-t">'
               f'<div class="t1">iCReaT 코드북 데이터 자동 기탁 시스템</div>'
               f'<div class="t2">Clinical &amp; Bio-Data Curation &amp; Mapping Support System</div></div></div>')
        with h2:
            with st.container(key="stepnav"):
                cols = st.columns([1, 1.15, 1.5, 0.95], gap="small", vertical_alignment="center")
                cur_i = [k for k, _ in STEPS].index(cur)
                done = [k for k, _ in STEPS][:cur_i] if ready else []
                if done:  # 지나온 단계: 파란 글씨 + ✓
                    md("<style>" + ",".join(f'.st-key-stp_{k} .stButton > button' for k in done)
                       + "{color:#004c97 !important;font-weight:600 !important;}</style>")
                for i, ((key, label), col) in enumerate(zip(STEPS, cols)):
                    with col:
                        with st.container(key=f"stp_{key}"):
                            mark = "✓" if key in done else f"{i + 1}"
                            st.button(f"{mark}\u2002{label}", key=f"stepbtn_{key}", use_container_width=True,
                                      type="primary" if key == cur else "secondary",
                                      disabled=(key != "upload" and not ready), on_click=go, args=(key,))


# ══════════════════════════════════════════════════════════════
# 6. ① 첫 화면 : 파일 업로드 + SOP
# ══════════════════════════════════════════════════════════════
SOP = [
    ("1", "파일 준비 · 업로드", "upload", [
        "iCReaT <b>과제 정보</b> 화면에서 <kbd>Ctrl</kbd> + <kbd>P</kbd> → 대상 'PDF로 저장'",
        "iCReaT에서 <b>eCRF 코드북(xlsx)</b> 내려받기",
        "두 파일을 아래 <b>파일 업로드에 업로드</b>"]),
    ("2", "변수 매핑", "layers", [
        "eCRF 코드북의 변수명 → CODA 코드북의 변수명으로 자동 변환",
        "과제정보 PDF → 과제정보 항목 자동 추출",
        "<b> 매핑 결과·검토 필요 항목 확인 </b>"]),
    ("3", "과제 및 데이터 입력 폼", "edit", [
        "<b> 자동 입력된 과제정보 확인 · 수정 </b>",
        "임상 변수별 <b>공개 / 비공개 지정 </b>",
        "대상자 수 · 기탁 담당자 · 기탁일 등 입력"]),
    ("4", "결과 확인 · 다운로드", "download", [
        "② 매핑 결과 + ③ 입력 내용을 합친 <b>CODA 기탁용 임상정보 코드북(xlsx)</b> 다운로드",
        "<span style='font-size:10px'><b> ※ 동의서와 데이터는 기탁 사이트 양식으로 직접 제출 </b>"]),
]


def render_sop():
    cards = ""
    for n, t, ic, items in SOP:
        k, lis = 0, ""
        for x in items:
            if "※" in x:   # 안내 문구는 번호 없이
                lis += f'<li class="note">{x}</li>'
            else:
                k += 1
                lis += f'<li><span class="sop-n">{n}-{k}</span><span>{x}</span></li>'

        cards += (f'<div class="sop-step"><div class="sop-num">{n}</div><div class="sop-body">'
                  f'<div class="sop-t">{icon(ic, 18, "#004c97")} {t}</div><ul>{lis}</ul></div></div>')
    md(f'<div class="sop"><div class="sop-h">{icon("clip", 22, "#004c97")} CODA 기탁 SOP '
       f'<small>iCReaT 과제 → CODA 임상정보 기탁까지 4단계</small></div><div class="sop-grid">{cards}</div></div>')


def render_landing():
    render_sop()
    left, right = st.columns([2.05, 1], gap="large")
    with left:
        with st.container(key="upcard"):
            md(f'<div class="pg-head">{icon("upload", 26, "#4b5563")}<div><div class="h">파일 업로드</div>'
               f'<div class="s">과제정보 PDF와 eCRF 코드북(xlsx)을 함께 끌어다 놓으세요.</div></div></div>')
            up_key = f"up_files_{ss.get('up_ver', 0)}"
            prev = ss.get(up_key) or []
            both = any(f.name.lower().endswith(".pdf") for f in prev) and \
                any(f.name.lower().endswith((".xlsx", ".csv")) for f in prev)
            with st.container(key="upzone_done" if both else "upzone"):
                files = st.file_uploader("업로드 파일", type=["pdf", "xlsx", "csv"], accept_multiple_files=True,
                                         key=up_key, label_visibility="collapsed")
            if both:
                with st.container(key="reupbtn"):
                    if st.button(":material/restart_alt: 파일 다시 선택", use_container_width=True):
                        ss.up_ver = ss.get("up_ver", 0) + 1
                        st.rerun()
            pdf_f = next((f for f in reversed(files or []) if f.name.lower().endswith(".pdf")), None)
            ecrf_f = next((f for f in reversed(files or []) if f.name.lower().endswith((".xlsx", ".csv"))), None)
            pp = parse_pdf(pdf_f.getvalue(), pdf_f.name) if pdf_f else None
            ep = parse_ecrf_file(ecrf_f.getvalue(), ecrf_f.name) if ecrf_f else None
            if pp or ep:
                c1, c2 = st.columns(2, gap="medium")
                with c1:
                    render_pdf_result(pp)
                with c2:
                    render_ecrf_result(ep)
                ok = bool(ep and ep["ok"] and pp and not pp.get("error"))
                if ep and ep["ok"] and not pp:
                    md(f'<div class="up-warn">{icon("alert", 16, "#3a3f47")} 과제정보 PDF가 아직 없습니다. '
                       'iCReaT 과제 정보 화면에서 Ctrl+P → PDF로 저장 후 함께 올려 주세요.</div>')
                with st.container(key="startbtn"):
                    if st.button(":material/arrow_forward: 변수명 매핑 시작 (2단계로)", disabled=not ok,
                                 use_container_width=True):
                        start_mapping(ep, pp)
                        st.rerun()
    with right:
        render_side(pp if files else None, ep if files else None)


def render_pdf_result(pp):
    if not pp:
        md(f'<div class="fbox empty">{icon("pdf", 22, "#9ca3af")}<div><b>과제정보 PDF</b><span>업로드 대기</span></div></div>')
        return
    if pp.get("error"):
        md(f'<div class="fbox bad">{icon("pdf", 22, "#111418")}<div><b>{E(pp["name"])}</b><span>{E(pp["error"])}</span></div></div>')
        return
    vals = pp["values"]
    n = sum(1 for k in META_KEYS if vals.get(k))
    if pp["scanned"]:
        md(f'<div class="fbox warn">{icon("pdf", 22, "#3a3f47")}<div><b>{E(pp["name"])}</b>'
           '<span>텍스트가 없는 PDF(스캔본)입니다. 3단계 입력 폼에서 직접 입력해 주세요.</span></div>'
           '<span class="ps cln">직접 입력</span></div>')
        return
    md(f'<div class="fbox ok">{icon("pdf", 22, "#004c97")}<div><b>{E(pp["name"])}</b>'
       f'<span></span></div>'
       f'<span class="ps {"map" if n == len(META_KEYS) else "cln"}">{"추출 완료" if n == len(META_KEYS) else "일부 추출"}</span></div>')
    trs = "".join(
        f'<tr><th>{FIELD_LABELS[k]}</th><td>{E(str(vals[k]))[:80] if vals.get(k) else "<span class=miss>미추출 · 직접 입력</span>"}</td></tr>'
        for k in META_KEYS + ["n_subjects"])
    md(f'<div class="tbl-wrap" style="max-height:none"><table class="info mini"><tbody>{trs}</tbody></table></div>')


def render_ecrf_result(ep):
    if not ep:
        md(f'<div class="fbox empty">{icon("table", 22, "#9ca3af")}<div><b>iCReaT eCRF 코드북</b><span>업로드 대기</span></div></div>')
        return
    if not ep["rows"]:
        md(f'<div class="fbox bad">{icon("table", 22, "#111418")}<div><b>{E(ep["name"])}</b><span>{E(ep["error"])}</span></div></div>')
        return
    rows = ep["rows"]
    md(f'<div class="fbox {"ok" if ep["ok"] else "bad"}">{icon("table", 22, "#004c97" if ep["ok"] else "#111418")}<div>'
       f'<b>{E(ep["name"])}</span></div>'
       f'<span class="ps {"map" if ep["ok"] else "rev"}">{"인식 완료" if ep["ok"] else "필수 컬럼 누락"}</span></div>')
    trs = ""
    for key, _, target in REQ_COLS:
        hit = key in ep["hit"]
        trs += (f'<tr><td>{E(ep["hit"][key][1][:18]) if hit else "<span class=miss>없음</span>"}</td>'
                f'<td class="ar">{icon("arrow", 13, "#9ca3af", 2)}</td><td>{E(target)}</td>'
                f'<td><span class="ps {"map" if hit else "rev"}">{"인식" if hit else "누락"}</span></td></tr>')
    md(f'<div class="tbl-wrap" style="max-height:none"><table class="pv up-map"><thead><tr><th>파일 컬럼</th><th></th>'
       f'<th>CODA 양식 필드</th><th>상태</th></tr></thead><tbody>{trs}</tbody></table></div>')


def render_side(pp, ep):
    def st_(ok, label):
        return f'<span class="ps {"map" if ok else "none"}">{"첨부됨" if ok else label}</span>'
    pdf_ok = bool(pp and not pp.get("error"))
    ecrf_ok = bool(ep and ep["ok"])
    md(f"""<div class="card tpl"><div class="tpl-h">{icon("clip", 20, "#004c97")} 필수 첨부 파일</div>
      <div class="tpl-sub">다음 파일이 첨부되어야 합니다.</div>
      <div class="need"><span class="nn">1</span><div><b>과제정보 파일 (PDF)</b></div>{st_(pdf_ok, "필요")}</div>
      <div class="need"><span class="nn">2</span><div><b>iCReaT eCRF 파일 (xlsx)</b></div>{st_(ecrf_ok, "필요")}</div>
      <div class="tpl-h2">자동 매핑 규칙</div>
      <div class="tpl-rule pdfr"><b>과제정보 PDF</b>{icon("arrow", 12, "#9ca3af", 2)}과제정보 시트</div>
      <div class="tpl-rule"><b>eCRF명</b>{icon("arrow", 12, "#9ca3af", 2)}테이블명(영문/국문)</div>
      <div class="tpl-rule"><b>변수명</b>{icon("arrow", 12, "#9ca3af", 2)}변수명(영문)</div>
      <div class="tpl-rule"><b>항목명</b>{icon("arrow", 12, "#9ca3af", 2)}변수 설명</div>
      <div class="tpl-rule"><b>데이터 유형</b>{icon("arrow", 12, "#9ca3af", 2)}변수 타입</div>
      <div class="tpl-rule"><b>코드리스트</b>{icon("arrow", 12, "#9ca3af", 2)}변수값(코드) 설명</div>
      <div class="tpl-rule cur"><b>큐레이션</b> No. · 통합 설문지 생성 · 공개여부</div></div>""")


# ══════════════════════════════════════════════════════════════
# 7. 화면 진입
# ══════════════════════════════════════════════════════════════
header()
# 업로드 전이면 첫 화면만 그린다. st.stop() 에만 의존하지 않고 if 분기로 처리
# (VS Code 디버거 / python 직접 실행처럼 st.stop() 이 동작하지 않는 환경에서도 아래 코드로 넘어가지 않게)
LANDING = ss.stage == "upload" or not ss.get("upload")
if LANDING:
    with st.container(key="body"):
        render_landing()
    st.stop()
    pid = ECRF = form = META = CODA = None
    PCT, DONE, CNT = {"key": 0, "map": 0, "rev": 0}, 0, Counter()
else:
    pid = ss.upload["pid"]
    ECRF = ss.upload["rows"]
    if pid not in ss.forms:  # 세션이 꼬였을 때 안전장치
        ss.forms[pid] = new_form({"pid": pid}, ECRF)
    form = ss.forms[pid]
    META = form["meta"]
    CODA = build_coda(ECRF, form["vars"], form["table_ko"])
    PCT, DONE, CNT = dist(CODA)


# ══════════════════════════════════════════════════════════════
# 8. 좌측 패널
# ══════════════════════════════════════════════════════════════
def donut_svg(segs, center, sub="평균 완성도"):
    r, c, off, arcs = 70, 2 * 3.14159 * 70, 0, ""
    for pct, col in segs:
        ln = c * pct / 100
        arcs += (f'<circle cx="90" cy="90" r="{r}" fill="none" stroke="{col}" stroke-width="18" '
                 f'stroke-dasharray="{ln:.1f} {c - ln:.1f}" stroke-dashoffset="{-off:.1f}" transform="rotate(150 90 90)"/>')
        off += ln
    return (f'<svg width="180" height="180" viewBox="0 0 180 180"><circle cx="90" cy="90" r="{r}" fill="none" '
            f'stroke="#eef0f3" stroke-width="18"/>{arcs}<text x="90" y="92" text-anchor="middle" font-size="27" '
            f'font-weight="700" fill="#1a1d23">{center}</text><text x="90" y="116" text-anchor="middle" '
            f'font-size="12" fill="#9ca3af">{sub}</text></svg>')


def left_panel():
    with st.container(key="menucard"):
        md('<div class="menu-title">메뉴 관리</div>')
        for key, label in [("mapping", "변수 매핑"), ("form", "과제 및 데이터 입력 폼"), ("result", "결과 확인")]:
            with st.container(key=f"nav_{key}"):
                st.button(label, key=f"btn_{key}", type="primary" if ss.page == key else "secondary",
                          use_container_width=True, on_click=go, args=(key,))
    md(f'<div class="card curfile"><div class="cf-h">{icon("file", 17, "#4b5563")} 현재 과제 <code>{E(pid)}</code></div>'
       f'<div class="cf-row">{icon("pdf", 15, "#004c97")}<span>{E(ss.upload["pdf_name"] or "과제정보 PDF 없음")}</span></div>'
       f'<div class="cf-row">{icon("table", 15, "#004c97")}<span>{E(ss.upload["ecrf_name"])}</span></div></div>')
    items = [("매핑 완료", PCT["map"], "#004c97"), ("큐레이션", PCT["key"], "#e6002e"), ("누락", PCT["rev"], "#d1d5db")]
    legend = "".join(f'<div class="legend-row"><span><i style="background:{c}"></i>{n}</span><b>{v}%</b></div>'
                     for n, v, c in items)
    filled = info_filled()
    legend += (f'<div class="legend-row sub"><span>과제정보 입력</span><b>{filled}/{len(INFO_ROWS)}</b></div>'
               f'<div class="legend-row"><span>공개 변수</span><b>{var_counts()[1]}/{var_counts()[0]}</b></div>')
    md(f'<div class="card dist"><div class="dist-head"><b>{icon("pie", 19, "#4b5563")} 변환/정제 항목 분포</b>'
       f'<span class="badge-live">실시간</span></div><div class="donut">'
       f'{donut_svg([(PCT["map"], "#004c97"), (PCT["key"], "#e6002e"), (PCT["rev"], "#d1d5db")], f"{readiness()}%", "기탁 준비도")}</div>{legend}</div>')


# ══════════════════════════════════════════════════════════════
# 9. 공통 렌더러
# ══════════════════════════════════════════════════════════════
def var_counts():
    v = form["vars"]
    pub = sum(1 for x in v.values() if x)
    return len(v), pub, len(v) - pub


# 상태: 큐레이션만 완료(키 변수 등 eCRF에 없는 행) → 큐레이션 / 큐레이션 후 매핑까지 완료 → 매핑 완료 / 둘 다 아님 → 누락
STATUS_BADGE = {"map": ("map", "매핑 완료"), "cln": ("map", "매핑 완료"), "rev": ("rev", "누락"), "key": ("key", "큐레이션")}


def coda_table_html(rows):
    head = "".join(f"<th>{c}</th>" for c in CODA_COLS) + "<th>상태</th><th>원본 (eCRF)</th>"
    body = ""
    for r in rows:
        pub = r["공개여부"]
        pb = {"공개": "pub", "비공개": "pri", "": "none"}[pub]
        codes = E(r["변수값(코드) 설명"]).replace("\n", "<br>") or "-"
        var = E(r["변수명(영문)"]) or '<span class="pb none">누락</span>'
        cls, txt = STATUS_BADGE[r["_status"]]
        iss = "".join(f'<span class="issue">· {E(x)}</span>' for x in r["_issues"]) + \
            "".join(f'<span class="curn">· {E(x)}</span>' for x in r.get("_cur", []))
        survey = E(r["통합 설문지"]).replace("\n", "<br>")
        src = '<span class="pb cur">큐레이션</span>' if r["_status"] == "key" else E(r["_src"])
        body += (f'<tr class="{"key" if r["_status"] == "key" else "var"}"><td class="no">{r["No."]}</td>'
                 f'<td><span class="pb {pb}">{pub}</span></td><td class="mono">{E(r["테이블명(영문)"])}</td>'
                 f'<td>{E(r["테이블명(국문)"]) or "-"}</td><td class="v">{var}</td><td class="item">{E(r["변수 설명"])}</td>'
                 f'<td class="wrap">{codes}</td><td>{E(r["변수 타입"])}</td><td class="survey">{survey}</td>'
                 f'<td><span class="ps {cls}">{txt}</span>{iss}</td><td class="src">{src}</td></tr>')
    return f'<div class="tbl-wrap coda-tbl"><table class="pv"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def ecrf_table_html(rows):
    head = "".join(f"<th>{c}</th>" for c in ECRF_COLS)
    body = "".join(
        f'<tr><td class="mono">{E(r["eCRF"])}</td><td class="item">{E(r["항목명"])}</td><td class="v">{E(r["변수명"]) or "-"}</td>'
        f'<td>{E(r["항목 유형"])}</td><td class="mono">{E(r["데이터 유형"])}</td><td class="no">{E(r["길이"]) or "-"}</td>'
        f'<td>{E(r["단위"]) or "-"}</td><td class="wrap">{E(r["코드리스트"]).replace("|", "<br>") or "-"}</td></tr>'
        for r in rows)
    return f'<div class="tbl-wrap ecrf-tbl"><table class="pv"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def ecrf_names():
    return list(dict.fromkeys(r["eCRF"] for r in ECRF))


def filter_by_ecrf(sel):
    if sel.startswith("전체"):
        return CODA, ECRF
    return ([r for r in CODA if r["_status"] == "key" or r.get("_ecrf") == sel], [r for r in ECRF if r["eCRF"] == sel])


def next_btn(label, page, key):
    with st.container(key=f"next_{key}"):
        st.button(f":material/arrow_forward: {label}", key=f"nb_{key}", on_click=go, args=(page,), use_container_width=True)


# ══════════════════════════════════════════════════════════════
# 10. ② 변수 매핑
# ══════════════════════════════════════════════════════════════
def page_mapping():
    with st.container(key="preview"):
        with st.container(key="prevtop"):
            a, b = st.columns([3, 1.1], vertical_alignment="center")
            with a:
                md(f'<div class="pv-title">{icon("table", 22, "#004c97")} CODA 코드북 미리보기</div>'
                   f'</div>')
            with b:
                with st.container(key="pvfilter"):
                    sel = st.selectbox("eCRF", [f"전체 eCRF ({len(ecrf_names())}종)"] + ecrf_names(),
                                       label_visibility="collapsed", key=f"pv_sel_{pid}")
        coda_v, ecrf_v = filter_by_ecrf(sel)
        t1, t2 = st.tabs([f"iCReaT eCRF 원본 ({len(ecrf_v)})", f"CODA 코드북 미리보기 ({len(coda_v)})"])
        with t1:
            md(ecrf_table_html(ecrf_v))
        with t2:
            md(coda_table_html(coda_v))
    render_table_ko()
    next_btn("3단계 · 과제·데이터 입력으로", "form", "m")


def cb_table_ko(key, names):
    for i, ch in ss.get(key, {}).get("edited_rows", {}).items():
        n = names[int(i)]
        t = form["table_ko"][n]
        if "테이블명(영문)" in ch:
            t["en"] = (ch["테이블명(영문)"] or "").strip() or n
        if "테이블명(국문)" in ch:
            t["ko"] = (ch["테이블명(국문)"] or "").strip()
        if "테이블명(영문)" in ch or "테이블명(국문)" in ch:
            t["status"], t["note"] = "수동", "연구자 수정"
    bump()


def render_table_ko():
    names = list(form["table_ko"])
    rows = [{"테이블명(영문)": form["table_ko"][n].get("en") or n, "테이블명(국문)": form["table_ko"][n]["ko"],
             "번역": form["table_ko"][n]["status"]} for n in names]
    with st.container(key="tkocard"):
        n_rev = sum(1 for n in names if form["table_ko"][n]["status"] == "검토")
        md(f'<div class="pv-title">{icon("edit", 20, "#004c97")} 테이블명 자동 번역 (영문 → 국문)</div>'
           f'<div class="pv-sub">영문·국문 칸을 눌러 바로 수정할 수 있습니다. 수정한 행은 번역이 \'수동\'으로 바뀝니다.'
           f'{f" · 확인 필요 {n_rev}건" if n_rev else ""}</div>')
        key = f"w_tko_{pid}_{ss.ever.get(pid, 0)}"
        st.data_editor(pd.DataFrame(rows), key=key, hide_index=True, use_container_width=True, num_rows="fixed",
                       disabled=["번역"], on_change=cb_table_ko, args=(key, names),
                       height=38 + 35 * len(rows),
                       column_config={"테이블명(영문)": st.column_config.TextColumn("테이블명(영문) ✎", required=True),
                                      "테이블명(국문)": st.column_config.TextColumn("테이블명(국문) ✎", required=True),
                                      "번역": st.column_config.TextColumn("번역", width="small")})


# ══════════════════════════════════════════════════════════════
# 11. ③ 과제·데이터 입력
# ══════════════════════════════════════════════════════════════
def bump():
    ss.ever[pid] = ss.ever.get(pid, 0) + 1


def cb_edit(ed_key, uids):
    for i, ch in ss.get(ed_key, {}).get("edited_rows", {}).items():
        if "공개" in ch:
            form["vars"][uids[int(i)]] = bool(ch["공개"])
    bump()


def cb_bulk(mode):
    for uid in form["vars"]:
        form["vars"][uid] = mode == "pub"
    bump()


def meta_input(key, area=False):
    label = FIELD_LABELS[key] if key != "org" else "과제 수행기관(데이터 기탁기관)"
    found = ss.upload["pdf_found"].get(key)
    orig = form["pdf_meta"].get(key, "")
    cur = META.get(key, "")
    if not orig:
        tag = '<span class="mtag none">미추출 · 직접 입력</span>'
    elif cur != orig:
        tag = '<span class="mtag edit">수정됨</span>'
    else:
        tag = f'<span class="mtag pdf">PDF 자동 입력 · \'{E(found or "")}\'</span>'
    md(f'<div class="mlabel">{label}{tag}</div>')
    fn = st.text_area if area else st.text_input
    kw = dict(height=78) if area else {}
    META[key] = fn(label, value=cur, key=f"w_meta_{key}_{pid}", label_visibility="collapsed", **kw)


def page_form():
    with st.container(key="formcard"):
        a, b = st.columns([3, 1], vertical_alignment="center")
        with a:
            md(f'<div class="pg-head">{icon("file", 26, "#4b5563")}<div><div class="h">연구 과제 메타데이터 &amp; 데이터 입력 폼</div>'
               f'<div class="s">과제정보 PDF에서 자동 입력된 항목을 확인·수정하고, 연구자 입력 항목과 변수 공개 여부를 채웁니다.</div></div></div>')
        with b:
            with st.container(key="savebtn"):
                if st.button(":material/save: 설정 및 폼 양식 저장", key="save"):
                    ss.saved[pid] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
                    st.toast("입력 폼이 저장되었습니다.", icon="✅")
                if pid in ss.saved:
                    md(f'<div class="saved">마지막 저장: {ss.saved[pid]}</div>')
        md('<div class="hr"></div>')

        # ── 자동 입력 (PDF) ──
        src = ss.upload["pdf_name"] or "PDF 없음"
        note = "스캔본 PDF · 직접 입력 필요" if ss.upload["pdf_scanned"] else f"{src}에서 자동 추출 · 수정 가능"
        md(f'<div class="sect auto">{icon("pdf", 19, "#004c97")} [자동 입력] 과제정보 PDF 추출 데이터'
           f'<span class="note">{icon("edit", 14, "#6b7280")} {E(note)}</span></div>')
        c1, c2 = st.columns(2, gap="medium")
        with c1:
            meta_input("pid")
        with c2:
            meta_input("org")
        c1, c2 = st.columns(2, gap="medium")
        with c1:
            meta_input("title")
        with c2:
            meta_input("pi")
        c1, c2 = st.columns(2, gap="medium")
        with c1:
            meta_input("period")
        with c2:
            meta_input("disease")
        meta_input("goal", area=True)
        meta_input("content", area=True)
        md('<div class="hr" style="margin:14px 0 6px"></div>')

        # ── 연구자 입력 ──
        md(f'<div class="sect user">{icon("check", 19, "#004c97")} [연구자 입력 폼] 없는 데이터 채우기 · 기탁 정보 입력</div>')
        md(f'<div class="subh">{icon("send", 18, "#4b5563")} 데이터 기탁 정보</div>')
        c1, c2 = st.columns(2, gap="medium")
        with c1:
            form["n_subjects"] = st.number_input("대상자 수 (샘플 수)", min_value=0, step=1,
                                                 value=int(form["n_subjects"]), key=f"w_n_{pid}")
        with c2:
            form["date"] = st.date_input("데이터 기탁일", value=form["date"], key=f"w_date_{pid}", format="YYYY-MM-DD")
        c1, c2 = st.columns(2, gap="medium")
        with c1:
            form["mgr"] = st.text_input("데이터 기탁 담당자", value=form["mgr"], key=f"w_mgr_{pid}", placeholder="성명")
        with c2:
            form["mgr_tel"] = st.text_input("기탁 담당자 연락처", value=form["mgr_tel"], key=f"w_mtel_{pid}",
                                            placeholder="010-0000-0000 / 이메일")
        form["desc"] = st.text_area("데이터 설명 (특이사항)", value=form["desc"], key=f"w_desc_{pid}", height=78,
                                    placeholder="결측 코딩 방식, 추적 기간, 데이터 수집 시 특이사항 등")

        md(f'<div class="subh">{icon("user", 18, "#4b5563")} CODA 데이터 기탁 신청자 <small>기탁 담당자와 다른 경우 입력</small></div>')
        with st.container(key="coda"):
            form["same"] = st.checkbox("데이터 기탁 담당자와 동일", value=form["same"], key=f"w_same_{pid}")
            c1, c2, c3 = st.columns(3, gap="medium")
            if form["same"]:
                with c1:
                    st.text_input("신청자 성명", value=form["mgr"], disabled=True, key=f"w_cn_same_{pid}")
                with c3:
                    st.text_input("신청자 연락처", value=form["mgr_tel"], disabled=True, key=f"w_ct_same_{pid}")
            else:
                with c1:
                    form["coda_name"] = st.text_input("신청자 성명", value=form["coda_name"], key=f"w_cn_{pid}")
                with c3:
                    form["coda_tel"] = st.text_input("신청자 연락처", value=form["coda_tel"], key=f"w_ct_{pid}")
            with c2:
                form["coda_id"] = st.text_input("CODA 아이디", value=form["coda_id"], key=f"w_cid_{pid}")

        # ── 임상 변수 리스트 (eCRF 전체 변수 · 체크 = 공개) ──
        md(f'<div class="subh">{icon("list", 18, "#4b5563")} 임상 변수 리스트 및 공개 여부 '
           f'<small>eCRF의 모든 변수가 표시됩니다 · 체크 = 공개, 해제 = 비공개 → CODA 코드북 \'공개여부\'에 반영</small></div>')
        n_var, n_pub, n_pri = var_counts()
        with st.container(key="varbulk"):
            c1, c2, c3 = st.columns([4.2, 1, 1], vertical_alignment="center")
            with c1:
                md(f'<div class="cnt">임상 변수 개수 <span class="n">{n_var}개</span> = 공개 <b>{n_pub}</b> + 비공개 <b>{n_pri}</b>'
                   f'<span class="f">= SUM(eCRF 변수)</span></div>')  # 박스 스타일은 .st-key-varbulk
            with c2:
                with st.container(key="pubbtn"):
                    st.button(":material/check_box: 전체 공개", use_container_width=True, on_click=cb_bulk, args=("pub",))
            with c3:
                st.button(":material/check_box_outline_blank: 전체 비공개", use_container_width=True,
                          on_click=cb_bulk, args=("pri",))
        uids = list(form["vars"].keys())
        df = pd.DataFrame([{"공개": form["vars"][u], "No.": n, "eCRF": ECRF[u]["eCRF"], "변수명": ECRF[u]["변수명"] or "(없음)",
                            "항목명": ECRF[u]["항목명"], "공개여부": "공개" if form["vars"][u] else "비공개"}
                           for n, u in enumerate(uids, 1)],
                          columns=["공개", "No.", "eCRF", "변수명", "항목명", "공개여부"])
        ed_key = f"w_ed_{pid}_{ss.ever.get(pid, 0)}"
        with st.container(key="vared"):
            st.data_editor(df, key=ed_key, hide_index=True, use_container_width=True, num_rows="fixed",
                           height=min(38 + 35 * max(len(df), 1), 460), on_change=cb_edit, args=(ed_key, uids),
                           disabled=["No.", "eCRF", "변수명", "항목명", "공개여부"],
                           column_config={"공개": st.column_config.CheckboxColumn("공개 ✓", width="small"),
                                          "No.": st.column_config.NumberColumn("No.", width="small"),
                                          "공개여부": st.column_config.TextColumn("공개여부", width="small")})
    next_btn("4단계 · 결과 확인으로", "result", "f")


# ══════════════════════════════════════════════════════════════
# 12. ④ 결과 확인 = 과제정보 + 코드북 통합 → CODA 기탁 xlsx
# ══════════════════════════════════════════════════════════════
def info_values():
    n_var, n_pub, n_pri = var_counts()
    names = [ECRF[u]["변수명"] or f'({ECRF[u]["항목명"]})' for u in form["vars"]]
    coda_txt = ("기탁 담당자와 동일" + (f" / {form['coda_id']}" if form["coda_id"] else "")) if form["same"] else \
        " / ".join(x for x in (form["coda_name"], form["coda_id"], form["coda_tel"]) if x)
    text = {**{k: META.get(k, "") for k in META_KEYS},
            "var_list": ", ".join(names), "var_count": f"{n_var}개 (공개 {n_pub}, 비공개 {n_pri})" if n_var else "",
            "n_subjects": f"{form['n_subjects']:,}명" if form["n_subjects"] else "",
            "mgr": " / ".join(x for x in (form["mgr"], form["mgr_tel"]) if x),
            "date": form["date"].strftime("%Y-%m-%d") if form["date"] else "", "desc": form["desc"],
            "coda": coda_txt if (form["same"] and form["coda_id"]) or not form["same"] else ""}
    chips = "".join(f'<span class="{"" if form["vars"][u] else "pri"}">{E(ECRF[u]["변수명"] or "(없음)")}</span>'
                    for u in list(form["vars"])[:60])
    more = f'<span class="muted">… 외 {n_var - 60}개</span>' if n_var > 60 else ""
    html_ = dict(text)
    html_["var_list"] = (f'<div class="vchips">{chips}{more}</div><div class="muted" style="margin-top:4px">'
                         f'취소선 = 비공개 변수 · Excel에는 전체 목록 입력</div>') if n_var else ""
    return text, html_


def build_xlsx(text):
    """CODA 기탁_임상정보 코드북 양식에 과제정보 + 코드북을 채워 bytes 반환"""
    if TEMPLATE_FILE.exists():
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
    else:
        wb = openpyxl.Workbook()
        wb.active.title = "과제정보"
        wb.create_sheet("CODA 기탁 데이터 코드북 양식_임상정보")
    ws = wb["과제정보"]
    for row, label, key, _ in INFO_ROWS:
        ws.cell(row, 1).value = ws.cell(row, 1).value or label
        c = ws.cell(row, 2)
        c.value = text[key]
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if ws.cell(row, 11).value == 0:
            ws.cell(row, 11).value = None
    wc = wb["CODA 기탁 데이터 코드북 양식_임상정보"]
    for r in range(4, max(wc.max_row, 4) + 1):
        for col in range(1, 11):
            wc.cell(r, col).value = None
    for n, rr in enumerate(CODA):
        vals = [rr["No."], rr["공개여부"], rr["테이블명(영문)"], rr["테이블명(국문)"], rr["변수명(영문)"], rr["변수 설명"],
                rr["변수값(코드) 설명"], rr["변수 타입"], rr["통합 설문지"], ", ".join(rr["_issues"]) or None]
        for col, v in enumerate(vals, 1):
            c = wc.cell(4 + n, col)
            c.value = v
            c.alignment = Alignment(wrap_text=True, vertical="center")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def src_badge(key, src):
    if src == "pdf":
        orig = form["pdf_meta"].get(key, "")
        if not orig:
            return '<span class="sb user">직접 입력</span>'
        return '<span class="sb edit">PDF · 수정됨</span>' if META.get(key, "") != orig else '<span class="sb auto">PDF 자동</span>'
    return {"user": '<span class="sb user">연구자 입력</span>', "calc": '<span class="sb calc">자동 계산</span>'}[src]


def render_result_kpis(filled):
    n_var, n_pub, n_pri = var_counts()
    pct_pub = round(n_pub * 100 / max(n_var, 1))
    pct_info = round(filled * 100 / len(INFO_ROWS))
    cards = [
        ("hero", "기탁 준비도", f"{readiness()}", "%", "코드북(큐레이션·매핑) 완료율 + 과제정보 입력률", readiness(), "check"),
        ("hero", "과제정보 입력", f"{filled}", f"/ {len(INFO_ROWS)}", "CODA '과제정보' 시트", pct_info, "clip"),
        ("hero", "임상 변수", f"{n_var}", "개", f"공개 {n_pub} · 비공개 {n_pri}", pct_pub, "list"),
        ("hero", "누락", f"{CNT['rev']}", "건", "변수명 누락·중복 · 영문/국문 불일치", PCT["rev"], "alert"),
    ]
    html_ = ""
    for cls, label, val, unit, sub, bar, ic in cards:
        barh = f'<div class="k2-bar"><i style="width:{bar}%"></i></div>' if bar is not None else '<div class="k2-bar k2-nobar"></div>'
        color = "#ffffff" if cls == "hero" else "#004c97"
        html_ += (f'<div class="k2 {cls}"><div class="k2-top"><span>{label}</span><span class="k2-ic">{icon(ic, 18, color, 2)}</span></div>'
                  f'<div class="k2-val">{val}<small>{unit}</small></div>{barh}<div class="k2-sub">{sub}</div></div>')
    md(f'<div class="k2-grid">{html_}</div>')


def info_filled():
    text, _ = info_values()
    return sum(1 for _, _, k, _ in INFO_ROWS if text[k] or k == "desc")


def readiness():
    """기탁 준비도 = (코드북 변수 매핑·정제 완료율 + 과제정보 입력률) / 2"""
    return round((DONE + info_filled() * 100 / len(INFO_ROWS)) / 2)


def page_result():
    text, html_ = info_values()
    filled = sum(1 for _, _, k, _ in INFO_ROWS if text[k] or k == "desc")
    render_result_kpis(filled)
    with st.container(key="resultcard"):
        with st.container(key="restop"):
            a, b = st.columns([3, 1.15], vertical_alignment="center")
            with a:
                md(f'<div class="pg-head">{icon("layers", 28, "#004c97")}<div><div class="h">결과 확인 — CODA 기탁용 임상정보 코드북</div>'
                   f'<div class="s"> 내려받아 CODA에 제출하세요.</div></div></div>')
            with b:
                with st.container(key="xlsxbtn"):
                    st.download_button(":material/download: CODA 기탁 코드북 (xlsx)", build_xlsx(text),
                                       file_name=f"{pid}_기탁_임상정보_코드북_{dt.date.today():%Y%m%d}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
        md('<div class="hr"></div>')
        pct = round(filled * 100 / len(INFO_ROWS))
        md(f'<div class="res-sect"><span style="display:flex;align-items:center;gap:8px">{icon("clip", 20, "#4b5563")} ① CODA 데이터 기탁 과제 정보 '
           f'<small>엑셀 \'과제정보\' 시트</small></span><span class="prog">입력 {filled}/{len(INFO_ROWS)}'
           f'<span class="bar"><i style="width:{pct}%"></i></span></span></div>')
        trs = ""
        for _, label, key, src in INFO_ROWS:
            v = html_[key] if key == "var_list" else E(str(html_[key]))
            if not text[key]:
                v = '<span class="miss">미입력</span>' if key != "desc" else '<span class="muted">(선택 항목)</span>'
            trs += f'<tr><th>{E(label)}</th><td class="val">{v}</td><td class="src">{src_badge(key, src)}</td></tr>'
        md(f'<div class="tbl-wrap" style="max-height:none"><table class="info"><tbody>{trs}</tbody></table></div>')

        md(f'<div class="res-sect" style="margin-top:16px"><span style="display:flex;align-items:center;gap:8px">{icon("table", 20, "#4b5563")} '
           f'② CODA 기탁 데이터 코드북 (임상정보) <small>엑셀 코드북 시트 · 검토 사항은 비고 열 기재</small></span>'
           f'<span class="prog">큐레이션 {CNT["key"]} · 매핑 완료 {CNT["map"]} · 누락 {CNT["rev"]}</span></div>')
        c1, c2, _ = st.columns([1.2, 1.3, 2.5], vertical_alignment="center")
        with c1:
            sel = st.selectbox("eCRF", [f"전체 eCRF ({len(ecrf_names())}종)"] + ecrf_names(),
                               label_visibility="collapsed", key=f"res_sel_{pid}")
        with c2:
            only = st.radio("보기", ["전체", "누락만"], horizontal=True, label_visibility="collapsed", key=f"res_only_{pid}")
        rows, _ = filter_by_ecrf(sel)
        if only == "누락만":
            rows = [r for r in rows if r["_status"] == "rev"]
        md(coda_table_html(rows))


# ══════════════════════════════════════════════════════════════
# 13. 본문
# ══════════════════════════════════════════════════════════════
if not LANDING:
    with st.container(key="body"):
        left, main = st.columns([1, 3.1])
        with main:   # 입력 폼 값이 같은 실행에서 좌측 도넛에 반영되도록 본문을 먼저 그림
            {"mapping": page_mapping, "form": page_form, "result": page_result}.get(ss.page, page_mapping)()
        with left:
            left_panel()
