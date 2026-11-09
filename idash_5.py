"""iCReaT 연구데이터 정제 및 기탁 대시보드 (Streamlit)

SOP 흐름 (헤더 단계 버튼과 동일)
  ① 파일 업로드          : iCReaT 과제정보 PDF(Ctrl+P 저장) + eCRF 코드북(xlsx) drag & drop
  ② CODA 양식 매핑       : eCRF → CODA 코드북 자동 변환 미리보기, PDF → 과제정보 자동 추출
  ③ 과제 및 데이터 입력 폼 : PDF 자동 입력값 확인·수정, 임상 변수 공개/비공개, 기탁 정보 입력
  ④ 결과 확인            : ②+③을 합친 CODA 기탁용 임상정보 코드북(xlsx) 다운로드

파일
  pdf_extract.py                              : 과제정보 PDF 텍스트 추출 (OCR 없이 텍스트 레이어 사용)
  style.css                                   : 화면 디자인
  data/CODA_clinical_codebook_template.xlsx   : CODA 기탁_임상정보 코드북 양식 (결과 Excel의 틀)
  data/C170005_ecrf_codebook.xlsx             : 샘플 eCRF 코드북
  data/C170005_project_info_sample.pdf        : 샘플 과제정보 PDF
"""
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

st.set_page_config(page_title="iCReaT 연구데이터 정제 및 기탁 대시보드", page_icon="🗄️", layout="wide",
                   initial_sidebar_state="collapsed")
E = html.escape
BASE = Path(__file__).parent
DATA = BASE / "data"
TEMPLATE_FILE = DATA / "CODA_clinical_codebook_template.xlsx"
SAMPLE_ECRF = DATA / "C170005_ecrf_codebook.xlsx"
SAMPLE_PDF = DATA / "C170005_project_info_sample.pdf"

# ══════════════════════════════════════════════════════════════
# 1. 규칙 / 설정  (CODA 규칙 확정 시 여기만 수정)
# ══════════════════════════════════════════════════════════════
ECRF_KO = {  # eCRF명 → CODA 테이블명(국문). 영문 테이블명은 eCRF명 그대로
    "BaseLineInfo": "기본정보(이식)", "Baseline Information": "기본정보", "Baseline Information_": "기본정보(추가)",
    "BaseLine": "기본정보(BaseLine)", "Laboratory Values": "검사 결과", "Demographics": "인구학적 정보",
    "Drinking and Past-Medical History": "음주 및 과거 병력",
    "dringking and past-medical history_COPY": "음주 및 과거 병력(사본)",
    "Study Drug": "연구 약물", "AE": "이상반응", "그리드 반복": "그리드 반복",
}


def coda_type(dtype, codelist):
    """iCReaT 데이터 유형 → CODA 변수 타입"""
    if codelist:
        return "범주형"
    return {"Integer": "연속형", "Float": "연속형", "String": "문자형", "Date": "날짜형"}.get(dtype, "검토 필요")


def coda_codes(codelist):
    """'1=Normal|2=Abnormal' → '1=Normal,\\n2=Abnormal' (CODA 양식 표기)"""
    if not codelist:
        return ""
    return ",\n".join(p.strip() for p in codelist.replace(";", "|").split("|") if p.strip())


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
STEPS = [("upload", "파일 업로드"), ("mapping", "CODA 양식 매핑"), ("form", "과제 및 데이터 입력 폼"), ("result", "결과 확인")]

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
for k, v in dict(stage="upload", page="mapping", forms={}, saved={}, ever={}, dels={}, use_sample=False).items():
    ss.setdefault(k, v)


def new_form(meta, ecrf, sample=False):
    f = dict(meta={k: str(meta.get(k, "") or "") for k in META_KEYS}, pdf_meta={k: str(meta.get(k, "") or "") for k in META_KEYS},
             vars={}, n_subjects=int(meta["n_subjects"]) if str(meta.get("n_subjects", "")).isdigit() else 0,
             mgr="", mgr_tel="", date=dt.date.today() + dt.timedelta(days=30), desc="", same=True,
             coda_name="", coda_id="", coda_tel="")
    if sample:  # 샘플 체험용 예시값 (AE·그리드 반복은 미입력 → '검토 필요' 예시)
        f["vars"] = {i: not (r["변수명"].startswith(("BIRTHD", "CRCRA")) or r["항목명"] == "생년월일")
                     for i, r in enumerate(ecrf) if r["eCRF"] not in ("AE", "그리드 반복")}
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
def build_coda(ecrf, public_map):
    dup = {k for k, v in Counter(r["변수명"] for r in ecrf if r["변수명"]).items() if v > 1}
    rows = [{"No.": 0, "공개여부": "공개", "테이블명(영문)": KEY_TABLE[0], "테이블명(국문)": KEY_TABLE[1],
             "변수명(영문)": k, "변수 설명": d, "변수값(코드) 설명": "", "변수 타입": t,
             "_src": "큐레이션", "_status": "key", "_issues": []} for k, d, t in KEY_VARS]
    for no, (i, r) in enumerate(enumerate(ecrf), 1):
        issues = []
        if not r["변수명"]:
            issues.append("변수명 누락")
        elif r["변수명"] in dup:
            issues.append("변수명 중복")
        pub = public_map.get(i)
        if pub is None:
            issues.append("공개여부 미입력")
        typ = coda_type(r["데이터 유형"], r["코드리스트"])
        if typ == "검토 필요":
            issues.append("타입 미정")
        codes = coda_codes(r["코드리스트"])
        rows.append({"No.": no, "공개여부": "" if pub is None else ("공개" if pub else "비공개"),
                     "테이블명(영문)": r["eCRF"], "테이블명(국문)": ECRF_KO.get(r["eCRF"], ""),
                     "변수명(영문)": r["변수명"], "변수 설명": r["항목명"], "변수값(코드) 설명": codes, "변수 타입": typ,
                     "_src": f'{r["eCRF"]}.{r["변수명"] or "?"}',
                     "_status": "rev" if issues else ("cln" if codes else "map"), "_issues": issues})
    return rows


def dist(rows):
    v = [r for r in rows if r["_status"] != "key"]
    n = max(len(v), 1)
    c = Counter(r["_status"] for r in v)
    return {k: round(c[k] * 100 / n) for k in ("map", "cln", "rev")}, round((c["map"] + c["cln"]) * 100 / n), c


CODA_COLS = ["No.", "공개여부", "테이블명(영문)", "테이블명(국문)", "변수명(영문)", "변수 설명", "변수값(코드) 설명", "변수 타입"]
CODA_SRC = {"No.": "큐레이션", "공개여부": "← 연구자 입력", "테이블명(영문)": "← eCRF명", "테이블명(국문)": "← eCRF명",
            "변수명(영문)": "← 변수명", "변수 설명": "← 항목명", "변수값(코드) 설명": "← 코드리스트", "변수 타입": "← 데이터 유형"}


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


def go(page):
    if page == "upload":
        ss.stage = "upload"
    else:
        ss.stage, ss.page = "main", page


def header():
    """헤더: 브랜드 + SOP 단계 버튼(① 업로드 ② 매핑 ③ 입력 폼 ④ 결과)"""
    cur = "upload" if ss.stage == "upload" else ss.page
    ready = bool(ss.get("upload"))
    with st.container(key="topbar"):
        h1, h2 = st.columns([1, 1.25], vertical_alignment="center")
        with h1:
            md(f'<div class="brand"><div class="brand-ic">{icon("db", 28, "#1c2a24", 2)}</div><div class="brand-t">'
               f'<div class="t1">iCReaT 연구데이터 정제 및 기탁 대시보드</div>'
               f'<div class="t2">Clinical &amp; Bio-Data Curation &amp; Mapping Support System</div></div></div>')
        with h2:
            with st.container(key="stepnav"):
                cols = st.columns([1, 1.15, 1.5, 0.95], gap="small", vertical_alignment="center")
                for i, ((key, label), col) in enumerate(zip(STEPS, cols)):
                    with col:
                        with st.container(key=f"stp_{key}"):
                            st.button(f"{i + 1} {label}", key=f"stepbtn_{key}", use_container_width=True,
                                      type="primary" if key == cur else "secondary",
                                      disabled=(key != "upload" and not ready), on_click=go, args=(key,))


# ══════════════════════════════════════════════════════════════
# 6. ① 첫 화면 : 파일 업로드 + SOP
# ══════════════════════════════════════════════════════════════
SOP = [
    ("1", "파일 준비 · 업로드", "upload", [
        "iCReaT <b>과제 정보</b> 화면에서 <kbd>Ctrl</kbd> + <kbd>P</kbd> → 대상 <b>'PDF로 저장'</b> → 과제정보 PDF 저장",
        "iCReaT에서 <b>eCRF 코드북(xlsx)</b> 내려받기",
        "두 파일을 아래 영역에 끌어다 놓거나 <b>파일 선택</b>으로 추가"]),
    ("2", "CODA 양식 매핑", "layers", [
        "eCRF 코드북 → <b>CODA 임상정보 코드북</b> 자동 변환",
        "과제정보 PDF → <b>과제정보 항목</b> 자동 추출",
        "매핑 결과·검토 필요 항목 미리보기로 확인"]),
    ("3", "과제 및 데이터 입력 폼", "edit", [
        "PDF에서 자동 입력된 과제정보 확인 · 수정",
        "임상 변수별 <b>공개 / 비공개</b> 지정",
        "대상자 수 · 기탁 담당자 · 기탁일 등 입력"]),
    ("4", "결과 확인 · 다운로드", "download", [
        "②매핑 결과 + ③입력 내용을 합친 <b>CODA 기탁용 임상정보 코드북(xlsx)</b> 다운로드",
        "CODA 기탁 사이트에 제출",
        "동의서는 기탁 사이트 양식으로 <b>직접 제출</b>"]),
]


def render_sop():
    cards = ""
    for n, t, ic, items in SOP:
        lis = "".join(f"<li>{x}</li>" for x in items)
        cards += (f'<div class="sop-step"><div class="sop-num">{n}</div><div class="sop-body">'
                  f'<div class="sop-t">{icon(ic, 18, "#4e8f78")} {t}</div><ul>{lis}</ul></div></div>')
    md(f'<div class="sop"><div class="sop-h">{icon("clip", 22, "#4f6bdc")} CODA 기탁 SOP '
       f'<small>iCReaT 과제 → CODA 임상정보 기탁까지 4단계</small></div><div class="sop-grid">{cards}</div></div>')


def render_landing():
    render_sop()
    left, right = st.columns([2.05, 1], gap="large")
    with left:
        with st.container(key="upcard"):
            md(f'<div class="pg-head">{icon("upload", 26, "#4b5563")}<div><div class="h">파일 업로드</div>'
               f'<div class="s">과제정보 PDF와 eCRF 코드북(xlsx)을 함께 끌어다 놓으세요. 두 파일을 한 번에 올려도 됩니다.</div></div></div>')
            files = st.file_uploader("업로드 파일", type=["pdf", "xlsx", "csv"], accept_multiple_files=True,
                                     key="up_files", label_visibility="collapsed")
            pdf_f = next((f for f in reversed(files or []) if f.name.lower().endswith(".pdf")), None)
            ecrf_f = next((f for f in reversed(files or []) if f.name.lower().endswith((".xlsx", ".csv"))), None)
            pp = parse_pdf(pdf_f.getvalue(), pdf_f.name) if pdf_f else None
            ep = parse_ecrf_file(ecrf_f.getvalue(), ecrf_f.name) if ecrf_f else None
            if ss.use_sample and not files:
                pp = parse_pdf(SAMPLE_PDF.read_bytes(), "C170005_과제정보.pdf") if SAMPLE_PDF.exists() else None
                ep = parse_ecrf_file(SAMPLE_ECRF.read_bytes(), "C170005_전체_코드북_20261023_수정.xlsx")
            if pp or ep:
                c1, c2 = st.columns(2, gap="medium")
                with c1:
                    render_pdf_result(pp)
                with c2:
                    render_ecrf_result(ep)
                ok = bool(ep and ep["ok"] and pp and not pp.get("error"))
                if ep and ep["ok"] and not pp:
                    md(f'<div class="up-warn">{icon("alert", 16, "#a77a0c")} 과제정보 PDF가 아직 없습니다. '
                       'iCReaT 과제 정보 화면에서 Ctrl+P → PDF로 저장 후 함께 올려 주세요.</div>')
                with st.container(key="startbtn"):
                    if st.button(":material/arrow_forward: CODA 양식 매핑 시작 (2단계로)", disabled=not ok,
                                 use_container_width=True):
                        start_mapping(ep, pp, sample=ss.use_sample and not files)
                        st.rerun()
    with right:
        render_side(pp if (files or ss.use_sample) else None, ep if (files or ss.use_sample) else None)


def render_pdf_result(pp):
    if not pp:
        md(f'<div class="fbox empty">{icon("pdf", 22, "#9aa1ad")}<div><b>과제정보 PDF</b><span>업로드 대기</span></div></div>')
        return
    if pp.get("error"):
        md(f'<div class="fbox bad">{icon("pdf", 22, "#c0392b")}<div><b>{E(pp["name"])}</b><span>{E(pp["error"])}</span></div></div>')
        return
    vals = pp["values"]
    n = sum(1 for k in META_KEYS if vals.get(k))
    if pp["scanned"]:
        md(f'<div class="fbox warn">{icon("pdf", 22, "#a77a0c")}<div><b>{E(pp["name"])}</b>'
           '<span>텍스트가 없는 PDF(스캔본)입니다. 3단계 입력 폼에서 직접 입력해 주세요.</span></div>'
           '<span class="ps cln">직접 입력</span></div>')
        return
    md(f'<div class="fbox ok">{icon("pdf", 22, "#4e8f78")}<div><b>{E(pp["name"])}</b>'
       f'<span>과제정보 {n}/{len(META_KEYS)}개 항목 자동 추출 (OCR 없이 텍스트 추출)</span></div>'
       f'<span class="ps {"map" if n == len(META_KEYS) else "cln"}">{"추출 완료" if n == len(META_KEYS) else "일부 추출"}</span></div>')
    trs = "".join(
        f'<tr><th>{FIELD_LABELS[k]}</th><td>{E(str(vals[k]))[:80] if vals.get(k) else "<span class=miss>미추출 · 직접 입력</span>"}</td></tr>'
        for k in META_KEYS + ["n_subjects"])
    md(f'<div class="tbl-wrap" style="max-height:none"><table class="info mini"><tbody>{trs}</tbody></table></div>')


def render_ecrf_result(ep):
    if not ep:
        md(f'<div class="fbox empty">{icon("table", 22, "#9aa1ad")}<div><b>iCReaT eCRF 코드북</b><span>업로드 대기</span></div></div>')
        return
    if not ep["rows"]:
        md(f'<div class="fbox bad">{icon("table", 22, "#c0392b")}<div><b>{E(ep["name"])}</b><span>{E(ep["error"])}</span></div></div>')
        return
    rows = ep["rows"]
    md(f'<div class="fbox {"ok" if ep["ok"] else "bad"}">{icon("table", 22, "#4e8f78" if ep["ok"] else "#c0392b")}<div>'
       f'<b>{E(ep["name"])}</b><span>시트 \'{E(ep["sheet"])}\' · 컬럼 {len(ep["hit"])}/{len(REQ_COLS)}개 인식</span></div>'
       f'<span class="ps {"map" if ep["ok"] else "rev"}">{"인식 완료" if ep["ok"] else "필수 컬럼 누락"}</span></div>')
    n_ecrf = len(dict.fromkeys(r["eCRF"] for r in rows))
    md('<div class="up-stats">' + "".join(
        f'<div><span>{t}</span><b>{v}</b></div>' for t, v in
        [("eCRF", f"{n_ecrf}종"), ("변수", f"{len(rows)}개"), ("코드리스트", f"{sum(1 for r in rows if r['코드리스트'])}개"),
         ("변수명 누락", f"{sum(1 for r in rows if not r['변수명'])}건")]) + "</div>")
    trs = ""
    for key, _, target in REQ_COLS:
        hit = key in ep["hit"]
        trs += (f'<tr><td>{E(ep["hit"][key][1][:18]) if hit else "<span class=miss>없음</span>"}</td>'
                f'<td class="ar">{icon("arrow", 13, "#9aa1ad", 2)}</td><td>{E(target)}</td>'
                f'<td><span class="ps {"map" if hit else "rev"}">{"인식" if hit else "누락"}</span></td></tr>')
    md(f'<div class="tbl-wrap" style="max-height:none"><table class="pv up-map"><thead><tr><th>파일 컬럼</th><th></th>'
       f'<th>CODA 양식 필드</th><th>상태</th></tr></thead><tbody>{trs}</tbody></table></div>')


def render_side(pp, ep):
    def st_(ok, label):
        return f'<span class="ps {"map" if ok else "none"}">{"첨부됨" if ok else label}</span>'
    pdf_ok = bool(pp and not pp.get("error"))
    ecrf_ok = bool(ep and ep["ok"])
    md(f"""<div class="card tpl"><div class="tpl-h">{icon("clip", 20, "#4e8f78")} 필수 첨부 파일</div>
      <div class="tpl-sub">다음 파일이 첨부되어야 합니다.</div>
      <div class="need"><span class="nn">1</span><div><b>과제정보 파일 (PDF)</b><small>iCReaT 과제 정보 화면 → Ctrl+P → PDF로 저장</small></div>{st_(pdf_ok, "필요")}</div>
      <div class="need"><span class="nn">2</span><div><b>iCReaT eCRF 파일 (xlsx)</b><small>eCRF명·항목명·변수명·항목유형·데이터유형·길이·단위·코드리스트</small></div>{st_(ecrf_ok, "필요")}</div>
      <div class="tpl-h2">자동 매핑 규칙</div>
      <div class="tpl-rule pdfr"><b>과제정보 PDF</b>{icon("arrow", 12, "#9aa1ad", 2)}과제정보 시트(과제명·번호·기관·책임자 등)</div>
      <div class="tpl-rule"><b>eCRF명</b>{icon("arrow", 12, "#9aa1ad", 2)}테이블명(영문·국문)</div>
      <div class="tpl-rule"><b>변수명</b>{icon("arrow", 12, "#9aa1ad", 2)}변수명(영문)</div>
      <div class="tpl-rule"><b>항목명</b>{icon("arrow", 12, "#9aa1ad", 2)}변수 설명</div>
      <div class="tpl-rule"><b>데이터 유형</b>{icon("arrow", 12, "#9aa1ad", 2)}변수 타입</div>
      <div class="tpl-rule"><b>코드리스트</b>{icon("arrow", 12, "#9aa1ad", 2)}변수값(코드) 설명</div>
      <div class="tpl-rule cur"><b>큐레이션</b> No. · 공개여부(3단계 연구자 입력)</div></div>""")
    with st.container(key="samplebtn"):
        if st.button(":material/science: 샘플 파일로 체험 (C170005 PDF + eCRF)", use_container_width=True,
                     disabled=not (SAMPLE_ECRF.exists() and SAMPLE_PDF.exists())):
            ss.use_sample = True
            st.rerun()


# ══════════════════════════════════════════════════════════════
# 7. 화면 진입
# ══════════════════════════════════════════════════════════════
header()
if ss.stage == "upload" or not ss.get("upload"):
    with st.container(key="body"):
        render_landing()
    st.stop()

pid = ss.upload["pid"]
ECRF = ss.upload["rows"]
form = ss.forms[pid]
META = form["meta"]
CODA = build_coda(ECRF, form["vars"])
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
            f'font-weight="700" fill="#1e2433">{center}</text><text x="90" y="116" text-anchor="middle" '
            f'font-size="12" fill="#9aa1ad">{sub}</text></svg>')


def left_panel():
    with st.container(key="menucard"):
        md('<div class="menu-title">메뉴 관리</div>')
        for key, label in [("mapping", "CODA 양식 매핑"), ("form", "과제 및 데이터 입력 폼"), ("result", "결과 확인")]:
            with st.container(key=f"nav_{key}"):
                st.button(label, key=f"btn_{key}", type="primary" if ss.page == key else "secondary",
                          use_container_width=True, on_click=go, args=(key,))
    md(f'<div class="card curfile"><div class="cf-h">{icon("file", 17, "#4b5563")} 현재 과제 <code>{E(pid)}</code></div>'
       f'<div class="cf-row">{icon("pdf", 15, "#4e8f78")}<span>{E(ss.upload["pdf_name"] or "과제정보 PDF 없음")}</span></div>'
       f'<div class="cf-row">{icon("table", 15, "#4e8f78")}<span>{E(ss.upload["ecrf_name"])}</span></div></div>')
    items = [("매핑 완료", PCT["map"], "#7fb4a1"), ("정제 완료", PCT["cln"], "#e5c25a"), ("검토 필요 항목", PCT["rev"], "#d6dae1")]
    legend = "".join(f'<div class="legend-row"><span><i style="background:{c}"></i>{n}</span><b>{v}%</b></div>'
                     for n, v, c in items)
    md(f'<div class="card dist"><div class="dist-head"><b>{icon("pie", 19, "#4b5563")} 변환/정제 항목 분포</b>'
       f'<span class="badge-live">실시간</span></div><div class="donut">'
       f'{donut_svg([(PCT["map"], "#7fb4a1"), (PCT["cln"], "#e5c25a"), (PCT["rev"], "#eef0f3")], f"{DONE}%")}</div>{legend}</div>')


# ══════════════════════════════════════════════════════════════
# 9. 공통 렌더러
# ══════════════════════════════════════════════════════════════
def var_counts():
    v = form["vars"]
    pub = sum(1 for x in v.values() if x)
    return len(v), pub, len(v) - pub


STATUS_BADGE = {"map": ("map", "매핑 완료"), "cln": ("cln", "정제 완료"), "rev": ("rev", "검토 필요"), "key": ("key", "큐레이션")}


def coda_table_html(rows):
    head = "".join(
        f'<th>{c}<small class="{"c" if CODA_SRC[c] in ("큐레이션", "← 연구자 입력") else ""}">{CODA_SRC[c]}</small></th>'
        for c in CODA_COLS) + '<th>정제 상태<small>자동 점검</small></th><th>원본<small>iCReaT eCRF</small></th>'
    body = ""
    for r in rows:
        pub = r["공개여부"]
        pb = {"공개": "pub", "비공개": "pri", "": "none"}[pub]
        codes = E(r["변수값(코드) 설명"]).replace("\n", "<br>") or "-"
        var = E(r["변수명(영문)"]) or '<span class="pb none">누락</span>'
        cls, txt = STATUS_BADGE[r["_status"]]
        iss = "".join(f'<span class="issue">· {E(x)}</span>' for x in r["_issues"])
        src = '<span class="pb cur">큐레이션</span>' if r["_status"] == "key" else E(r["_src"])
        body += (f'<tr class="{"key" if r["_status"] == "key" else "var"}"><td class="no">{r["No."]}</td>'
                 f'<td><span class="pb {pb}">{pub or "미입력"}</span></td><td class="mono">{E(r["테이블명(영문)"])}</td>'
                 f'<td>{E(r["테이블명(국문)"]) or "-"}</td><td class="v">{var}</td><td class="item">{E(r["변수 설명"])}</td>'
                 f'<td class="wrap">{codes}</td><td>{E(r["변수 타입"])}</td>'
                 f'<td><span class="ps {cls}">{txt}</span>{iss}</td><td class="src">{src}</td></tr>')
    return f'<div class="tbl-wrap"><table class="pv"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def ecrf_table_html(rows):
    head = "".join(f"<th>{c}</th>" for c in ECRF_COLS)
    body = "".join(
        f'<tr><td class="mono">{E(r["eCRF"])}</td><td class="item">{E(r["항목명"])}</td><td class="v">{E(r["변수명"]) or "-"}</td>'
        f'<td>{E(r["항목 유형"])}</td><td class="mono">{E(r["데이터 유형"])}</td><td class="no">{E(r["길이"]) or "-"}</td>'
        f'<td>{E(r["단위"]) or "-"}</td><td class="wrap">{E(r["코드리스트"]).replace("|", "<br>") or "-"}</td></tr>'
        for r in rows)
    return f'<div class="tbl-wrap"><table class="pv"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def ecrf_names():
    return list(dict.fromkeys(r["eCRF"] for r in ECRF))


def filter_by_ecrf(sel):
    if sel.startswith("전체"):
        return CODA, ECRF
    return ([r for r in CODA if r["_status"] == "key" or r["테이블명(영문)"] == sel], [r for r in ECRF if r["eCRF"] == sel])


def next_btn(label, page, key):
    with st.container(key=f"next_{key}"):
        st.button(f":material/arrow_forward: {label}", key=f"nb_{key}", on_click=go, args=(page,), use_container_width=True)


# ══════════════════════════════════════════════════════════════
# 10. ② CODA 양식 매핑
# ══════════════════════════════════════════════════════════════
def page_mapping():
    n_var, n_pub, n_pri = var_counts()
    n_meta = sum(1 for k in META_KEYS if META.get(k))
    kpis = [
        ("과제정보 자동 입력", "pdf", "#4f6bdc", f"{n_meta} / {len(META_KEYS)}",
         f'<span style="color:#4f6bdc">과제정보 PDF에서 추출 · 3단계에서 수정 가능</span>'),
        ("eCRF 변수 (매핑 대상)", "layers", "#7fb4a1", f"{len(ECRF)}개",
         f'{icon("check", 17, "#4e8f78")}<span style="color:#4e8f78">eCRF {len(ecrf_names())}종 · CODA 변환 완료</span>'),
        ("검토 필요 항목", "alert", "#c98a1e", f"{CNT['rev']}건",
         '<span style="color:#c98a1e">변수명 누락 · 공개여부 미입력 등</span>'),
        ("임상 변수 리스트", "list", "#7a68e6", f"{n_var}개 변수",
         f'<span style="color:#6b7280">공개 {n_pub} • 비공개 {n_pri} • 미입력 {len(ECRF) - n_var}</span>'),
    ]
    md('<div class="kpis">' + "".join(
        f'<div class="card kpi"><div class="kpi-top">{t}{icon(ic, 22, c)}</div><div class="kpi-val">{v}</div>'
        f'<div class="kpi-sub">{s}</div></div>' for t, ic, c, v, s in kpis) + "</div>")

    with st.container(key="preview"):
        with st.container(key="prevtop"):
            a, b = st.columns([3, 1.1], vertical_alignment="center")
            with a:
                md(f'<div class="pv-title">{icon("table", 22, "#4e8f78")} iCReaT eCRF → CODA 코드북 매핑 미리보기</div>'
                   f'<div class="pv-sub">[{E(pid)}] eCRF 코드북을 정제하여 CODA 기탁 코드북(임상정보) 형식으로 변환한 결과입니다.</div>')
            with b:
                with st.container(key="pvfilter"):
                    sel = st.selectbox("eCRF", [f"전체 eCRF ({len(ecrf_names())}종)"] + ecrf_names(),
                                       label_visibility="collapsed", key=f"pv_sel_{pid}")
        chips = [("eCRF명", "테이블명(영문·국문)"), ("변수명", "변수명(영문)"), ("항목명", "변수 설명"),
                 ("데이터 유형", "변수 타입"), ("코드리스트", "변수값(코드) 설명")]
        ar = icon("arrow", 13, "#9aa1ad", 2)
        md('<div class="rules">' + "".join(f'<span class="rchip"><b>{s}</b>{ar}{t}</span>' for s, t in chips)
           + '<span class="rchip cur"><b>큐레이션</b> No. (ID·FAMILY_ID·PNTTM=0, 이후 1부터)</span>'
             '<span class="rchip cur"><b>연구자 입력</b> 공개여부</span></div>')
        coda_v, ecrf_v = filter_by_ecrf(sel)
        t1, t2 = st.tabs([f"CODA 코드북 미리보기 ({len(coda_v)})", f"iCReaT eCRF 원본 ({len(ecrf_v)})"])
        with t1:
            md(coda_table_html(coda_v))
        with t2:
            md(ecrf_table_html(ecrf_v))
    next_btn("3단계 · 과제 및 데이터 입력 폼으로", "form", "m")


# ══════════════════════════════════════════════════════════════
# 11. ③ 과제 및 데이터 입력 폼
# ══════════════════════════════════════════════════════════════
def bump():
    ss.ever[pid] = ss.ever.get(pid, 0) + 1


def cb_add(ms_key, pub_key):
    for uid in ss.get(ms_key, []):
        form["vars"][uid] = ss.get(pub_key, "공개") == "공개"
    ss[ms_key] = []
    bump()


def cb_add_ecrf(uids, pub_key):
    for uid in uids:
        form["vars"][uid] = ss.get(pub_key, "공개") == "공개"
    bump()


def cb_edit(ed_key, uids):
    dels = ss.dels.setdefault(pid, set())
    for i, ch in ss.get(ed_key, {}).get("edited_rows", {}).items():
        uid = uids[int(i)]
        if "공개여부" in ch:
            form["vars"][uid] = ch["공개여부"] == "공개"
        if "선택" in ch:
            (dels.add if ch["선택"] else dels.discard)(uid)
    bump()


def cb_bulk(mode):
    dels = ss.dels.setdefault(pid, set())
    if mode == "del":
        for uid in list(dels):
            form["vars"].pop(uid, None)
        dels.clear()
    else:
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
        md(f'<div class="sect auto">{icon("pdf", 19, "#4f6bdc")} [자동 입력] 과제정보 PDF 추출 데이터'
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
        md(f'<div class="sect user">{icon("check", 19, "#4e8f78")} [연구자 입력 폼] 없는 데이터 채우기 · 기탁 정보 입력</div>')
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

        # ── 임상 변수 리스트 ──
        md(f'<div class="subh">{icon("list", 18, "#4b5563")} 임상 변수 리스트 및 공개 여부 '
           f'<small>추가한 변수의 공개 여부가 CODA 코드북 \'공개여부\'에 반영됩니다</small></div>')
        ver = ss.ever.get(pid, 0)
        with st.container(key="varadd"):
            c1, c2, c3, c4, c5 = st.columns([1.3, 2.6, 1.2, 0.75, 1.25], vertical_alignment="bottom")
            with c1:
                ef = st.selectbox("eCRF", ecrf_names(), key=f"w_ef_{pid}")
            rem = [i for i, r in enumerate(ECRF) if r["eCRF"] == ef and i not in form["vars"]]
            ms_key, pub_key = f"w_ms_{pid}", f"w_pubnew_{pid}"
            with c2:
                st.multiselect("추가할 변수", rem, key=ms_key, placeholder=f"{len(rem)}개 선택 가능",
                               format_func=lambda i: f'{ECRF[i]["변수명"] or "(변수명 없음)"} · {ECRF[i]["항목명"]}')
            with c3:
                st.radio("공개 여부", ["공개", "비공개"], horizontal=True, key=pub_key)
            with c4:
                with st.container(key="addbtn"):
                    st.button(":material/add: 추가", use_container_width=True, on_click=cb_add, args=(ms_key, pub_key))
            with c5:
                st.button(f":material/playlist_add: {ef[:10]} 전체", use_container_width=True, disabled=not rem,
                          on_click=cb_add_ecrf, args=(rem, pub_key), help=f"{ef}의 미추가 변수 {len(rem)}개를 모두 추가")

        n_var, n_pub, n_pri = var_counts()
        md(f'<div class="cnt">임상 변수 개수 <span class="n">{n_var}개</span> = 공개 <b>{n_pub}</b> + 비공개 <b>{n_pri}</b>'
           f'<span class="f">= SUM(추가된 변수)</span><span style="margin-left:auto">미추가 eCRF 변수 '
           f'<b>{len(ECRF) - n_var}</b>개 → CODA 코드북에서 \'검토 필요\'</span></div>')

        uids = list(form["vars"].keys())
        dels = ss.dels.setdefault(pid, set())
        df = pd.DataFrame([{"선택": u in dels, "No.": n, "eCRF": ECRF[u]["eCRF"], "변수명": ECRF[u]["변수명"] or "(없음)",
                            "항목명": ECRF[u]["항목명"], "공개여부": "공개" if form["vars"][u] else "비공개"}
                           for n, u in enumerate(uids, 1)],
                          columns=["선택", "No.", "eCRF", "변수명", "항목명", "공개여부"])
        ed_key = f"w_ed_{pid}_{ver}"
        with st.container(key="vared"):
            st.data_editor(df, key=ed_key, hide_index=True, use_container_width=True, num_rows="fixed",
                           height=min(38 + 35 * max(len(df), 1), 420), on_change=cb_edit, args=(ed_key, uids),
                           disabled=["No.", "eCRF", "변수명", "항목명"],
                           column_config={"선택": st.column_config.CheckboxColumn("선택", width="small"),
                                          "No.": st.column_config.NumberColumn("No.", width="small"),
                                          "공개여부": st.column_config.SelectboxColumn("공개여부", options=["공개", "비공개"],
                                                                                   required=True, width="small")})
        with st.container(key="varbulk"):
            c1, c2, c3, _ = st.columns([1.1, 1, 1, 2.5])
            with c1:
                with st.container(key="delbtn"):
                    st.button(f":material/delete: 선택 삭제 ({len(dels)})", use_container_width=True, disabled=not dels,
                              on_click=cb_bulk, args=("del",))
            with c2:
                st.button("전체 공개", use_container_width=True, on_click=cb_bulk, args=("pub",))
            with c3:
                st.button("전체 비공개", use_container_width=True, on_click=cb_bulk, args=("pri",))
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
                rr["변수값(코드) 설명"], rr["변수 타입"], None, ", ".join(rr["_issues"]) or None]
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


def page_result():
    text, html_ = info_values()
    filled = sum(1 for _, _, k, _ in INFO_ROWS if text[k] or k == "desc")
    with st.container(key="resultcard"):
        with st.container(key="restop"):
            a, b = st.columns([3, 1.15], vertical_alignment="center")
            with a:
                md(f'<div class="pg-head">{icon("layers", 28, "#4e8f78")}<div><div class="h">결과 확인 — CODA 기탁용 임상정보 코드북</div>'
                   f'<div class="s">②매핑 결과와 ③입력 폼 내용을 CODA 기탁 양식 그대로 합쳤습니다. 내려받아 CODA에 제출하세요.</div></div></div>')
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
           f'<span class="prog">매핑 완료 {CNT["map"]} · 정제 완료 {CNT["cln"]} · 검토 필요 {CNT["rev"]}</span></div>')
        c1, c2, _ = st.columns([1.2, 1.3, 2.5], vertical_alignment="center")
        with c1:
            sel = st.selectbox("eCRF", [f"전체 eCRF ({len(ecrf_names())}종)"] + ecrf_names(),
                               label_visibility="collapsed", key=f"res_sel_{pid}")
        with c2:
            only = st.radio("보기", ["전체", "검토 필요만"], horizontal=True, label_visibility="collapsed", key=f"res_only_{pid}")
        rows, _ = filter_by_ecrf(sel)
        if only == "검토 필요만":
            rows = [r for r in rows if r["_status"] == "rev"]
        md(coda_table_html(rows))


# ══════════════════════════════════════════════════════════════
# 13. 본문
# ══════════════════════════════════════════════════════════════
with st.container(key="body"):
    left, main = st.columns([1, 3.1])
    with left:
        left_panel()
    with main:
        {"mapping": page_mapping, "form": page_form, "result": page_result}.get(ss.page, page_mapping)()
