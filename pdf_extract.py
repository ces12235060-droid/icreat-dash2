"""iCReaT '과제 정보' 화면 PDF(Ctrl+P → PDF로 저장) → 과제 메타데이터 추출

OCR 없이 PDF 내부 텍스트 레이어를 읽는다 (pdfminer.six).
브라우저 인쇄로 저장한 PDF는 텍스트가 그대로 들어 있으므로 OCR이 필요 없다.
스캔·이미지 PDF처럼 텍스트가 없으면 scanned=True 를 돌려주고, 화면에서 직접 입력하도록 안내한다.

표 형태(라벨 셀 | 값 셀) 화면을 위해 좌표 기반으로 라벨 오른쪽·같은 행 범위의 텍스트를 값으로 묶는다.
"""
import io
import logging
import re

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextContainer, LTTextLine

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# 필드 → iCReaT 화면 라벨 후보 (공백 제거 후 비교). 실제 화면 라벨이 다르면 여기에 추가
FIELD_ALIASES = {
    "pid": ["과제번호", "연구번호", "과제코드"],
    "title": ["과제명(국문)", "과제명", "연구과제명", "연구제목(국문)", "연구제목", "연구명"],
    "pi": ["연구책임자", "과제책임자", "책임연구자"],
    "org": ["과제수행기관", "수행기관", "소속기관", "연구기관", "주관기관"],
    "period": ["연구기간", "과제기간", "수행기간", "과제수행기간"],
    "goal": ["연구목적", "과제목표", "연구목표"],
    "content": ["연구내용", "과제주요내용", "주요내용", "연구내용요약"],
    "disease": ["대상질환", "대상질환명", "연구질환명", "질환명"],
    "n_subjects": ["목표대상자수", "대상자수", "목표피험자수", "예정대상자수"],
}
FIELD_LABELS = {"pid": "과제 번호", "title": "과제명", "pi": "과제 책임자", "org": "과제 수행기관", "period": "과제 수행기간",
                "goal": "과제 목표", "content": "과제 주요내용", "disease": "연구 질환명", "n_subjects": "대상자 수"}
# 라벨로만 쓰이고 값으로 오면 안 되는 문구 (다른 셀의 라벨)
OTHER_LABELS = ["과제명(영문)", "연구제목(영문)", "과제상태", "연구유형", "연구단계", "등록일", "IRB승인번호"]

_ALIAS = {}
for f, al in FIELD_ALIASES.items():
    for a in al:
        _ALIAS[re.sub(r"\s+", "", a)] = f
_OTHER = {re.sub(r"\s+", "", a) for a in OTHER_LABELS}


def _n(s):
    return re.sub(r"\s+", "", s or "")


def _is_hangul(ch):
    return "가" <= ch <= "힣"


def _join(parts):
    out = ""
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if out and not (_is_hangul(out[-1]) and _is_hangul(p[0])):
            out += " "
        out += p
    return out


def _lines(data):
    pages = []
    for page in extract_pages(io.BytesIO(data), laparams=LAParams(line_margin=0.3)):
        items = []
        for el in page:
            if isinstance(el, LTTextContainer):
                for ln in el:
                    if isinstance(ln, LTTextLine) and ln.get_text().strip():
                        items.append(dict(x0=ln.x0, x1=ln.x1, y0=ln.y0, y1=ln.y1, yc=(ln.y0 + ln.y1) / 2,
                                          t=ln.get_text().strip()))
        pages.append((page.width, items))
    return pages


def extract_project_info(data):
    """PDF bytes → dict(values, found, scanned, n_chars, text)"""
    try:
        pages = _lines(data)
    except Exception as e:  # 손상/암호화 PDF
        return dict(values={}, found={}, scanned=False, error=f"PDF를 읽지 못했습니다: {e}", n_chars=0, text="")
    text = "\n".join(it["t"] for _, items in pages for it in items)
    n_chars = len(_n(text))
    if n_chars < 20:
        return dict(values={}, found={}, scanned=True, error="", n_chars=n_chars, text=text)

    values, found = {}, {}
    for width, items in pages:
        labels = []
        for it in items:
            k = _n(it["t"])
            f = _ALIAS.get(k)
            inline = ""
            if not f:  # '라벨: 값' 같은 줄
                m = re.match(r"^(.{2,14}?)\s*[:：]\s*(.+)$", it["t"])
                if m and _n(m.group(1)) in _ALIAS:
                    f, inline = _ALIAS[_n(m.group(1))], m.group(2).strip()
            if f or k in _OTHER:
                labels.append(dict(it, f=f, inline=inline))
        lab_ids = {id(x) for x in labels}
        rows_y = sorted({round(lb["yc"]) for lb in labels}, reverse=True)
        for lb in labels:
            if not lb["f"] or lb["f"] in values:
                continue
            if lb["inline"]:
                values[lb["f"]], found[lb["f"]] = lb["inline"], lb["t"]
                continue
            yc = round(lb["yc"])
            above = [y for y in rows_y if y > yc + 3]
            below = [y for y in rows_y if y < yc - 3]
            top = (yc + min(above)) / 2 if above else lb["y1"] + 40
            bottom = (yc + max(below)) / 2 if below else lb["y0"] - 40
            right_labels = [o["x0"] for o in labels if o is not lb and abs(o["yc"] - lb["yc"]) < 4 and o["x0"] > lb["x1"]]
            right = min(right_labels) if right_labels else width
            vals = [it for it in items
                    if bottom < it["yc"] < top and it["x0"] >= lb["x1"] - 2 and it["x1"] <= right + 2
                    and _n(it["t"]) not in _ALIAS and _n(it["t"]) not in _OTHER]
            vals.sort(key=lambda it: (-round(it["yc"]), it["x0"]))
            v = _join([it["t"] for it in vals])
            if v:
                values[lb["f"]], found[lb["f"]] = v, lb["t"]

    if "pid" not in values:
        m = re.search(r"\b([A-Z]\d{6})\b", text)
        if m:
            values["pid"], found["pid"] = m.group(1), "(본문 패턴)"
    if "n_subjects" in values:
        m = re.search(r"[\d,]+", values["n_subjects"])
        values["n_subjects"] = int(m.group().replace(",", "")) if m else values["n_subjects"]
    return dict(values=values, found=found, scanned=False, error="", n_chars=n_chars, text=text)
