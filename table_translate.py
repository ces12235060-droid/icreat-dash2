"""CODA 테이블명(영문) → 테이블명(국문) 자동 번역 (오프라인 · 규칙 기반)

내부망에서 외부 번역 API를 쓸 수 없으므로 임상 용어 사전 + 규칙으로 번역한다.
  1) 전체 문구 사전 매칭  (예: 'Laboratory Values' → '검사 결과')
  2) CamelCase / '_' / '-' 분리 후 긴 구절부터 사전 매칭  (예: 'BaseLineInfo' → 'Base Line Info' → '기본 정보')
  3) 사전에 없는 단어는 철자 유사도로 보정 (예: 'dringking' → 'drinking' → '음주')
  4) 접미사 규칙: '_COPY' → '(사본)', 끝의 '_' → '(추가)'
  5) 그래도 남는 영문 단어는 그대로 두고 status='검토' 로 표시 → 화면에서 수정
한글 eCRF명은 그대로 사용한다. 새 용어는 GLOSSARY 에 추가하면 된다.
"""
import difflib
import re

# 소문자 키. 여러 단어 구절을 먼저 정의 (긴 구절 우선 매칭)
GLOSSARY = {
    # 구절
    "baseline information": "기본정보", "base line information": "기본정보", "base line info": "기본정보",
    "baseline info": "기본정보", "base line": "기본정보",
    "past medical history": "과거 병력", "past-medical history": "과거 병력", "medical history": "병력",
    "family history": "가족력", "laboratory values": "검사 결과", "laboratory test": "임상검사", "lab test": "임상검사",
    "vital signs": "활력징후", "vital sign": "활력징후", "physical examination": "신체검사", "physical exam": "신체검사",
    "adverse event": "이상반응", "adverse events": "이상반응", "serious adverse event": "중대한 이상반응",
    "study drug": "연구 약물", "concomitant medication": "병용 약물", "prior medication": "이전 약물",
    "informed consent": "동의서", "inclusion exclusion": "선정 제외 기준", "inclusion criteria": "선정 기준",
    "exclusion criteria": "제외 기준", "end of study": "연구 종료", "follow up": "추적관찰", "follow-up": "추적관찰",
    "case report form": "증례기록서", "quality of life": "삶의 질", "subject information": "대상자 정보",
    # 단어
    "baseline": "기본정보", "info": "정보", "information": "정보", "demographics": "인구학적 정보",
    "demographic": "인구학적 정보", "laboratory": "임상검사", "lab": "임상검사", "values": "결과", "value": "결과",
    "drinking": "음주", "alcohol": "음주", "smoking": "흡연", "tobacco": "흡연", "history": "이력",
    "past": "과거", "medical": "의학", "and": "및", "or": "또는", "ae": "이상반응", "sae": "중대한 이상반응",
    "drug": "약물", "medication": "약물", "study": "연구", "visit": "방문", "screening": "스크리닝",
    "randomization": "무작위배정", "enrollment": "등록", "registration": "등록", "consent": "동의",
    "diagnosis": "진단", "treatment": "치료", "therapy": "치료", "surgery": "수술", "transplant": "이식",
    "transplantation": "이식", "kidney": "신장", "renal": "신장", "liver": "간", "heart": "심장",
    "blood": "혈액", "urine": "소변", "imaging": "영상", "image": "영상", "genetic": "유전", "genomic": "유전체",
    "biomarker": "바이오마커", "sample": "검체", "specimen": "검체", "questionnaire": "설문", "survey": "설문",
    "score": "점수", "outcome": "결과", "result": "결과", "results": "결과", "test": "검사", "exam": "검사",
    "examination": "검사", "assessment": "평가", "evaluation": "평가", "grid": "그리드", "repeat": "반복",
    "general": "일반", "summary": "요약", "comment": "의견", "comments": "의견", "subject": "대상자",
    "patient": "환자", "termination": "종료", "completion": "완료", "death": "사망", "pregnancy": "임신",
    "nutrition": "영양", "diet": "식이", "exercise": "운동", "sleep": "수면", "symptom": "증상", "symptoms": "증상",
    "ecg": "심전도", "echo": "심초음파", "echocardiography": "심초음파", "pathology": "병리", "biopsy": "생검",
}
SUFFIX_RULES = [(r"[_\s-]*copy$", "(사본)"), (r"_+$", "(추가)"), (r"[_\s-]*v(\d+)$", r"(v\1)")]
_WORDS = sorted({w for k in GLOSSARY for w in k.split()})


def _split(name):
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)          # CamelCase 분리
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    return [t for t in re.split(r"[\s_\-/]+", s.lower()) if t]


def _fix(tok):
    if tok in GLOSSARY or tok in _WORDS or tok.isdigit():
        return tok, False
    m = difflib.get_close_matches(tok, _WORDS, n=1, cutoff=0.82)
    return (m[0], True) if m else (tok, False)


def translate_table_name(name):
    """→ dict(ko, status: 자동|보정|검토|원문, note)"""
    raw = (name or "").strip()
    if not raw:
        return dict(ko="", status="검토", note="이름 없음")
    if re.search(r"[가-힣]", raw):
        return dict(ko=raw, status="원문", note="한글 eCRF명 그대로 사용")
    base, suffix = raw, ""
    for pat, rep in SUFFIX_RULES:
        m = re.search(pat, base, flags=re.I)
        if m:
            suffix = re.sub(pat, rep, m.group(0), flags=re.I).strip() if "\\1" in rep else rep
            base = base[:m.start()]
            break
    whole = base.strip().lower()
    if whole in GLOSSARY:
        return dict(ko=GLOSSARY[whole] + suffix, status="자동", note="용어 사전")
    toks, fixed = [], False
    for t in _split(base):
        t2, f = _fix(t)
        toks.append(t2)
        fixed |= f
    out, i, unknown = [], 0, []
    while i < len(toks):  # 긴 구절 우선 매칭
        for L in range(min(4, len(toks) - i), 0, -1):
            ph = " ".join(toks[i:i + L])
            if ph in GLOSSARY:
                out.append(GLOSSARY[ph])
                i += L
                break
        else:
            out.append(toks[i])
            if not toks[i].isdigit():
                unknown.append(toks[i])
            i += 1
    ko = re.sub(r"\s+", " ", " ".join(out)).strip()
    ko = re.sub(r"(정보) \1", r"\1", ko)
    ko = ko.replace("기본정보 정보", "기본정보").replace("과거 의학 이력", "과거 병력").replace("의학 이력", "병력")
    ko += suffix
    if unknown:
        return dict(ko=ko, status="검토", note="사전에 없는 단어: " + ", ".join(unknown))
    return dict(ko=ko, status="보정" if fixed else "자동", note="철자 보정 후 번역" if fixed else "규칙 번역")
