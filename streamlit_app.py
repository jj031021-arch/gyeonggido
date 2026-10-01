
import importlib.util
import subprocess
import sys

# 배포 환경에 openpyxl이 없으면 자동 설치 (requirements.txt가 적용되지 않은 경우 대비)
for _pkg in ("openpyxl",):
    if importlib.util.find_spec(_pkg) is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", _pkg])

import io
import re
from collections import Counter
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

# ===================== 비교 로직 =====================
def _clean(t) -> str:
    return re.sub(r"[\(\)\[\]\-_/·,.]", " ", str(t).lower())


def _bigrams(s: str) -> set:
    return {s[i:i + 2] for i in range(len(s) - 1)} or {s}


def similarity(a: str, b: str) -> float:
    """글자 2-gram 자카드 + 순서 기반 비율 + 글자 Dice 의 평균 (0~1)"""
    if not a or not b:
        return 0.0
    ba, bb = _bigrams(a), _bigrams(b)
    jac = len(ba & bb) / len(ba | bb)
    seq = SequenceMatcher(None, a, b).ratio()
    dice = 2 * len(set(a) & set(b)) / (len(set(a)) + len(set(b)))
    return (jac + seq + dice) / 3


def read_files(files, col_no, col_member, col_item, match_tabs_by):
    """files: [(파일명, bytes)] → ({탭키: [행…]}, [경고…])"""
    tab_rows, warnings = {}, []
    for fname, content in files:
        name = fname.rsplit(".", 1)[0]
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None)
        for idx, (sheet, df) in enumerate(sheets.items(), start=1):
            missing = [c for c in (col_no, col_member, col_item) if c not in df.columns]
            if missing:
                warnings.append(f"{name} 파일의 '{sheet}' 탭에 {missing} 열이 없어 건너뜀")
                continue
            key = {"name": sheet, "order": f"{idx}번째 탭", "all": "전체"}[match_tabs_by]
            for _, r in df.dropna(subset=[col_item]).iterrows():
                tab_rows.setdefault(key, []).append({
                    "파일": name, "탭": sheet, "연번": r[col_no],
                    "의원명": r[col_member], "자료명": str(r[col_item])})
    return tab_rows, warnings


def find_groups(rows, threshold, min_files, auto_stopwords=True,
                auto_ratio=0.3, manual_stopwords=()):
    stop = set(manual_stopwords)
    if auto_stopwords and rows:
        cnt = Counter(w for r in rows for w in set(_clean(r["자료명"]).split()))
        stop |= {w for w, c in cnt.items() if c / len(rows) >= auto_ratio}

    def normalize(t):
        toks = _clean(t).split()
        s = "".join(w for w in toks if w not in stop)
        for w in manual_stopwords:
            s = s.replace(w, "")
        return s or "".join(toks)

    for r in rows:
        r["_norm"] = normalize(r["자료명"])

    parent = list(range(len(rows)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i in range(len(rows)):
        for j in range(i + 1, len(rows)):
            if rows[i]["파일"] != rows[j]["파일"] and \
               similarity(rows[i]["_norm"], rows[j]["_norm"]) >= threshold:
                parent[find(i)] = find(j)

    groups = {}
    for i, r in enumerate(rows):
        groups.setdefault(find(i), []).append(r)
    return [g for g in groups.values() if len({x["파일"] for x in g}) >= min_files]


def analyze(files, col_no="연번", col_member="의원명", col_item="자료명",
            match_tabs_by="name", threshold=0.6, min_files=2,
            auto_stopwords=True, auto_ratio=0.3, manual_stopwords=()):
    """→ (요약 DataFrame, 상세 DataFrame, 경고 리스트)"""
    tab_rows, warnings = read_files(files, col_no, col_member, col_item, match_tabs_by)
    summary, detail, gid = [], [], 0
    for key, rows in tab_rows.items():
        if len({r["파일"] for r in rows}) < min_files:
            continue
        for g in find_groups(rows, threshold, min_files, auto_stopwords,
                             auto_ratio, manual_stopwords):
            gid += 1
            rep = min((x["자료명"] for x in g), key=len)
            members = sorted({str(x["의원명"]) for x in g})
            fs = sorted({x["파일"] for x in g})
            summary.append({"그룹": gid, "탭": key, "대표자료명": rep,
                            "겹친 파일 수": len(fs), "요구 의원 수": len(members),
                            "요구 의원": ", ".join(members), "겹친 파일": ", ".join(fs)})
            for x in sorted(g, key=lambda x: (x["파일"], str(x["연번"]).zfill(8))):
                detail.append({"그룹": gid, "탭": key, "대표자료명": rep, "파일": x["파일"],
                               "연번": x["연번"], "의원명": x["의원명"],
                               "자료명(원문)": x["자료명"]})
    s_cols = ["그룹", "탭", "대표자료명", "겹친 파일 수", "요구 의원 수", "요구 의원", "겹친 파일"]
    d_cols = ["그룹", "탭", "대표자료명", "파일", "연번", "의원명", "자료명(원문)"]
    s = pd.DataFrame(summary, columns=s_cols)
    if len(s):
        s = s.sort_values(["겹친 파일 수", "요구 의원 수", "그룹"],
                          ascending=[False, False, True]).reset_index(drop=True)
    return s, pd.DataFrame(detail, columns=d_cols), warnings


def to_excel_bytes(summary: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="요약", index=False)
        detail.to_excel(w, sheet_name="상세(의원+연번+자료명)", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col) * 1.6 + 2
                ws.column_dimensions[col[0].column_letter].width = min(50, width)
            ws.freeze_panes = "A2"
    return buf.getvalue()


# ===================== 화면 =====================
st.set_page_config(page_title="요구자료 겹침 찾기", page_icon="📑", layout="wide")
st.title("📑 요구자료 겹침 찾기")
st.caption("엑셀 여러 개를 올리면 자료명이 겹치는(비슷한) 항목을 의원명·연번과 함께 뽑아줍니다.")

with st.sidebar:
    st.header("⚙️ 설정")
    col_no = st.text_input("연번 열 이름", "연번")
    col_member = st.text_input("의원명 열 이름", "의원명")
    col_item = st.text_input("자료명 열 이름", "자료명")
    tab_label = st.radio("탭 비교 방식", ["같은 이름의 탭끼리", "같은 순서의 탭끼리", "탭 구분 없이 전체"])
    match_tabs_by = {"같은 이름의 탭끼리": "name", "같은 순서의 탭끼리": "order",
                     "탭 구분 없이 전체": "all"}[tab_label]
    threshold = st.slider("유사도 기준", 0.3, 1.0, 0.6, 0.05,
                          help="높을수록 엄격(비슷한 것만), 낮을수록 느슨(더 많이 묶임)")
    min_files = st.number_input("최소 겹침 파일 수", 2, 20, 2,
                                help="예: 4로 하면 4개 파일 모두에서 겹친 것만")
    with st.expander("불용어 (선택)"):
        auto_stop = st.checkbox("자주 나오는 단어 자동 제외", True)
        manual = st.text_input("직접 지정 (쉼표로 구분)", "", placeholder="사본, 제출")

uploaded = st.file_uploader("엑셀 파일 업로드 (여러 개 선택)", type=["xlsx", "xls"],
                            accept_multiple_files=True)

if len(uploaded) < 2:
    st.info("엑셀 파일을 2개 이상 올려 주세요. 예시 파일은 저장소의 `examples/` 폴더에 있습니다.")
    st.stop()

files = [(f.name, f.getvalue()) for f in uploaded]
manual_list = [w.strip() for w in manual.split(",") if w.strip()]

with st.spinner("비교 중..."):
    summary, detail, warnings = analyze(
        files, col_no, col_member, col_item, match_tabs_by, threshold,
        int(min_files), auto_stop, 0.3, manual_list)

for w in warnings:
    st.warning(w)
if summary.empty:
    st.error("겹치는 자료가 없습니다. 유사도 기준을 낮추거나 열 이름을 확인해 주세요.")
    st.stop()

n_files = len(uploaded)
c1, c2, c3 = st.columns(3)
c1.metric("겹치는 자료", f"{len(summary)}건")
c2.metric("모든 파일에서 겹침", f"{(summary['겹친 파일 수'] == n_files).sum()}건")
c3.metric("관련 요구 건수", f"{len(detail)}건")

st.download_button("⬇️ 결과 엑셀 다운로드", to_excel_bytes(summary, detail),
                   file_name="겹치는_자료_의원별.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary")

tab1, tab2, tab3 = st.tabs(["📋 요약", "🔍 자료별 상세", "👤 의원별"])

with tab1:
    st.bar_chart(summary.set_index("대표자료명")[["겹친 파일 수", "요구 의원 수"]])
    st.dataframe(summary, hide_index=True)

with tab2:
    for _, s in summary.iterrows():
        with st.expander(f"{s['대표자료명']}  ·  {s['겹친 파일 수']}개 파일 / {s['요구 의원 수']}명  ({s['탭']})"):
            st.dataframe(detail[detail["그룹"] == s["그룹"]][["파일", "연번", "의원명", "자료명(원문)"]],
                         hide_index=True)

with tab3:
    per_member = detail.groupby("의원명").size().sort_values(ascending=False)
    st.caption("다른 의원과 겹치는 자료를 몇 건 요구했는지")
    st.bar_chart(per_member)
    pick = st.selectbox("의원 선택", per_member.index)
    st.dataframe(detail[detail["의원명"] == pick][["대표자료명", "파일", "연번", "자료명(원문)"]],
                 hide_index=True)
