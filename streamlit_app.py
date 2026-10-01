import streamlit as st

from overlap_core import analyze, to_excel_bytes

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
