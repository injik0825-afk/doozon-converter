"""
streamlit_app.py
분개장 자동생성 앱 메인 UI

실행: streamlit run streamlit_app.py
"""
import streamlit as st
import pandas as pd
from datetime import datetime

from app.companies import get_company_config, list_companies
from app.core.portfolio import Portfolio
from app.core.converter import Converter
from app.core.month_end import MonthEndProcessor
from app.parsers import get_parser
from app.utils import (
    load_excel_sheets, export_journal_to_duzone,
    Flag, ST_BG_COLOR, LABEL_KO, PRIORITY,
)


st.set_page_config(
    page_title="분개장 자동생성",
    page_icon="📒",
    layout="wide",
)

st.title("📒 분개장 자동생성 시스템")
st.caption("거래내역 파일만 올리면 분개장을 자동으로 생성합니다")


# ========================================
# 사이드바: 기본 설정
# ========================================
with st.sidebar:
    st.header("⚙️ 기본 설정")

    companies = list_companies()
    company_code = st.selectbox(
        "회사 선택",
        options=list(companies.keys()),
        format_func=lambda x: companies[x],
    )

    col1, col2 = st.columns(2)
    with col1:
        year = st.number_input("결산 연도", min_value=2020, max_value=2099,
                                value=datetime.now().year)
    with col2:
        month = st.number_input("결산 월", min_value=1, max_value=12,
                                value=datetime.now().month)

    st.divider()
    st.caption(f"🏢 {companies[company_code]}")
    st.caption(f"📅 {year}년 {month}월 결산")

    st.divider()
    with st.expander("🎨 색상 안내"):
        st.markdown("""
        분개 결과에 다음 색상이 표시됩니다:

        - 🔴 **빨강**: 자동분개 실패 / 차대변 불일치 → 수동 입력 필요
        - 🟡 **노랑**: 취득가액 없음 → 처분손익 검토
        - 🔵 **파랑**: 신규 종목 → 거래처 등록 필요
        - 🟠 **주황**: 계정/거래처 추정 → 검토
        - ⚪ **회색**: 시스템 자동생성 (정상)
        """)


# ========================================
# 탭 구성
# ========================================
tab1, tab2, tab3, tab4 = st.tabs(
    ["1. 기초 포트폴리오", "2. 거래내역 업로드", "3. 월말평가", "4. 결과 확인"]
)


# ========================================
# Tab 1: 기초 포트폴리오
# ========================================
with tab1:
    st.header("📊 전월말 기준 포트폴리오")
    st.markdown("""
    매도 시 처분이익을 정확히 계산하려면 **전월말 종목별 보유수량·평균단가·평가금액**이 필요합니다.

    💡 처음 시작하는 경우 빈 상태로 두고 **2단계**로 이동해도 됩니다 (취득가액이 없는 매도는 노란색으로 표시됨).
    """)

    upload_mode = st.radio("입력 방법",
                            ["엑셀 업로드", "수동 입력", "빈 상태로 시작"],
                            horizontal=True)

    if 'portfolio_data' not in st.session_state:
        st.session_state['portfolio_data'] = []

    if upload_mode == "엑셀 업로드":
        pf_file = st.file_uploader(
            "기초 포트폴리오 엑셀 (종목명, 종목유형, 시장구분, 수량, 평균단가, 전월말_평가금액, 거래처코드)",
            type=['xlsx'], key='pf_upload'
        )
        if pf_file:
            pf_df = pd.read_excel(pf_file)
            st.session_state['portfolio_data'] = pf_df.to_dict('records')
            st.success(f"✅ {len(pf_df)}개 종목 로드")
            st.dataframe(pf_df, use_container_width=True)

    elif upload_mode == "수동 입력":
        st.info("테이블에서 직접 종목 추가/편집")
        edited = st.data_editor(
            pd.DataFrame(st.session_state['portfolio_data']) if st.session_state['portfolio_data']
            else pd.DataFrame({
                '종목명': [''], '종목유형': ['주식'], '시장구분': ['코스피'],
                '수량': [0], '평균단가': [0.0], '전월말_평가금액': [0],
                '거래처코드': [''],
            }),
            num_rows='dynamic',
            use_container_width=True,
            key='pf_editor',
        )
        if st.button("포트폴리오 저장"):
            st.session_state['portfolio_data'] = edited.to_dict('records')
            st.success("저장됨")

    else:
        st.info("4월 결산이면 3월말 포트폴리오가 필요합니다. 없어도 진행 가능하나 처분이익 계산이 부정확합니다 (해당 라인은 노란색 표시).")
        st.session_state['portfolio_data'] = []

    st.caption(f"현재 {len(st.session_state['portfolio_data'])}개 종목")


# ========================================
# Tab 2: 거래내역 업로드
# ========================================
with tab2:
    st.header("📁 거래내역 파일 업로드")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("🏦 은행/카드")
        bank_card_file = st.file_uploader(
            "은행·카드 거래내역", type=['xlsx'], key='bank_upload')
    with col2:
        st.subheader("📈 증권사")
        sec_file = st.file_uploader(
            "증권사 거래내역", type=['xlsx'], key='sec_upload')
    with col3:
        st.subheader("📋 청약 (선택)")
        ipo_file = st.file_uploader(
            "공모주 청약내역", type=['xlsx'], key='ipo_upload')

    st.divider()

    if st.button("🚀 변환 실행", type='primary', use_container_width=True):
        if not bank_card_file and not sec_file:
            st.error("최소 하나의 거래내역 파일을 업로드하세요.")
        else:
            with st.spinner("거래내역을 분개장으로 변환하는 중..."):
                # 1. 회사 설정 + 포트폴리오
                company = get_company_config(company_code)
                portfolio = Portfolio()
                if st.session_state['portfolio_data']:
                    portfolio.load_opening_positions(st.session_state['portfolio_data'])

                # 2. Converter
                converter = Converter(company, portfolio)

                # 3. 월초 평가상계
                month_end = MonthEndProcessor(company, portfolio, converter.book)
                month_end.generate_opening_reversal(year, month)

                # 4. 파일 파싱
                all_events = []
                with st.expander("📋 파싱 로그", expanded=True):
                    if bank_card_file:
                        sheets = load_excel_sheets(bank_card_file.read())
                        for sheet_name, df in sheets.items():
                            if df.empty:
                                continue
                            if 'IBK기업은행' in sheet_name or 'IBK 기업은행' in sheet_name:
                                events = get_parser('ibk_bank').parse(df, account_id='IBK_메인')
                                all_events.extend(events)
                                st.write(f"- {sheet_name}: {len(events)}건")
                            elif 'IBK카드' in sheet_name:
                                events = get_parser('ibk_card').parse(df)
                                all_events.extend(events)
                                st.write(f"- {sheet_name}: {len(events)}건")
                            elif sheet_name == '카드이용내역':
                                events = get_parser('card_summary').parse(df)
                                all_events.extend(events)
                                st.write(f"- {sheet_name}: {len(events)}건")

                    if sec_file:
                        sheets = load_excel_sheets(sec_file.read())
                        for sheet_name, df in sheets.items():
                            if df.empty:
                                continue
                            sec_account = company.get_securities_by_pattern(sheet_name)
                            if sec_account is None:
                                continue
                            if '키움' in sheet_name:
                                parser = get_parser('kiwoom')
                            elif '교보' in sheet_name:
                                parser = get_parser('kyobo')
                            elif '한국' in sheet_name or '한투' in sheet_name:
                                parser = get_parser('hanto')
                            elif '메리츠' in sheet_name:
                                parser = get_parser('meritz')
                            else:
                                st.warning(f"⚠️ {sheet_name}: 미지원 증권사")
                                continue
                            events = parser.parse(df, account_id=sec_account.별칭)
                            all_events.extend(events)
                            st.write(f"- {sheet_name}: {len(events)}건")

                    st.info(f"**총 이벤트: {len(all_events)}건**")

                # 5. 변환
                converter.convert(all_events)

                # 세션 저장
                st.session_state['converter'] = converter
                st.session_state['portfolio'] = portfolio
                st.session_state['company'] = company

                st.success(f"✅ 변환 완료! 총 {len(converter.book.transactions)}개 거래")

                # 플래그 요약 즉시 표시
                flag_counts = converter.book.flag_counts()
                if flag_counts:
                    st.markdown("##### 시각적 경고 표시")
                    cols = st.columns(len(flag_counts) if len(flag_counts) <= 5 else 5)
                    for i, (f, cnt) in enumerate(flag_counts.items()):
                        col = cols[i % len(cols)]
                        col.markdown(
                            f'<div style="background-color:{ST_BG_COLOR[f]};'
                            f'padding:8px;border-radius:5px;text-align:center;">'
                            f'<b>{LABEL_KO[f]}</b><br/>{cnt}건</div>',
                            unsafe_allow_html=True
                        )


# ========================================
# Tab 3: 월말평가
# ========================================
with tab3:
    st.header("📈 월말평가 가격 입력")
    st.markdown("월말 종가를 입력하면 평가손익을 자동으로 분개합니다.")

    if 'portfolio' not in st.session_state:
        st.info("먼저 **2단계**에서 거래내역을 변환하세요.")
    else:
        portfolio = st.session_state['portfolio']
        snapshot = portfolio.snapshot()

        if snapshot.empty:
            st.info("보유 중인 종목이 없습니다.")
        else:
            snapshot['월말종가'] = 0.0

            edited_prices = st.data_editor(
                snapshot[['종목명', '수량', '평균단가', '장부가액', '월말종가', '신규']],
                use_container_width=True,
                disabled=['종목명', '수량', '평균단가', '장부가액', '신규'],
                key='price_editor',
            )

            if st.button("🧮 월말평가 분개 생성", type='primary'):
                current_prices = dict(zip(
                    edited_prices['종목명'],
                    edited_prices['월말종가']
                ))
                current_prices = {k: v for k, v in current_prices.items() if v > 0}

                if not current_prices:
                    st.warning("월말종가를 입력한 종목이 없습니다.")
                else:
                    converter = st.session_state['converter']
                    company = st.session_state['company']
                    processor = MonthEndProcessor(company, portfolio, converter.book)
                    processor.generate_month_end_valuation(year, month, current_prices)
                    st.success(f"✅ 월말평가 분개 생성 완료 ({len(current_prices)}개 종목)")


# ========================================
# Tab 4: 결과 확인/다운로드
# ========================================
with tab4:
    st.header("📊 결과 확인 및 다운로드")

    if 'converter' not in st.session_state:
        st.info("먼저 **2단계**에서 거래내역을 변환하세요.")
    else:
        converter = st.session_state['converter']
        book = converter.book
        summary = book.summary()

        # 기본 요약
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("총 거래 수", f"{summary['총 거래 수']:,}")
        col2.metric("총 분개 라인", f"{summary['총 분개 라인']:,}")
        col3.metric("차변 합계", f"{summary['차변 합계']:,}")
        col4.metric("대변 합계", f"{summary['대변 합계']:,}")

        if summary['차/대변 일치']:
            st.success("✅ 차변/대변 일치")
        else:
            st.error(f"❌ 차/대변 불일치 (차이: {summary['차변 합계'] - summary['대변 합계']:,})")

        # 시각적 경고 요약 (색상 표시)
        flag_counts = book.flag_counts()
        if flag_counts:
            st.divider()
            st.subheader("⚠️ 검토 필요 사항 요약")
            
            cols = st.columns(min(len(flag_counts), 4))
            for i, (f, cnt) in enumerate(flag_counts.items()):
                with cols[i % len(cols)]:
                    st.markdown(
                        f'<div style="background-color:{ST_BG_COLOR[f]};'
                        f'padding:12px;border-radius:8px;text-align:center;'
                        f'margin-bottom:8px;font-size:14px;">'
                        f'<b>{LABEL_KO[f]}</b><br/>'
                        f'<span style="font-size:24px;">{cnt}</span> 건</div>',
                        unsafe_allow_html=True
                    )

        # 거래 유형별 통계
        st.divider()
        st.subheader("거래 유형별 분포")
        type_df = pd.DataFrame(
            [(k, v) for k, v in summary['거래 유형별 건수'].items()],
            columns=['거래 유형', '건수']
        ).sort_values('건수', ascending=False)
        st.dataframe(type_df, use_container_width=True, hide_index=True)

        # 검토 필요 라인만 보기
        if flag_counts:
            st.divider()
            st.subheader("🔍 검토 필요 라인 (색상별)")
            
            # 플래그 필터
            flag_options = [f.value for f in flag_counts.keys()]
            selected_flags = st.multiselect(
                "표시할 플래그 선택",
                options=flag_options,
                default=flag_options,
                format_func=lambda v: LABEL_KO[Flag(v)],
            )
            
            if selected_flags:
                selected_flag_set = {Flag(v) for v in selected_flags}
                flagged_entries = [
                    e for e in book.all_entries()
                    if any(f in selected_flag_set for f in e.flags)
                ]
                
                if flagged_entries:
                    st.caption(f"총 {len(flagged_entries)}개 라인")
                    
                    # 컬러 표시 DataFrame
                    rows = []
                    bg_colors = []
                    for e in flagged_entries:
                        d = e.to_dict()
                        d['_경고'] = ', '.join(e.flag_labels)
                        d['_메모'] = e.메모
                        rows.append(d)
                        bg_colors.append(ST_BG_COLOR.get(e.top_flag, ''))
                    
                    review_df = pd.DataFrame(rows)
                    
                    def highlight_review(row):
                        idx = row.name
                        if idx < len(bg_colors) and bg_colors[idx]:
                            return [f'background-color: {bg_colors[idx]}'] * len(row)
                        return [''] * len(row)
                    
                    styled = review_df.style.apply(highlight_review, axis=1)
                    st.dataframe(styled, use_container_width=True, height=400)

        # 미처리 이벤트
        if converter.unhandled_events:
            with st.expander(f"🔧 미처리 이벤트 ({len(converter.unhandled_events)}건)"):
                unhandled_df = pd.DataFrame([
                    {
                        '날짜': e.날짜, '유형': e.event_type, '원천': e.원천,
                        '금액': e.총액, '종목': e.종목명, '적요': e.적요원본,
                    }
                    for e in converter.unhandled_events
                ])
                st.dataframe(unhandled_df, use_container_width=True)

        # 분개장 미리보기 (전체)
        st.divider()
        st.subheader("📝 분개장 미리보기 (전체 - 컬러 표시)")
        all_entries = book.all_entries()
        preview_count = min(200, len(all_entries))
        
        rows = []
        bg_colors = []
        for e in all_entries[:preview_count]:
            d = e.to_dict()
            rows.append(d)
            bg_colors.append(ST_BG_COLOR.get(e.top_flag, ''))
        
        if rows:
            preview_df = pd.DataFrame(rows)
            
            def highlight_preview(row):
                idx = row.name
                if idx < len(bg_colors) and bg_colors[idx]:
                    return [f'background-color: {bg_colors[idx]}'] * len(row)
                return [''] * len(row)
            
            styled = preview_df.style.apply(highlight_preview, axis=1)
            st.dataframe(styled, use_container_width=True, height=400)
            
            if len(all_entries) > preview_count:
                st.caption(f"... 전체 {len(all_entries):,}행 중 처음 {preview_count}행 표시 (전체는 다운로드)")

        # 다운로드
        st.divider()
        st.subheader("📥 다운로드")
        col_d1, col_d2 = st.columns(2)

        with col_d1:
            excel_bytes = export_journal_to_duzone(
                book, include_legend=True, include_review_sheet=True
            )
            st.download_button(
                label="📥 더존 업로드용 엑셀 다운로드 (색상포함)",
                data=excel_bytes,
                file_name=f"더존_일반전표_{year}년{month:02d}월_{company_code}.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
                help="색상이 입혀진 엑셀 파일. '검토필요' 시트에 플래그 라인만 모아서 보여줍니다."
            )

        with col_d2:
            if 'portfolio' in st.session_state:
                pf_snapshot = st.session_state['portfolio'].snapshot()
                csv = pf_snapshot.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 현재 포트폴리오 다운로드 (다음달용)",
                    data=csv,
                    file_name=f"포트폴리오_{year}년{month:02d}월말.csv",
                    mime='text/csv',
                    use_container_width=True,
                )
