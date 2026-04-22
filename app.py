import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import io
import tempfile
import os

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="centered")
st.title("📊 더존 위하고 전표 변환기")
st.markdown("거래내역 파일을 더존 일반전표 업로드 형식으로 변환합니다.")
st.divider()

TEMPLATE_PATH = "더존위하고_일반전표입력_엑셀_업로드_Template.xlsx"

# 계정과목 코드 매핑
ACCOUNT_MAP = {
    '보통예금':         ('10301', '보통예금'),
    '예치금':           ('10800', '예치금'),
    '선급금':           ('13200', '선급금'),
    '미수금':           ('10600', '미수금'),
    '미지급금':         ('25300', '미지급금'),
    '미지급비용':       ('25200', '미지급비용'),
    '미지급세금':       ('25600', '미지급세금'),
    '예수금':           ('25400', '예수금'),
    '선납세금':         ('13600', '선납세금'),
    '단기매매증권':     ('11100', '단기매매증권'),
    '이자수익':         ('91100', '이자수익'),
    '잡이익':           ('91900', '잡이익'),
    '복리후생비':       ('84700', '복리후생비'),
    '접대비':           ('84300', '접대비'),
    '여비교통비':       ('84400', '여비교통비'),
    '통신비':           ('84600', '통신비'),
    '임차료':           ('85000', '임차료'),
    '지급수수료':       ('85200', '지급수수료'),
    '세금과공과금':     ('82700', '세금과공과금'),
    '보험료':           ('85100', '보험료'),
    '급여':             ('81200', '급여'),
    '임원급여':         ('81100', '임원급여'),
    '법인세비용':       ('99900', '법인세비용'),
    '광고선전비':       ('84100', '광고선전비'),
    '소모품비':         ('85400', '소모품비'),
}

# 거래내용 키워드 → 계정과목 자동 매핑
KEYWORD_RULES = [
    (['출고수수료'],                                    ('85200', '지급수수료')),
    (['KT', '통신비', 'KT유선', 'SKT', 'LGU'],         ('84600', '통신비')),
    (['관리비'],                                        ('85000', '임차료')),
    (['임차료', '임대료'],                              ('85000', '임차료')),
    (['복리후생비', '식대', '부식비', '중식', '석식'],  ('84700', '복리후생비')),
    (['접대비', '경조비', '접 대'],                     ('84300', '접대비')),
    (['출장비', '여비', '교통비', '주차'],              ('84400', '여비교통비')),
    (['회계감사', '세무조정', '수수료'],                ('85200', '지급수수료')),
    (['급여', '임금'],                                  ('81200', '급여')),
    (['임원급여'],                                      ('81100', '임원급여')),
    (['산재보험', '고용보험', '합산보험', '건강보험'],  ('82700', '세금과공과금')),
    (['이자수익', '이자'],                              ('91100', '이자수익')),
    (['연말정산', '소득세 환급', '지방소득세 환급'],    ('25400', '예수금')),
    (['법인세'],                                        ('13600', '선납세금')),
    (['국세', '지방세'],                               ('82700', '세금과공과금')),
    (['납입금', '수익금출금', '수익금입금'],            ('10800', '예치금')),
]

def get_account_from_keyword(text1, text2):
    combined = ' '.join(filter(None, [str(text1 or ''), str(text2 or '')])).strip()
    if not combined or combined == 'nan':
        return None
    # 거래내용2에 계정과목명이 직접 명시된 경우 우선 처리
    for key, val in ACCOUNT_MAP.items():
        t2 = str(text2 or '')
        if key in t2:
            return val
    # 키워드 규칙 적용
    for keywords, account in KEYWORD_RULES:
        for kw in keywords:
            if kw in combined:
                return account
    return None

def parse_date(date_val):
    try:
        if isinstance(date_val, str):
            parts = date_val.split('-')
            return int(parts[1]), int(parts[2][:2])
        else:
            return date_val.month, date_val.day
    except:
        return None, None

def process_ibk(df):
    rows = []
    unmapped = []

    # 헤더 찾기
    header_row = None
    for i, row in df.iterrows():
        if '거래일시' in str(row.values):
            header_row = i
            break
    if header_row is None:
        return rows, unmapped

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row+1:].reset_index(drop=True)

    for _, row in df.iterrows():
        try:
            date_val = row.get('거래일시')
            if pd.isna(date_val) or '합계' in str(date_val):
                continue

            month, day = parse_date(str(date_val))
            if not month:
                continue

            try:
                out_amt = int(float(str(row.get('출금', 0) or 0)))
                in_amt  = int(float(str(row.get('입금', 0) or 0)))
            except:
                continue

            text1 = str(row.get('거래내용1') or '').strip()
            text2 = str(row.get('거래내용2') or '').strip()
            memo  = text2 if text2 and text2 != 'nan' else text1

            account = get_account_from_keyword(text1, text2)

            if out_amt > 0:
                if account:
                    # 차변: 해당 계정 / 대변: 보통예금
                    rows.append([month, day, '차변', account[0], account[1], '', '', memo, out_amt, ''])
                    rows.append([month, day, '대변', '10301', '보통예금', '', '', memo, '', out_amt])
                else:
                    unmapped.append({'날짜': f'{month}/{day}', '출금': out_amt, '입금': 0, '거래내용1': text1, '거래내용2': text2})

            elif in_amt > 0:
                if account:
                    # 차변: 보통예금 / 대변: 해당 계정
                    rows.append([month, day, '차변', '10301', '보통예금', '', '', memo, in_amt, ''])
                    rows.append([month, day, '대변', account[0], account[1], '', '', memo, '', in_amt])
                else:
                    unmapped.append({'날짜': f'{month}/{day}', '출금': 0, '입금': in_amt, '거래내용1': text1, '거래내용2': text2})
        except:
            continue

    return rows, unmapped

def process_card(df):
    rows = []
    unmapped = []

    # 헤더 찾기
    header_row = None
    for i, row in df.iterrows():
        if '거래내용1' in str(row.values):
            header_row = i
            break
    if header_row is None:
        return rows, unmapped

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row+1:].reset_index(drop=True)

    for _, row in df.iterrows():
        try:
            date_val = row.get('결제일') or row.get('거래일')
            if pd.isna(date_val):
                continue

            month, day = parse_date(str(date_val))
            if not month:
                continue

            try:
                amt = int(float(str(row.get('승인금액', 0) or 0)))
            except:
                continue
            if amt <= 0:
                continue

            text1 = str(row.get('거래내용1') or '').strip()
            text2 = str(row.get('거래내용2') or '').strip()
            memo  = text1

            account = get_account_from_keyword(text1, text2)

            if account:
                # 차변: 비용계정 / 대변: 미지급금
                rows.append([month, day, '차변', account[0], account[1], '', '', memo, amt, ''])
                rows.append([month, day, '대변', '25300', '미지급금', '', '', memo, '', amt])
            else:
                unmapped.append({'날짜': f'{month}/{day}', '출금': amt, '입금': 0, '거래내용1': text1, '거래내용2': text2})
        except:
            continue

    return rows, unmapped

def create_output_excel(all_rows):
    if os.path.exists(TEMPLATE_PATH):
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.value = None
        for i, row_data in enumerate(all_rows, start=2):
            for j, val in enumerate(row_data, start=1):
                if val != '':
                    ws.cell(row=i, column=j, value=val)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        headers = ['월','일','구분','계정과목코드','계정과목명','거래처코드','거래처명','적요명','차변(출금)','대변(입금)']
        ws.append(headers)
        for row_data in all_rows:
            ws.append([v if v != '' else None for v in row_data])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ── UI ──────────────────────────────────────────────────────────────────────

st.subheader("거래내역 파일 업로드")
uploaded_file = st.file_uploader("거래내역 (.xlsx)", type=['xlsx'], key="txn")

st.divider()

if uploaded_file:
    if st.button("🔄 변환 시작", type="primary", use_container_width=True):
        with st.spinner("변환 중..."):
            try:
                xl = pd.ExcelFile(uploaded_file)
                sheet_names = xl.sheet_names

                all_rows = []
                all_unmapped = []

                # IBK 기업은행 처리
                ibk_sheets = [s for s in sheet_names if 'IBK' in s or '기업은행' in s]
                for s in ibk_sheets:
                    df = pd.read_excel(xl, sheet_name=s, header=None)
                    rows, unmapped = process_ibk(df)
                    all_rows.extend(rows)
                    all_unmapped.extend([{**u, '출처': s} for u in unmapped])

                # 비씨카드 처리
                card_sheets = [s for s in sheet_names if '비씨' in s or '카드' in s or '세부' in s]
                for s in card_sheets:
                    df = pd.read_excel(xl, sheet_name=s, header=None)
                    rows, unmapped = process_card(df)
                    all_rows.extend(rows)
                    all_unmapped.extend([{**u, '출처': s} for u in unmapped])

                if not all_rows:
                    st.error("변환된 데이터가 없습니다. 파일 형식을 확인해 주세요.")
                else:
                    output_excel = create_output_excel(all_rows)

                    total_dr = sum(r[8] for r in all_rows if r[8] != '')
                    total_cr = sum(r[9] for r in all_rows if r[9] != '')

                    st.success("✅ 변환 완료!")

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("전표 행 수", f"{len(all_rows):,}행")
                    col_b.metric("미매핑 건수", f"{len(all_unmapped):,}건")
                    col_c.metric("차대변 일치", "✅" if total_dr == total_cr else "❌ 불일치")

                    st.info(f"차변 합계: {total_dr:,.0f}원  |  대변 합계: {total_cr:,.0f}원")

                    if all_unmapped:
                        st.warning(f"⚠️ 아래 {len(all_unmapped)}건은 계정과목을 자동으로 판단하지 못했습니다. 수동으로 추가해 주세요.")
                        st.dataframe(pd.DataFrame(all_unmapped), use_container_width=True)

                    fname = uploaded_file.name.replace('.xlsx', '')
                    out_name = f"더존업로드_{fname}.xlsx"

                    st.download_button(
                        label="📥 변환 파일 다운로드",
                        data=output_excel,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )

            except Exception as e:
                st.error(f"오류 발생: {str(e)}")
else:
    st.info("거래내역 파일을 업로드하면 변환 버튼이 활성화됩니다.")

st.divider()
st.caption("계정과목 추가/수정이 필요하면 Claude에게 요청해 주세요.")
