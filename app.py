import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import io
import tempfile
import os

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="centered")

st.title("📊 더존 위하고 전표 변환기")
st.markdown("분개장 파일을 더존 일반전표 업로드 형식으로 변환합니다.")
st.divider()

# 계정과목 코드 매핑
ACCOUNT_MAP = {
    '보통예금':                     ('10301', '보통예금'),
    '예치금':                       ('10800', '예치금'),
    '선급금':                       ('13200', '선급금'),
    '미수금':                       ('10600', '미수금'),
    '단기매매증권':                  ('11100', '단기매매증권'),
    '단기매매증권평가이익':           ('90600', '단기매매증권평가이익'),
    '단기매매증권평가손실':           ('83700', '단기매매증권평가손실'),
    '단기매매증권처분이익':           ('90700', '단기매매증권처분이익'),
    '단기매매증권처분손실':           ('83800', '단기매매증권처분손실'),
    '이자수익(금융)':                ('91100', '이자수익'),
    '이자수익':                      ('91100', '이자수익'),
    '잡이익':                       ('91900', '잡이익'),
    '잡손실':                       ('89900', '잡손실'),
    '예수금':                       ('25400', '예수금'),
    '미지급금':                     ('25300', '미지급금'),
    '미지급비용':                   ('25200', '미지급비용'),
    '미지급세금':                   ('25600', '미지급세금'),
    '선납세금':                     ('13600', '선납세금'),
    '복리후생비(판)':               ('84700', '복리후생비'),
    '복리후생비':                   ('84700', '복리후생비'),
    '여비교통비(판)':               ('84400', '여비교통비'),
    '여비교통비':                   ('84400', '여비교통비'),
    '접대비(기업업무추진비)(판)':    ('84300', '접대비'),
    '접대비':                       ('84300', '접대비'),
    '기업업무추진비':               ('84300', '접대비'),
    '지급수수료(판)':               ('85200', '지급수수료'),
    '지급수수료':                   ('85200', '지급수수료'),
    '주식거래수수료(판)':           ('85200', '지급수수료'),
    '주식거래수수료':               ('85200', '지급수수료'),
    '세금과공과금(판)':             ('82700', '세금과공과금'),
    '세금과공과금':                 ('82700', '세금과공과금'),
    '보험료(판)':                   ('85100', '보험료'),
    '보험료':                       ('85100', '보험료'),
    '임원급여(판)':                 ('81100', '임원급여'),
    '임원급여':                     ('81100', '임원급여'),
    '직원급여(판)':                 ('81200', '급여'),
    '급여':                         ('81200', '급여'),
    '통신비':                       ('84600', '통신비'),
    '통신비(판)':                   ('84600', '통신비'),
    '임차료':                       ('85000', '임차료'),
    '임차료(판)':                   ('85000', '임차료'),
    '감가상각비':                   ('86200', '감가상각비'),
    '광고선전비':                   ('84100', '광고선전비'),
    '소모품비':                     ('85400', '소모품비'),
    '도서인쇄비':                   ('85500', '도서인쇄비'),
    '법인세비용':                   ('99900', '법인세비용'),
    '현금':                         ('10100', '현금'),
    '법인카드':                     ('10400', '보통예금'),
}

def clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except:
        pass
    s = str(v).strip()
    return None if s in ('', 'nan', 'NaN', 'None') else s

def to_int(v):
    try:
        if v is None or pd.isna(v):
            return 0
        return int(float(str(v).replace(',', '')))
    except:
        return 0

def get_account(name):
    if name is None:
        return ('', name or '')
    name_clean = name.strip()
    if name_clean in ACCOUNT_MAP:
        return ACCOUNT_MAP[name_clean]
    # 부분 매칭
    for key, val in ACCOUNT_MAP.items():
        if key in name_clean or name_clean in key:
            return val
    return ('', name_clean)

def parse_and_convert(uploaded_journal):
    """분개장 파일을 파싱하여 더존 업로드 형식으로 변환"""

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as tmp:
        tmp.write(uploaded_journal.read())
        tmp_path = tmp.name

    try:
        df_raw = pd.read_excel(tmp_path, engine='xlrd', header=None)
    except Exception as e:
        try:
            df_raw = pd.read_excel(tmp_path, engine='openpyxl', header=None)
        except:
            raise Exception(f"파일을 읽을 수 없습니다: {e}")
    finally:
        os.unlink(tmp_path)

    data = df_raw.values

    # 전표 파싱
    vouchers = []
    current_date_str = None
    current_seq_str = None
    current_voucher_lines = []

    for row in data:
        col0 = clean(row[0]) if len(row) > 0 else None
        col1 = clean(row[1]) if len(row) > 1 else None
        col2 = to_int(row[2]) if len(row) > 2 else 0
        col3 = clean(row[3]) if len(row) > 3 else None
        col4 = clean(row[4]) if len(row) > 4 else None
        col5 = to_int(row[5]) if len(row) > 5 else 0

        # 헤더/합계/제목 스킵
        if col0 in ('구     분', '월/일', '합 계'):
            continue
        if col3 and '분   개   장' in str(col3):
            continue
        if col3 and '분개장' in str(col3):
            continue

        # 날짜가 있는 행 = 새 전표 시작
        if col0 and '/' in str(col0):
            if current_voucher_lines:
                vouchers.append((current_date_str, current_seq_str, current_voucher_lines))
            current_date_str = col0
            current_seq_str = col1
            current_voucher_lines = [(col2, col3, col4, col5)]
        elif col0 is None and current_date_str is not None:
            current_voucher_lines.append((col2, col3, col4, col5))

    if current_voucher_lines:
        vouchers.append((current_date_str, current_seq_str, current_voucher_lines))

    # 더존 업로드 형식으로 변환
    upload_rows = []
    unmapped = set()

    for date_str, seq_str, lines in vouchers:
        try:
            parts = str(date_str).split('/')
            month = int(parts[0])
            day = int(parts[1])
        except:
            continue

        debit_entries = []
        credit_entries = []
        memo = ''

        for amt_dr, acct_dr, acct_cr, amt_cr in lines:
            has_dr = amt_dr > 0
            has_cr = amt_cr > 0

            if has_dr and acct_dr:
                debit_entries.append({'amt': amt_dr, 'acct': acct_dr, 'memo': ''})
            if has_cr and acct_cr:
                credit_entries.append({'amt': amt_cr, 'acct': acct_cr, 'memo': ''})

            # 적요 처리 (금액 없고 텍스트만 있는 경우)
            if not has_dr and not has_cr:
                if acct_dr and not memo:
                    memo = acct_dr
                if acct_cr and not memo:
                    memo = acct_cr

        # 적요를 각 항목에 적용
        for e in debit_entries:
            if not e['memo']:
                e['memo'] = memo
        for e in credit_entries:
            if not e['memo']:
                e['memo'] = memo

        # 더존 형식 행 생성
        for e in debit_entries:
            code, name = get_account(e['acct'])
            if not code:
                unmapped.add(e['acct'])
            upload_rows.append([
                month, day, '차변', code, name,
                '', '', e['memo'], e['amt'], ''
            ])

        for e in credit_entries:
            code, name = get_account(e['acct'])
            if not code:
                unmapped.add(e['acct'])
            upload_rows.append([
                month, day, '대변', code, name,
                '', '', e['memo'], '', e['amt']
            ])

    return upload_rows, len(vouchers), unmapped

def create_output_excel(upload_rows, template_file=None):
    """더존 업로드용 엑셀 파일 생성"""

    if template_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(template_file.read())
            tmp_path = tmp.name
        wb = load_workbook(tmp_path)
        ws = wb.active
        os.unlink(tmp_path)

        # 기존 데이터 행 삭제 (헤더 제외)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.value = None

        for i, row_data in enumerate(upload_rows, start=2):
            for j, val in enumerate(row_data, start=1):
                if val != '':
                    ws.cell(row=i, column=j, value=val)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "전표데이터"
        headers = ['월', '일', '구분', '계정과목코드', '계정과목명',
                   '거래처코드', '거래처명', '적요명', '차변(출금)', '대변(입금)']
        ws.append(headers)
        for row_data in upload_rows:
            ws.append([v if v != '' else None for v in row_data])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ── UI ──────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("① 분개장 파일")
    journal_file = st.file_uploader(
        "분개장 (.xls / .xlsx)",
        type=['xls', 'xlsx'],
        key="journal"
    )

with col2:
    st.subheader("② 더존 템플릿 (선택)")
    template_file = st.file_uploader(
        "템플릿 파일 (.xlsx)",
        type=['xlsx'],
        key="template",
        help="없으면 기본 형식으로 생성됩니다"
    )

st.divider()

if journal_file is not None:
    if st.button("🔄 변환 시작", type="primary", use_container_width=True):
        with st.spinner("변환 중..."):
            try:
                upload_rows, voucher_count, unmapped = parse_and_convert(journal_file)

                if not upload_rows:
                    st.error("변환된 데이터가 없습니다. 파일 형식을 확인해 주세요.")
                else:
                    output_excel = create_output_excel(upload_rows, template_file)

                    # 합계 계산
                    total_dr = sum(r[8] for r in upload_rows if r[8] != '')
                    total_cr = sum(r[9] for r in upload_rows if r[9] != '')

                    st.success("✅ 변환 완료!")

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("전표 수", f"{voucher_count:,}건")
                    col_b.metric("총 행 수", f"{len(upload_rows):,}행")
                    col_c.metric("차대변 일치", "✅" if total_dr == total_cr else "❌ 불일치")

                    st.info(f"차변 합계: {total_dr:,.0f}원  |  대변 합계: {total_cr:,.0f}원")

                    if unmapped:
                        st.warning(
                            f"⚠️ 계정과목 코드 미매핑 항목 ({len(unmapped)}개): "
                            + ", ".join(sorted(unmapped))
                            + "\n\n코드란이 공백으로 처리되었습니다. 더존에서 직접 수정하거나 매핑 추가를 요청해 주세요."
                        )

                    # 파일명에서 월 추출
                    fname = journal_file.name.replace('.xls', '').replace('.xlsx', '')
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
                st.info("파일 형식이 올바른지 확인해 주세요.")
else:
    st.info("분개장 파일을 업로드하면 변환 버튼이 활성화됩니다.")

st.divider()
st.caption("계정과목 코드 추가/수정이 필요하면 Claude에게 요청해 주세요.")
