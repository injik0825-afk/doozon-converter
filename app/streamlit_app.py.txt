import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import io, tempfile, os, re

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")
st.markdown("거래내역 파일을 더존 일반전표 업로드 형식으로 변환합니다.")

TEMPLATE_PATH = "더존위하고_일반전표입력_엑셀_업로드_Template.xlsx"

# ── 종목 DB ───────────────────────────────────────────────────────────────────
# 회사별로 분리. 자동 감지 후 STOCK_DB 변수에 할당
RIVERSIDE_STOCK_DB = {
    '케이뱅크':          '주식#코스피#케이뱅크',
    '에스팀':            '주식#코스닥#에스팀',
    '액스비스':          '주식#코스닥#액스비스',
    '티엠씨':            '주식#코스닥#티엠씨',
    'HC홈센타':          '주식#코스닥#HC홈센타',
    '제이엘케이':        '주식#코스닥#제이엘케이',
    '카나프테라퓨틱스':  '주식#코스닥#카나프테라퓨틱스',
    '아이엠바이오로직스':'주식#코스닥#아이엠바이오로직스',
    '메쥬':              '주식#코스닥#메쥬',
    '한패스':            '주식#코스닥#한패스',
    '리센스메디컬':      '주식#코스닥#리센스메디컬',
    '엔에이치스팩33호':  '주식#코스닥#엔에이치스팩33호',
    '신한제17호스팩':    '주식#코스닥#신한제17호스팩',
    '교보20호스팩':      '주식#코스닥#교보20호스팩',
    '인벤테라':          '주식#코스닥#인벤테라',
    '에스엔시스':        '주식#코스닥#에스엔시스',
    '두산퓨얼셀10-2':    '채권#상장#두산퓨얼셀10-2',
    '두산퓨얼셀9-2':     '채권#상장#두산퓨얼셀9-2',
    '두산에너빌리티79-2':'채권#상장#두산에너빌리티79-2',
    '한진117-2':         '채권#상장#한진117-2',
    '한국자산신탁8-1':   '채권#상장#한국자산신탁8-1',
}

DURI_STOCK_DB = {
    'SK리츠':              '주식#코스피#SK리츠',
    'TIGER200':            '주식#코스피#TIGER200',
    '교보스팩20호':        '주식#코스닥#교보스팩20호',
    '노타':                '주식#코스닥#노타',
    '더핑크퐁':            '주식#코스닥#더핑크퐁',
    '두산로보틱스':        '주식#코스피#두산로보틱스',
    '리센스메디컬':        '주식#코스닥#리센스메디컬',
    '맥쿼리인프라':        '주식#코스피#맥쿼리인프라',
    '메쥬':                '주식#코스닥#메쥬',
    '명인제약':            '주식#코스피#명인제약',
    '바이오프로테크':      '주식#코넥스#바이오프로테크',
    '삼성전자':            '주식#코스피#삼성전자',
    '삼양컴텍':            '주식#코스닥#삼양컴텍',
    '세미파이브':          '주식#코스닥#세미파이브',
    '신한스팩17호':        '주식#코스닥#신한스팩17호',
    '씨엠티엑스':          '주식#코스닥#씨엠티엑스',
    '아로마티카':          '주식#코스닥#아로마티카',
    '아이엠바이오로직스':  '주식#코스닥#아이엠바이오로직스',
    '아크릴':              '주식#코스닥#아크릴',
    '알지노믹스':          '주식#코스닥#알지노믹스',
    '액스비스':            '주식#코스닥#액스비스',
    '에스팀':              '주식#코스닥#에스팀',
    '에임드바이오':        '주식#코스닥#에임드바이오',
    '엔에이치스팩33호':    '주식#코스닥#엔에이치스팩33호',
    '인벤테라':            '주식#코스닥#인벤테라',
    '카나프테라퓨틱스':    '주식#코스닥#카나프테라퓨틱스',
    '케이뱅크':            '주식#코스피#케이뱅크',
    '코람코더원리츠':      '주식#코스피#코람코더원리츠',
    '큐리오시스':          '주식#코스닥#큐리오시스',
    '티엠씨':              '주식#코스닥#티엠씨',
    '한패스':              '주식#코스닥#한패스',
    '국민주택1종채권25-06': '채권#국민주택#국민주택1종채권25-06',
    '대한항공102-2':       '채권#회사채#대한항공102-2',
    '두산310-2':           '채권#회사채#두산310-2',
    '두산311-2':           '채권#회사채#두산311-2',
    '두산에너빌리티78-2':  '채권#회사채#두산에너빌리티78-2',
    '두산에너빌리티79-2':  '채권#회사채#두산에너빌리티79-2',
    '이마트24신종자본증권37':'채권#회사채#이마트24신종자본증권37',
    '한진115-2':           '채권#회사채#한진115-2',
    '한진117-2':           '채권#회사채#한진117-2',
    '한진123-2':           '채권#회사채#한진123-2',
    '한진칼13':            '채권#회사채#한진칼13',
}

# 자동 감지 후 할당. 기본은 리버사이드.
STOCK_DB = RIVERSIDE_STOCK_DB

ABBREV_LIST = [
    ('아이엠바이오', '아이엠바이오로직스'), ('카나프', '카나프테라퓨틱스'),
    ('두산퓨얼셀', '두산퓨얼셀10-2'),      ('인벤테', '인벤테라'),
    ('리센스', '리센스메디컬'),             ('신한17', '신한제17호스팩'),
    ('교보20', '교보20호스팩'),             ('NH33', '엔에이치스팩33호'),
    ('제이엘케이', '제이엘케이'),           ('아이엠', '아이엠바이오로직스'),
    ('한패스', '한패스'), ('케이뱅크', '케이뱅크'), ('액스비스', '액스비스'),
    ('티엠씨', '티엠씨'), ('에스팀', '에스팀'), ('메쥬', '메쥬'),
    ('HC홈센타', 'HC홈센타'), ('에스엔시스', '에스엔시스'),
]

# 두리 거래내역 종목 별칭 → 표준 종목명 (긴 키워드부터 매칭)
DURI_ABBREV = [
    ('카나프테라퓨틱스', '카나프테라퓨틱스'),
    ('카나프테라',       '카나프테라퓨틱스'),
    ('카나프',           '카나프테라퓨틱스'),
    ('아이엠바이오로직스','아이엠바이오로직스'),
    ('아이엠바이오',     '아이엠바이오로직스'),
    ('아이엠',           '아이엠바이오로직스'),
    ('엔에이치스팩33호', '엔에이치스팩33호'),
    ('NH스팩33',         '엔에이치스팩33호'),
    ('NH33',             '엔에이치스팩33호'),
    ('신한스팩17호',     '신한스팩17호'),
    ('신한스팩17',       '신한스팩17호'),
    ('신한17',           '신한스팩17호'),
    ('교보스팩20호',     '교보스팩20호'),
    ('교보스팩20',       '교보스팩20호'),
    ('교보스팩',         '교보스팩20호'),
    ('교보20호',         '교보스팩20호'),
    ('교보20',           '교보스팩20호'),
    ('인벤테라',         '인벤테라'),
    ('리센스메디컬',     '리센스메디컬'),
    ('리센스',           '리센스메디컬'),
    ('큐리오시스',       '큐리오시스'),
    ('한패스',           '한패스'),
    ('메쥬',             '메쥬'),
    ('티엠씨',           '티엠씨'),
    ('액스비스',         '액스비스'),
    ('에스팀',           '에스팀'),
    ('에임드바이오',     '에임드바이오'),
    ('알지노믹스',       '알지노믹스'),
    ('아크릴',           '아크릴'),
    ('세미파이브',       '세미파이브'),
    ('아로마티카',       '아로마티카'),
    ('씨엠티엑스',       '씨엠티엑스'),
    ('더핑크퐁',         '더핑크퐁'),
    ('노타',             '노타'),
    ('삼양컴텍',         '삼양컴텍'),
    ('바이오프로테크',   '바이오프로테크'),
    ('두산로보틱스',     '두산로보틱스'),
    ('맥쿼리인프라',     '맥쿼리인프라'),
    ('명인제약',         '명인제약'),
    ('케이뱅크',         '케이뱅크'),
    ('SK리츠',           'SK리츠'),
    ('TIGER200',         'TIGER200'),
    ('코람코더원리츠',   '코람코더원리츠'),
    ('코람코',           '코람코더원리츠'),
    ('삼성전자',         '삼성전자'),
]


def extract_duri_stock(text):
    """두리 거래내용에서 종목명 추출 (긴 별칭부터 매칭)"""
    text = str(text).strip()
    for abbrev, full in DURI_ABBREV:
        if abbrev in text:
            return full
    return ''

# ── 계정과목 코드 ─────────────────────────────────────────────────────────────
# 회사별로 분리. 자동 감지 후 AC 변수에 할당
# 리버사이드: 5자리 코드 (10301, 11100 등)
# 두리: 3자리 코드 (103, 107 등)
# 계정명은 두 회사 모두 (판) 접미사 + 이자수익(금융) 등 동일
RIVERSIDE_AC = {
    '보통예금':     ('10301', '보통예금'),
    '예치금':       ('10800', '예치금'),
    '선급금':       ('13200', '선급금'),
    '미수금':       ('10600', '미수금'),
    '단매증':       ('11100', '단기매매증권'),
    '단매이익':     ('90600', '단기매매증권평가이익'),
    '단매손실':     ('83700', '단기매매증권평가손실'),
    '처분이익':     ('90700', '단기매매증권처분이익'),
    '처분손실':     ('83800', '단기매매증권처분손실'),
    '이자수익':     ('91100', '이자수익(금융)'),
    '분배금수익':   ('91200', '분배금수익'),
    '잡이익':       ('91900', '잡이익'),
    '예수금':       ('25400', '예수금'),
    '미지급금':     ('25300', '미지급금'),
    '미지급비용':   ('25200', '미지급비용'),
    '미지급세금':   ('25600', '미지급세금'),
    '선납세금':     ('13600', '선납세금'),
    '복리후생비':   ('84700', '복리후생비(판)'),
    '접대비':       ('84300', '접대비(기업업무추진비)(판)'),
    '여비교통비':   ('84400', '여비교통비(판)'),
    '지급수수료':   ('85200', '지급수수료(판)'),
    '주식거래수수료': ('85201', '주식거래수수료(판)'),
    '세금과공과금': ('82700', '세금과공과금(판)'),
    '보험료':       ('85100', '보험료(판)'),
    '급여':         ('81200', '직원급여(판)'),
    '임원급여':     ('81100', '임원급여(판)'),
}

DURI_AC = {
    '보통예금':     ('103', '보통예금'),
    '단매증':       ('107', '단기매매증권'),
    '예치금':       ('125', '예치금'),
    '선급금':       ('131', '선급금'),
    '선납세금':     ('136', '선납세금'),
    '미지급금':     ('253', '미지급금'),
    '예수금':       ('254', '예수금'),
    '미지급세금':   ('261', '미지급세금'),
    '퇴직급여충당부채': ('295', '퇴직급여충당부채'),
    '처분이익':     ('412', '단기매매증권처분이익'),
    '단매이익':     ('413', '단기매매증권평가이익'),
    '투자일임수수료입':('416', '투자일임수수료'),
    '분배금수익':   ('419', '분배금수익'),
    '이자수익':     ('420', '이자수익(금융)'),
    '처분손실':     ('457', '단기매매증권처분손실'),
    '단매손실':     ('458', '단기매매증권평가손실'),
    '임원급여':     ('801', '임원급여(판)'),
    '급여':         ('802', '직원급여(판)'),
    '상여금임원':   ('803', '상여금(임원)(판)'),
    '상여금직원':   ('804', '상여금(직원)(판)'),
    '임원퇴직DC':   ('806', '임원퇴직급여(DC)(판)'),
    '직원퇴직DC':   ('808', '직원퇴직급여(DC)(판)'),
    '직원퇴직DB':   ('809', '직원퇴직급여(DB)(판)'),
    '복리후생비':   ('811', '복리후생비(판)'),
    '여비교통비':   ('812', '여비교통비(판)'),
    '접대비':       ('813', '접대비(기업업무추진비)(판)'),
    '통신비':       ('814', '통신비(판)'),
    '세금과공과금': ('817', '세금과공과금(판)'),
    '감가상각비':   ('818', '감가상각비(판)'),
    '지급임차료':   ('819', '지급임차료(판)'),
    '보험료':       ('821', '보험료(판)'),
    '차량유지비':   ('822', '차량유지비(판)'),
    '교육훈련비':   ('825', '교육훈련비(판)'),
    '도서인쇄비':   ('826', '도서인쇄비(판)'),
    '주식거래수수료': ('828', '주식거래수수료(판)'),
    '소모품비':     ('830', '소모품비(판)'),
    '지급수수료':   ('831', '지급수수료(판)'),
    '건물관리비':   ('837', '건물관리비(판)'),
    '잡이익':       ('930', '잡이익'),
    '잡손실':       ('960', '잡손실'),
    '법인세등':     ('998', '법인세등'),
    # 두리에는 미수금/미지급비용 없음 → 미지급금으로 매핑
    '미수금':       ('253', '미지급금'),
    '미지급비용':   ('253', '미지급금'),
}

# 자동 감지 후 할당. 기본은 리버사이드.
AC = RIVERSIDE_AC

# ── 색상 정의 ─────────────────────────────────────────────────────────────────
FILL_RED    = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')
FILL_ORANGE = PatternFill(start_color='FFB347', end_color='FFB347', fill_type='solid')
FILL_YELLOW = PatternFill(start_color='FFF59D', end_color='FFF59D', fill_type='solid')


# ── 회사 자동 감지 ────────────────────────────────────────────────────────────
def detect_company(sheet_names):
    """
    시트명 패턴으로 회사 자동 감지.
    - 두리: 시트명이 '키움(...)', '메리츠(...)', '한국(80...)', '교보(1...)' 등 계좌번호 형식
            또는 'IBK기업은행 내역', '카드이용내역', '납입내역', '우리카드내역' 등
    - 리버사이드: '한국투자증권', '교보증권', 'IBK 기업은행'(공백), '비씨', 'iM증권' 등
    """
    names = [str(s) for s in sheet_names]
    full = ' | '.join(names)

    # 두리 강한 마커
    duri_markers = ['키움(', '메리츠(', 'IBK기업은행 내역', '카드이용내역',
                    '납입내역', '우리카드내역', '하나카드내역', '은행거래내역', 'IBK카드내역']
    for m in duri_markers:
        if m in full:
            return 'duri'
    # 시트명에 '한국(80', '교보(1' 같은 계좌번호 형식
    for n in names:
        if re.match(r'^(한국|교보|키움|메리츠|KB)\(\d', n):
            return 'duri'

    # 리버사이드 마커
    riverside_markers = ['IBK 기업은행', '한국투자증권', '교보증권', '비씨카드',
                         '비씨', '세부 내역', 'iM증권', '미래에셋증권', '신한투자증권',
                         '삼성증권', 'NH투자증권', 'KB증권']
    for m in riverside_markers:
        if m in full:
            return 'riverside'

    return 'unknown'

# ── 유틸 ─────────────────────────────────────────────────────────────────────
def clean(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except: pass
    return str(v).strip()

def to_int(v):
    try: return int(float(str(v).replace(',', '').replace(' ', '')))
    except: return 0

def parse_date(v):
    s = clean(v)
    if not s: return None, None
    try:
        d = pd.to_datetime(s[:10])
        return d.month, d.day
    except: pass
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except: return None, None

def get_stock(name):
    return STOCK_DB.get(name, '')

def extract_stock_from_text(text):
    for abbrev, full in ABBREV_LIST:
        if abbrev in text:
            return full
    return None

def row(month, day, div, acct_key, cp, memo, debit, credit, flag=''):
    code, name = AC[acct_key]
    return [month, day, div, code, name, '', cp, memo,
            debit if debit else '', credit if credit else '', flag]

# ── IBK 기업은행 처리 ─────────────────────────────────────────────────────────
def process_ibk(df, stock_by_date):
    rows, unmapped = [], []
    hdr = None
    for i, r in df.iterrows():
        if '거래일시' in str(r.values):
            hdr = i; break
    if hdr is None: return rows, unmapped

    df.columns = df.iloc[hdr]
    df = df.iloc[hdr+1:].reset_index(drop=True)

    for _, r in df.iterrows():
        date_val = r.get('거래일시', '')
        if not clean(date_val) or '합계' in clean(date_val): continue
        month, day = parse_date(date_val)
        if not month: continue

        out_amt = to_int(r.get('출금', 0))
        in_amt  = to_int(r.get('입금', 0))
        if out_amt == 0 and in_amt == 0: continue

        t1       = clean(r.get('거래내용1', ''))
        t2       = clean(r.get('거래내용2', ''))
        partner  = clean(r.get('상대계좌예금주명', ''))
        combined = t1 + ' ' + t2

        result = classify_ibk(month, day, out_amt, in_amt, t1, t2,
                               combined, partner, stock_by_date)
        if result:
            rows.extend(result)
        elif result == []:
            # 의도적 무시 (예: 교보 시트에서 이미 처리되는 거래)
            pass
        else:
            unmapped.append({'날짜': f'{month}/{day}', '출금': out_amt,
                             '입금': in_amt, '거래내용1': t1, '거래내용2': t2})
    return rows, unmapped


def classify_ibk(month, day, out_amt, in_amt, t1, t2, combined, partner, stock_by_date):
    rows = []

    if out_amt > 0:
        if '출고수수료' in t1:
            stocks_today = stock_by_date.get((month, day), [])
            sl = get_stock(stocks_today[0]) if stocks_today else ''
            memo = f'{sl} 출고수수료' if sl else '출고수수료'
            rows += [row(month, day, '차변', '지급수수료', sl, memo, out_amt, 0),
                     row(month, day, '대변', '보통예금',   '',  memo, 0, out_amt)]
            return rows

        if any(k in combined for k in ['복리후생비', '식대', '부식비', '회식비']):
            if '카드' not in combined and '미지급금' not in combined:
                memo = t2 if t2 else t1
                rows += [row(month, day, '차변', '복리후생비', '', memo, out_amt, 0),
                         row(month, day, '대변', '보통예금',   '', memo, 0, out_amt)]
                return rows

        if any(k in combined for k in ['경조비', '접대비']):
            memo = t2 if t2 else t1
            rows += [row(month, day, '차변', '접대비',   '', memo, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', memo, 0, out_amt)]
            return rows

        if any(k in combined for k in ['출장비', '여비', '숙박']):
            memo = t2 if t2 else t1
            rows += [row(month, day, '차변', '여비교통비', '', memo, out_amt, 0),
                     row(month, day, '대변', '보통예금',   '', memo, 0, out_amt)]
            return rows

        if '유선' in t1 or ('KT' in t1 and '유선' in t1):
            rows += [row(month, day, '차변', '미지급금', '주식회사 케이티', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',               t1, 0, out_amt)]
            return rows

        if '관리비' in combined or '서린' in combined:
            rows += [row(month, day, '차변', '미지급금', '서린빌딩관리사무소', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',                  t1, 0, out_amt)]
            return rows

        if '임차료' in combined:
            rows += [row(month, day, '차변', '미지급금', '서린빌딩관리사무소', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',                  t1, 0, out_amt)]
            return rows

        if any(k in combined for k in ['회계감사', '세무조정']):
            rows += [row(month, day, '차변', '미지급금', '세화회계법인', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',             t1, 0, out_amt)]
            return rows

        if '산재보험' in t1:
            rows += [row(month, day, '차변', '보험료',   '', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', t1, 0, out_amt)]
            return rows

        if '고용보험' in t1:
            emp_rate = 21090 / 93430
            emp_amt  = min(round(out_amt * emp_rate), out_amt)
            boss_amt = out_amt - emp_amt
            rows += [row(month, day, '차변', '예수금',   '근로복지공단', t1, emp_amt,  0),
                     row(month, day, '차변', '보험료',   '',             t1, boss_amt, 0),
                     row(month, day, '대변', '보통예금', '',             t1, 0, out_amt)]
            return rows

        if '합산보험' in t1:
            health_emp  = round(out_amt * 292840 / 579920)
            tax_boss    = round(out_amt * 143540 / 579920)
            pension_emp = out_amt - health_emp - tax_boss
            rows += [row(month, day, '차변', '예수금',       '건강보험공단', t1, health_emp,  0),
                     row(month, day, '차변', '세금과공과금', '국민연금공단', t1, tax_boss,    0),
                     row(month, day, '차변', '예수금',       '국민연금공단', t1, pension_emp, 0),
                     row(month, day, '대변', '보통예금',     '',             t1, 0, out_amt)]
            return rows

        if t1 == '급여' or ('급여' in t1 and '등지급' not in t1):
            rows += [row(month, day, '차변', '미지급비용', '', '급여', out_amt, 0),
                     row(month, day, '대변', '보통예금',   '', '급여', 0, out_amt)]
            return rows

        if '법인세 납부' in t2 or '법인세납부' in t2 or \
           ('국세조회납부' in t1 and out_amt >= 1000000):
            rows += [row(month, day, '차변', '미지급세금', '영등포세무서', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금',   '',             t1, 0, out_amt)]
            return rows

        if '국세조회납부' in t1:
            rows += [row(month, day, '차변', '예수금',   '영등포세무서', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',             t1, 0, out_amt)]
            return rows

        if '지방세납부' in t1:
            rows += [row(month, day, '차변', '예수금',   '영등포구청', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',           t1, 0, out_amt)]
            return rows

        if '비씨카드출금' in t1:
            rows += [row(month, day, '차변', '미지급금', 'BC카드',  '비씨카드출금', out_amt, 0),
                     row(month, day, '대변', '보통예금', '',         '비씨카드출금', 0, out_amt)]
            return rows

        if re.match(r'^I[가-힣0-9A-Za-z]+납입$', t1):
            rows += [row(month, day, '차변', '예수금',   '', t1, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', t1, 0, out_amt)]
            return rows

        if re.match(r'^O[가-힣A-Za-z0-9]+납입$', t1):
            # IBK O 출금 = 위탁업자(금강투자일임 등)에 청약금 송금 → 청약 분개
            stock_name = extract_stock_from_text(t1)
            sl   = get_stock(stock_name) if stock_name else ''
            base = round(out_amt / 1.01)
            fee  = out_amt - base
            memo = f'{sl} 청약납입' if sl else t1
            rows += [row(month, day, '차변', '선급금',       sl, memo, base, 0),
                     row(month, day, '차변', '주식거래수수료', sl, memo, fee,  0),
                     row(month, day, '대변', '보통예금',     '',  memo, 0, out_amt)]
            return rows

        if re.match(r'^고유[가-힣A-Za-z0-9]+납입$', t1):
            stock_name = extract_stock_from_text(t1)
            sl   = get_stock(stock_name) if stock_name else ''
            base = round(out_amt / 1.01)
            fee  = out_amt - base
            memo = f'{sl} 청약납입' if sl else t1
            cp_hantoo  = '한국투자증권(81247132-01)'
            wts_stocks = stock_by_date.get('__wts__', set())
            if stock_name and stock_name in wts_stocks:
                rows += [row(month, day, '차변', '예치금',       cp_hantoo, t1,   out_amt, 0),
                         row(month, day, '대변', '보통예금',     '',         t1,   0, out_amt),
                         row(month, day, '차변', '선급금',       sl,         memo, base, 0),
                         row(month, day, '차변', '주식거래수수료', sl,         memo, fee,  0),
                         row(month, day, '대변', '예치금',       cp_hantoo,  memo, 0, out_amt)]
            else:
                rows += [row(month, day, '차변', '선급금',       sl,  memo, base, 0),
                         row(month, day, '차변', '주식거래수수료', sl,  memo, fee,  0),
                         row(month, day, '대변', '보통예금',     '',  memo, 0, out_amt)]
            return rows

    if in_amt > 0:
        if '리버사이드파트너스' == t1 or \
           ('리버사이드' in t1 and not t2) or \
           ('리버사이드' in partner):
            memo = '계좌이체[한투9969 -> 기업은행]'
            rows += [row(month, day, '차변', '보통예금', '',            memo, in_amt, 0),
                     row(month, day, '대변', '예치금',   '한국투자증권', memo, 0, in_amt)]
            return rows

        if re.match(r'^O[가-힣A-Za-z0-9]+(납입|수익금출금)$', t1):
            # IBK 입금 O하이/O채권 = 교보 시트의 '은행이체출금'에서 이미 분개됨(보통예금/예치금)
            # IBK 쪽에서 또 분개하면 중복이므로 무시
            return []

        if re.match(r'^에[가-힣]_', t1):
            rows += [row(month, day, '차변', '보통예금', '', t1, in_amt, 0),
                     row(month, day, '대변', '예수금',   '', t1, 0, in_amt)]
            return rows

        if '다올투자증권' in t1:
            rows += [row(month, day, '차변', '보통예금', '',            t1, in_amt, 0),
                     row(month, day, '대변', '예수금',   '다올투자증권', t1, 0, in_amt)]
            return rows

        if '영등포세무서' in t1 or '소득세' in t2:
            rows += [row(month, day, '차변', '보통예금', '',            t1, in_amt, 0),
                     row(month, day, '대변', '미수금',   '영등포세무서', t1, 0, in_amt)]
            return rows

        if '영등포지방소득' in t1 or '지방소득세' in t2 or '영등포구청' in t1:
            rows += [row(month, day, '차변', '보통예금', '',           t1, in_amt, 0),
                     row(month, day, '대변', '미수금',   '영등포구청', t1, 0, in_amt)]
            return rows

        if '이자수익' in t2 or '이자' in t2 or '결산' in t1:
            rows += [row(month, day, '차변', '보통예금', '', t1, in_amt, 0),
                     row(month, day, '대변', '이자수익', '', t1, 0, in_amt)]
            return rows

        if re.match(r'^I[가-힣0-9A-Za-z]+납입$', t1):
            rows += [row(month, day, '차변', '보통예금', '', t1, in_amt, 0),
                     row(month, day, '대변', '예수금',   '', t1, 0, in_amt)]
            return rows

    return None


# ── 두리 IBK기업은행 처리 ────────────────────────────────────────────────────
def process_duri_ibk(df, stock_by_date):
    """
    두리 IBK기업은행 내역 파서.
    리버사이드와 차이점:
    - 컬럼명: '거래내용1/2' 단일 컬럼 '거래내용'
    - 청약 패턴: '메쥬－일임'(전각), '카나프테라－고위험', '아이엠바이오일임'(하이픈없음), '웨스트메쥬'
    - 전각 하이픈('－')과 일반 하이픈('-') 모두 사용
    """
    rows, unmapped = [], []
    hdr = None
    for i, r in df.iterrows():
        if '거래일시' in str(r.values):
            hdr = i; break
    if hdr is None: return rows, unmapped

    df.columns = df.iloc[hdr]
    df = df.iloc[hdr+1:].reset_index(drop=True)

    for _, r in df.iterrows():
        date_val = r.get('거래일시', '')
        if not clean(date_val) or '합계' in clean(date_val): continue
        month, day = parse_date(str(date_val))
        if not month: continue

        out_amt = to_int(r.get('출금', 0))
        in_amt  = to_int(r.get('입금', 0))
        if out_amt == 0 and in_amt == 0: continue

        content = clean(r.get('거래내용', ''))
        partner = clean(r.get('상대계좌예금주명', ''))
        memo_field = clean(r.get('메모', ''))

        result = classify_duri_ibk(month, day, out_amt, in_amt, content, partner, memo_field, stock_by_date)
        if result:
            rows.extend(result)
        elif result == []:
            pass  # 의도적 무시
        else:
            unmapped.append({'날짜': f'{month}/{day}', '출금': out_amt,
                             '입금': in_amt, '거래내용': content, '상대': partner})
    return rows, unmapped


def classify_duri_ibk(month, day, out_amt, in_amt, content, partner, memo, stock_by_date):
    """두리 IBK 거래 분류"""
    rows = []

    # ── 출금 처리 ──
    if out_amt > 0:
        # 4대보험 - 국민연금 (직원분=세공 / 회사분=예수금 50/50)
        if '국민연금' in content:
            half = out_amt // 2
            rows += [row(month, day, '차변', '예수금',       '국민연금공단', content, half, 0),
                     row(month, day, '차변', '세금과공과금', '국민연금공단', content, out_amt - half, 0),
                     row(month, day, '대변', '보통예금',     '',             content, 0, out_amt)]
            return rows

        # 고용보험 (직원=예수금 18900/42800, 회사=보험료)
        if '고용보험' in content:
            emp = round(out_amt * 18900 / 42800)
            boss = out_amt - emp
            rows += [row(month, day, '차변', '예수금', '근로복지공단', content, emp, 0),
                     row(month, day, '차변', '보험료', '근로복지공단', content, boss, 0),
                     row(month, day, '대변', '보통예금', '',           content, 0, out_amt)]
            return rows

        # 산재보험 (전액 회사부담)
        if '산재보험' in content:
            rows += [row(month, day, '차변', '보험료',   '근로복지공단', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',             content, 0, out_amt)]
            return rows

        # 건강보험 (절반=예수금 직원분 / 절반=복리후생비 회사부담)
        if '국민건강' in content or '건강보험' in content:
            half = out_amt // 2
            rows += [row(month, day, '차변', '예수금',     '건강보험공단', content, half, 0),
                     row(month, day, '차변', '복리후생비', '건강보험공단', content, out_amt - half, 0),
                     row(month, day, '대변', '보통예금',   '',             content, 0, out_amt)]
            return rows

        # 국세/지방세
        if '국세조회' in content or '국세' in content:
            rows += [row(month, day, '차변', '예수금',   '영등포세무서', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',             content, 0, out_amt)]
            return rows
        if '지방세' in content:
            rows += [row(month, day, '차변', '예수금',   '영등포구청', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',           content, 0, out_amt)]
            return rows

        # 이체수수료
        if '이체수수료' in content:
            rows += [row(month, day, '차변', '지급수수료', '', content, out_amt, 0),
                     row(month, day, '대변', '보통예금',   '', content, 0, out_amt)]
            return rows

        # 퇴직연금수수료
        if '퇴직연금수수료' in content:
            rows += [row(month, day, '차변', '지급수수료', '', content, out_amt, 0),
                     row(month, day, '대변', '보통예금',   '', content, 0, out_amt)]
            return rows

        # 조의금
        if '조의금' in content:
            rows += [row(month, day, '차변', '접대비',   '', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', content, 0, out_amt)]
            return rows

        # 청약 출금 (종목명-유형 패턴)
        # 전각 하이픈("－") 또는 일반 하이픈("-")이거나 하이픈 없이도 가능
        # 유형: 고유, 일임, 고위험, 르네상스(르네상/르네예약 줄임), 채권
        if any(t in content for t in ['－고유','-고유','고유','－일임','-일임','일임',
                                        '－고위험','-고위험','고위험',
                                        '－르네상스','-르네상스','르네상스','르네상','르네예약',
                                        '－채권','-채권']):
            stock_name = extract_duri_stock(content)
            if stock_name:
                rows += [row(month, day, '차변', '예수금',   '수탁운용', content, out_amt, 0),
                         row(month, day, '대변', '보통예금', '',         content, 0, out_amt)]
                return rows

        # 회사 이체 출금 (보통예금 → 증권사로 자금 송금)
        # '두리인베스트먼트(' = 증권사로 보내는 자금이체
        if '두리인베스트먼트' in content:
            rows += [row(month, day, '차변', '예치금',   '', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', content, 0, out_amt)]
            return rows

        # 카드 결제일 대금 출금 (비씨카드출금 / 우리카드결제 / IBK카드 등)
        if any(k in content for k in ['비씨카드출금','우리카드결제','우리카드출금','IBK카드결제','카드결제','카드출금']):
            cp = ''
            if '비씨' in content: cp = 'BC카드'
            elif '우리' in content: cp = '우리카드'
            elif 'IBK' in content: cp = 'IBK카드'
            rows += [row(month, day, '차변', '미지급금', cp, content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', content, 0, out_amt)]
            return rows

        # 통신비 (SKT / KT / 010 핸드폰 번호 패턴)
        if 'SKT' in content or 'KT' in content or re.match(r'^010\d{8}', content):
            rows += [row(month, day, '차변', '통신비',   '', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '', content, 0, out_amt)]
            return rows

        # 관리비 (LG팰리스 빌딩 관리비 - "관리0003-0827" 같은 패턴)
        if '관리' in content and (re.search(r'\d{4}-\d{4}', content) or '관리비' in content):
            rows += [row(month, day, '차변', '미지급금', 'LG팰리스빌딩관리단', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',                  content, 0, out_amt)]
            return rows

        # 임대료
        if '임대료' in content or content == '김민영':
            rows += [row(month, day, '차변', '지급임차료', '', '임대료', out_amt, 0),
                     row(month, day, '대변', '보통예금',   '', '임대료', 0, out_amt)]
            return rows

        # 인명 출금 (박형준/허윤정 = 직원 급여 결제. 전월 미지급 급여 청산)
        # 분개장 패턴: 차변 미지급금(직원급여) / 대변 보통예금
        # 화이트리스트만 매칭 (다른 인명은 미분류로 두어 사용자 검토 유도)
        # - 신민자: 결산성(퇴직금 중간정산), 김민영: 이미 임대료 패턴에서 처리
        if content in ('박형준', '허윤정') and partner == content:
            rows += [row(month, day, '차변', '미지급금', '직원급여', content, out_amt, 0),
                     row(month, day, '대변', '보통예금', '',         content, 0, out_amt)]
            return rows

    # ── 입금 처리 ──
    if in_amt > 0:
        # 결산 이자
        if '결산' in content:
            rows += [row(month, day, '차변', '보통예금', '', content, in_amt, 0),
                     row(month, day, '대변', '이자수익', '', content, 0, in_amt)]
            return rows

        # 웨스트 청약 환급 (입금)
        if content.startswith('웨스트'):
            rows += [row(month, day, '차변', '보통예금', '',         content, in_amt, 0),
                     row(month, day, '대변', '예수금',   '수탁운용', content, 0, in_amt)]
            return rows

        # 청약 환급 (종목명-유형 입금)
        if any(t in content for t in ['－고유','-고유','고유','－일임','-일임','일임',
                                        '－고위험','-고위험','고위험',
                                        '－르네상스','-르네상스','르네상스','르네상','르네예약',
                                        '－채권','-채권']):
            stock_name = extract_duri_stock(content)
            if stock_name:
                rows += [row(month, day, '차변', '보통예금', '',         content, in_amt, 0),
                         row(month, day, '대변', '예수금',   '수탁운용', content, 0, in_amt)]
                return rows

        # 회사 이체 입금 (증권사 → 보통예금, 자금 회수)
        # 두리인베스트먼트(, 키움두리, KB증권두리 → 보통예금/예치금 분개
        if '두리인베스트먼트' in content or '키움두리' in content or 'KB증권두리' in content:
            cp = ''
            if '키움' in content: cp = '키움증권'
            elif 'KB증권' in content: cp = 'KB증권'
            rows += [row(month, day, '차변', '보통예금', '', content, in_amt, 0),
                     row(month, day, '대변', '예치금',   cp, content, 0, in_amt)]
            return rows

        # 비엔케이/에이치엔 (수탁자금 환급) → 보통예금/예수금 분개
        if '비엔케이' in content or '에이치엔' in content:
            rows += [row(month, day, '차변', '보통예금', '',         content, in_amt, 0),
                     row(month, day, '대변', '예수금',   '수탁운용', content, 0, in_amt)]
            return rows

    return None


# ── 두리 한국투자증권 처리 ────────────────────────────────────────────────────
def process_duri_hanguk(df, account_id, cost_basis):
    """
    두리 한국증권 시트 (한국(80...-XX) 형식) 파서.
    헤더 시프트 문제로 컬럼 위치를 인덱스로 직접 접근.
    
    실제 데이터 컬럼 (인덱스 기준):
      0: 거래일, 1: 번호, 2: 거래종류, 3: 종목명, 4: 잔고번호,
      5: 거래수량, 6: 거래단가, 7: 환율, 8: 거래금액, 9: 정산금액, 10: 수수료
    
    처리 대상 거래종류:
    - 사채이자입금: 차변 예치금/선납세금×2 / 대변 이자수익(금융)
    - HTS주식매도/채권매도: 차변 예치금/주식거래수수료/세공 / 대변 단매증/처분이익
    - 채권만기상환: 차변 예치금 / 대변 단매증
    - 예탁금이용료: 차변 예치금 / 대변 이자수익(금융)
    
    무시 (분개 안 함):
    - HTS당사이체출금/입금: 같은 회사 다른 계좌간 이동
    - HTS타사이체출금/입금: IBK에서 이미 분개됨
    - 타사이체입고/공모주입고/HTS타사이체출고/WTS추납대체청약: 자산 이전 (분개장에 별도 분개 없음)
    """
    rows, unmapped = [], []
    if df.empty: return rows, unmapped

    # 첫 행은 헤더, 데이터는 1행부터
    for i in range(1, len(df)):
        r = df.iloc[i]
        date_val = clean(r.iloc[0])
        if not date_val: continue
        m, d = parse_date(str(date_val))
        if not m: continue

        ttype = clean(r.iloc[2])
        stock = clean(r.iloc[3])
        qty = to_int(r.iloc[5])
        price = to_int(r.iloc[6])
        amount = to_int(r.iloc[8])     # 거래금액 (세전)
        settle = to_int(r.iloc[9])     # 정산금액 (세후)
        comm = to_int(r.iloc[10])      # 수수료

        sl = get_stock(stock) if stock else ''
        cp_sec = f'한국투자증권({account_id})'

        # 무시 거래 (자금 이동만)
        if ttype in ('HTS당사이체출금','HTS당사이체입금','HTS당사이체출고','HTS당사이체입고',
                     'HTS타사이체출금','HTS타사이체입금','HTS타사이체출고','HTS타사이체입고',
                     '타사이체입고','타사이체입금','공모주입고','WTS추납대체청약'):
            continue

        # 사채이자입금
        if '사채이자입금' in ttype or '이자입금' in ttype:
            interest = amount      # 세전 이자
            net      = settle      # 입금된 금액 (세후)
            tax_total = interest - net
            if tax_total > 0:
                # 90% 국세 / 10% 지방세 (15.4% 중 14% : 1.4%)
                tax_local = round(tax_total / 11)
                tax_natl  = tax_total - tax_local
            else:
                tax_local = tax_natl = 0
            memo     = f'{stock} 사채이자입금' if stock else '사채이자입금'
            memo_tax = f'{stock} 선납법인세'    if stock else '선납법인세'
            rows.append(row(m, d, '차변', '예치금',   cp_sec,         memo,     net, 0))
            if tax_natl > 0:
                rows.append(row(m, d, '차변', '선납세금', '마포세무서', memo_tax, tax_natl, 0))
            if tax_local > 0:
                rows.append(row(m, d, '차변', '선납세금', '마포구청',   memo_tax, tax_local, 0))
            rows.append(row(m, d, '대변', '이자수익', cp_sec,         memo,     0, interest))
            continue

        # 예탁금이용료
        if '예탁금이용료' in ttype:
            rows += [row(m, d, '차변', '예치금',   cp_sec, '예탁금이용료', settle, 0),
                     row(m, d, '대변', '이자수익', cp_sec, '예탁금이용료', 0, settle)]
            continue

        # 채권만기상환 (수량 큼, 정산금액 = 상환금)
        if '만기상환' in ttype or '만기' in ttype:
            cost_key = (stock, account_id)
            cost = cost_basis.get(cost_key, {}).get('unit_price', 0) * qty if cost_key in cost_basis else amount
            # 분개장 패턴: 차변 예치금 / 대변 단매증 (cost = 거래금액 추정)
            memo = f'{stock} 만기상환' if stock else '만기상환'
            rows += [row(m, d, '차변', '예치금', cp_sec, memo, settle, 0),
                     row(m, d, '대변', '단매증', sl,     memo, 0, amount or settle)]
            continue

        # 주식 매도 (HTS코스닥주식매도 / HTS거래소주식매도)
        if '매도' in ttype and qty > 0 and price > 0:
            # 거래세는 거래금액 - 정산금액 - 수수료
            tax = max(amount - settle - comm, 0)
            cost_key = (stock, account_id)
            if cost_key in cost_basis:
                cost = cost_basis[cost_key]['unit_price'] * qty
                gain = settle + comm + tax - cost  # 매도가 - 매수원가 (정산 + 수수료 + 세금 = 매도가)
            else:
                cost = None
                gain = None

            memo = f'{sl}({qty}주*@{price:,})매도#{account_id}' if sl else f'{stock} 매도'
            rows.append(row(m, d, '차변', '예치금',         cp_sec, memo, settle, 0))
            if comm > 0:
                rows.append(row(m, d, '차변', '주식거래수수료', sl,     memo, comm, 0))
            if tax > 0:
                rows.append(row(m, d, '차변', '세금과공과금',   sl,     memo, tax, 0))

            if cost is not None and cost > 0:
                rows.append(row(m, d, '대변', '단매증', sl, memo, 0, cost))
                if gain > 0:
                    rows.append(row(m, d, '대변', '처분이익', sl, memo, 0, gain))
                elif gain < 0:
                    rows.append(row(m, d, '차변', '처분손실', sl, memo, abs(gain), 0))
                # FIFO 제거
                del cost_basis[cost_key]
            else:
                # 취득가 모름 → 빨간색 처리
                rows.append(row(m, d, '대변', '단매증',   sl, f'{memo} [취득가확인필요]', 0, 0, 'RED'))
                rows.append(row(m, d, '대변', '처분이익', sl, f'{memo} [취득가확인필요]', 0, 0, 'RED'))
                unmapped.append({'날짜': f'{m}/{d}', '종목': stock, '수량': qty, '단가': price,
                                 '비고': '두리 한국증권 취득가 확인 필요'})
            continue

        # 미인식 거래
        unmapped.append({'날짜': f'{m}/{d}', '거래종류': ttype, '종목': stock,
                         '수량': qty, '단가': price, '거래금액': amount, '정산금액': settle})

    return rows, unmapped


# ── 두리 키움증권 처리 ───────────────────────────────────────────────────────
def process_duri_kiwoom(df, account_id, cost_basis):
    """
    두리 키움증권 시트 (키움(2770-XXXX) 형식) 파서.
    2행 헤더 + 2행 데이터 (한투 시트와 비슷한 구조).
    
    행 0 컬럼: 거래일자, 적요명, 수량/좌수, 거래금액, 수수료, 거래세/농특세, 정산금액, ...
    행 1 컬럼: 통화, 거래소, 종목명, 단가/환율, ...
    
    처리 대상:
    - 예탁금이용료(이자)입금: 차변 예치금 / 대변 이자수익(금융)
    - 장내매수/장내매도/KOSDAQ매도/거래소매도: 매수/매도 분개
    무시:
    - 이체입금/이체출금/대체입금/대체출금/타사대체입고: IBK 또는 자산이동
    """
    rows, unmapped = [], []
    if df.empty: return rows, unmapped

    # 데이터는 행 2부터 (2행씩 한 묶음)
    i = 2
    while i < len(df):
        r1 = df.iloc[i]      # 첫 행: 거래일/적요/수량/거래금액/수수료/세/정산
        r2 = df.iloc[i+1] if i+1 < len(df) else None

        date_val = clean(r1.iloc[0])
        if not date_val:
            i += 1; continue
        m, d = parse_date(str(date_val))
        if not m:
            i += 1; continue

        ttype  = clean(r1.iloc[2])      # 적요명 (인덱스 1은 빈 컬럼)
        qty    = to_int(r1.iloc[3])     # 수량/좌수
        amount = to_int(r1.iloc[4])     # 거래금액
        comm   = to_int(r1.iloc[5])     # 수수료
        tax    = to_int(r1.iloc[6])     # 거래세
        settle = to_int(r1.iloc[7])     # 정산금액

        stock = clean(r2.iloc[2]) if r2 is not None else ''
        price = to_int(r2.iloc[3]) if r2 is not None else 0

        sl = get_stock(stock) if stock else ''
        cp_sec = f'키움증권({account_id})'

        # 예탁금이용료
        if '예탁금이용료' in ttype or '이자입금' in ttype:
            if settle > 0:
                rows += [row(m, d, '차변', '예치금',   cp_sec, '예탁금이용료(이자)입금', settle, 0),
                         row(m, d, '대변', '이자수익', cp_sec, '예탁금이용료(이자)입금', 0, settle)]
            i += 2; continue

        # 자금 이체 (이체입금/이체출금/대체입금/대체출금) - IBK에서 처리
        if any(k in ttype for k in ['이체입금','이체출금','대체입금','대체출금']):
            i += 2; continue

        # 타사대체입고 (자산 이동 - 분개 없음)
        if '타사대체입고' in ttype or '입고' in ttype:
            i += 2; continue

        # 매수
        if '매수' in ttype and qty > 0 and price > 0:
            cost = qty * price
            memo = f'{sl}({qty}주*@{price:,})매수#{account_id}' if sl else f'{stock} 매수'
            rows.append(row(m, d, '차변', '단매증', sl, memo, cost, 0))
            if comm > 0:
                rows.append(row(m, d, '차변', '주식거래수수료', sl, memo, comm, 0))
            rows.append(row(m, d, '대변', '예치금', cp_sec, memo, 0, cost + comm))
            # 이동평균 취득가 갱신
            cost_key = (stock, account_id)
            if cost_key in cost_basis:
                old = cost_basis[cost_key]
                new_qty = old['qty'] + qty
                new_avg = (old['unit_price']*old['qty'] + price*qty) / new_qty if new_qty else 0
                cost_basis[cost_key] = {'unit_price': new_avg, 'qty': new_qty}
            else:
                cost_basis[cost_key] = {'unit_price': price, 'qty': qty}
            i += 2; continue

        # 매도
        if '매도' in ttype and qty > 0 and price > 0:
            cost_key = (stock, account_id)
            if cost_key in cost_basis and cost_basis[cost_key]['qty'] >= qty:
                avg = cost_basis[cost_key]['unit_price']
                cost = avg * qty
                gain = settle + comm + tax - cost
                # FIFO/이평 차감
                cost_basis[cost_key]['qty'] -= qty
                if cost_basis[cost_key]['qty'] == 0:
                    del cost_basis[cost_key]
            else:
                cost = None
                gain = None

            memo = f'{sl}({qty}주*@{price:,})매도#{account_id}' if sl else f'{stock} 매도'
            rows.append(row(m, d, '차변', '예치금', cp_sec, memo, settle, 0))
            if comm > 0:
                rows.append(row(m, d, '차변', '주식거래수수료', sl, memo, comm, 0))
            if tax > 0:
                rows.append(row(m, d, '차변', '세금과공과금', sl, memo, tax, 0))

            if cost is not None and cost > 0:
                rows.append(row(m, d, '대변', '단매증', sl, memo, 0, cost))
                if gain > 0:
                    rows.append(row(m, d, '대변', '처분이익', sl, memo, 0, gain))
                elif gain < 0:
                    rows.append(row(m, d, '차변', '처분손실', sl, memo, abs(gain), 0))
            else:
                rows.append(row(m, d, '대변', '단매증',   sl, f'{memo} [취득가확인필요]', 0, 0, 'RED'))
                rows.append(row(m, d, '대변', '처분이익', sl, f'{memo} [취득가확인필요]', 0, 0, 'RED'))
                unmapped.append({'날짜': f'{m}/{d}', '종목': stock, '수량': qty, '단가': price,
                                 '비고': '두리 키움 취득가 확인 필요'})
            i += 2; continue

        # 미인식
        unmapped.append({'날짜': f'{m}/{d}', '거래종류': ttype, '종목': stock,
                         '수량': qty, '단가': price, '거래금액': amount, '정산': settle})
        i += 2

    return rows, unmapped


# ── 두리 메리츠증권 처리 ──────────────────────────────────────────────────────
def process_duri_meritz(df, account_id, cost_basis):
    """
    메리츠증권 시트 (메리츠(3045-XXXX-XX) 형식).
    2행 헤더 + 2행 데이터, 거래 거의 없고 예탁금이용료만 있는 경우 많음.
    
    행 0: ['거래일자', '종목코드', '수량', '수수료', '세전이자', '거래금액', ...]
    행 1: ['거래적요', '종목명', '단가', '제세금', '신용이자', '반영금액', ...]
    
    분개장 패턴: 차변 예치금 / 대변 이자수익(금융)
    """
    rows, unmapped = [], []
    if df.empty: return rows, unmapped

    i = 2
    while i < len(df):
        r1 = df.iloc[i]
        r2 = df.iloc[i+1] if i+1 < len(df) else None

        date_val = clean(r1.iloc[0])
        if not date_val:
            i += 1; continue
        m, d = parse_date(str(date_val))
        if not m:
            i += 1; continue

        # 메리츠는 행0에 거래일자/금액들, 행1에 적요/단가
        ttype = clean(r2.iloc[0]) if r2 is not None else ''
        # 거래금액(반영금액)이 어디에 있나 - 테스트 데이터 기준 행1 인덱스 5에 반영금액
        amount = to_int(r2.iloc[5]) if r2 is not None else 0
        if amount == 0:
            amount = to_int(r1.iloc[5])  # fallback

        cp_sec = f'메리츠증권({account_id})'

        if '예탁금이용료' in ttype:
            if amount > 0:
                rows += [row(m, d, '차변', '예치금',   cp_sec, '예탁금이용료', amount, 0),
                         row(m, d, '대변', '이자수익', cp_sec, '예탁금이용료', 0, amount)]
            i += 2; continue

        # 은행이체출금 (대변 예치금만 - IBK 짝 분개에서 처리)
        if '은행이체' in ttype:
            i += 2; continue

        # 미인식
        unmapped.append({'날짜': f'{m}/{d}', '거래적요': ttype, '금액': amount})
        i += 2

    return rows, unmapped


# ── 두리 카드이용내역 처리 ────────────────────────────────────────────────────
def process_duri_card_usage(df):
    """
    두리 '카드이용내역' 시트 파서 (회사가 정리한 카드 사용 통합본).
    헤더: [날짜, 내역, 금액, 비고, 카드사]
    
    분개: 차변 비용계정 / 대변 미지급금
    카드사 → 거래처 매핑:
    - IBK기업 → 기업카드#5318 (또는 비슷)
    - 우리 → 우리카드#3939
    - 빈칸 → 단순 미지급금 (분개장 거래처 보고 매칭)
    
    비고 키워드별 비용계정:
    - 식대/탕비/생수/커피/연회비 → 복리후생비(판)
    - 접대/골프라운딩/조화 → 접대비(기업업무추진비)(판)
    - 주차비/주유비/대리운전/항공권 → 여비교통비(판)
    - 소모품/탕비용품 → 소모품비(판)
    - 등기부등본/인감/등기우편 → 지급수수료(판)
    """
    rows, unmapped = [], []
    if df.empty: return rows, unmapped

    # 헤더 찾기 (날짜/내역/금액/비고/카드사)
    hdr = None
    for i in range(min(5, len(df))):
        vals = [clean(v) for v in df.iloc[i].values]
        if '날짜' in vals and '내역' in vals and '금액' in vals:
            hdr = i; break
    if hdr is None: return rows, unmapped

    df.columns = df.iloc[hdr]
    df = df.iloc[hdr+1:].reset_index(drop=True)

    for _, r in df.iterrows():
        date_val = r.get('날짜', '')
        if not clean(date_val): continue
        m, d = parse_date(str(date_val))
        if not m: continue

        amt = to_int(r.get('금액', 0))
        if amt <= 0: continue

        item = clean(r.get('내역', ''))
        memo_field = clean(r.get('비고', ''))
        card = clean(r.get('카드사', ''))

        # 카드사 → 거래처 매핑
        if 'IBK' in card or '기업' in card:
            cp = '기업카드#5318'
        elif '우리' in card:
            cp = '우리카드#3939'
        else:
            cp = '카드'

        # 비용계정 분류
        memo_full = item + ' ' + memo_field
        if any(k in memo_full for k in ['식대','탕비','생수','커피','커피캡슐','연회비','부식','회식']):
            acct = '복리후생비'
        elif any(k in memo_full for k in ['접대','골프라운딩','골프','조화','회식']):
            acct = '접대비'
        elif any(k in memo_full for k in ['주차비','주유','대리운전','항공권','출장','여비','교통비']):
            acct = '여비교통비'
        elif any(k in memo_full for k in ['소모품','탕비용품','다이소']):
            acct = '소모품비'
        elif any(k in memo_full for k in ['등기부','인감','등기우편','우체국','발급']):
            acct = '지급수수료'
        else:
            unmapped.append({'날짜': f'{m}/{d}', '금액': amt,
                             '내역': item, '비고': memo_field, '카드사': card})
            continue

        memo = memo_field if memo_field else item
        rows += [row(m, d, '차변', acct,      '', memo, amt, 0),
                 row(m, d, '대변', '미지급금', cp, memo, 0, amt)]

    return rows, unmapped


# ── 두리 납입내역 처리 ────────────────────────────────────────────────────────
def process_duri_napip(df):
    """
    두리 '납입내역' 시트 파서.
    헤더: [납입일자, 구분, 종목, 수량, 단가, 납입금, 청약수수료, 납입총액, 납입처]
    
    분개장에서 청약 납입은 IBK 출금에서 이미 분개됨 (예수금/보통예금).
    이 시트는 회사 내부 정리 자료라 자동변환에서는 무시 (중복 방지).
    추후 IBK 누락 거래 보강용으로 활용 가능.
    """
    return [], []  # 무시


# ── 비씨카드 처리 ─────────────────────────────────────────────────────────────
def process_card(df):
    """
    비씨카드 시트는 두 영역으로 구성:
    - 첫 영역(헤더1=[No,이용구분,거래일,결제일,...]): 결제 완료분 (전월 사용)
    - 두 번째 영역(헤더2=[No,거래일,이용구분,결제예정일자,...]): 결제 예정분 (당월 사용)
    각 영역의 컬럼 순서가 다르므로 각각 헤더 인식해서 처리.
    """
    rows, unmapped = [], []

    # 헤더 후보 위치 모두 찾기
    hdr_positions = []
    for i, r in df.iterrows():
        vals = [clean(v) for v in r.values]
        if '거래내용1' in vals and '승인금액' in vals:
            hdr_positions.append(i)
    if not hdr_positions:
        return rows, unmapped

    # 각 영역을 차례로 처리
    # 정책: "결제예정일자" 컬럼 있는 영역만 처리 (당월 사용 = 미결제)
    # "결제일" 컬럼 영역 = 전월 사용분의 결제 완료. 이미 분개됐으므로 무시.
    for idx, hdr in enumerate(hdr_positions):
        next_hdr = hdr_positions[idx+1] if idx+1 < len(hdr_positions) else len(df)
        section = df.iloc[hdr+1:next_hdr].copy()
        section.columns = df.iloc[hdr]
        cols = [clean(c) for c in section.columns]

        # 결제 완료 영역(전월 사용분)은 건너뛰기
        if '결제예정일자' not in cols:
            continue

        for _, r in section.iterrows():
            date_val = r.get('거래일') or r.get('결제예정일자')
            if not clean(date_val): continue
            month, day = parse_date(str(date_val))
            if not month: continue

            amt = to_int(r.get('승인금액', 0))
            if amt <= 0: continue

            t1 = clean(r.get('거래내용1', ''))
            t2 = clean(r.get('거래내용2', ''))
            combined = t1 + ' ' + t2

            if any(k in combined for k in ['복리후생비', '식대', '부식비', '회식비']):
                acct = '복리후생비'
            elif '접대비' in combined:
                acct = '접대비'
            elif any(k in combined for k in ['교통비', '출장', '여비']):
                acct = '여비교통비'
            elif '지급수수료' in combined or '수수료' in combined:
                acct = '지급수수료'
            else:
                unmapped.append({'날짜': f'{month}/{day}', '금액': amt,
                                 '거래내용1': t1, '거래내용2': t2})
                continue

            memo = t2 if t2 else t1
            rows += [row(month, day, '차변', acct,      '',               memo, amt, 0),
                     row(month, day, '대변', '미지급금', '비씨(7964)카드', memo, 0,   amt)]
    return rows, unmapped


# ── 교보증권 처리 ─────────────────────────────────────────────────────────────
def process_kyobo(df, cost_basis):
    rows, unmapped = [], []
    current_acct = ''
    current_header_row = None
    col_map = {}

    for idx in range(len(df)):
        r    = df.iloc[idx]
        row0 = clean(r.iloc[0])

        if re.match(r'^\d{4}-\d{5}-\d{2}', row0):
            current_acct = row0; current_header_row = None; col_map = {}; continue

        if '거래일자' in row0:
            col_map = {clean(val): ci for ci, val in enumerate(r)}
            current_header_row = idx; continue

        if current_header_row is None or not col_map: continue

        date_val = r.iloc[col_map.get('거래일자', 0)]
        if not clean(date_val) or clean(date_val) in ('거래내역 없음', 'NaN'): continue
        month, day = parse_date(date_val)
        if not month: continue

        ttype     = clean(r.iloc[col_map.get('적요명', 1)])
        stock     = clean(r.iloc[col_map.get('종목명(거래상대명)', 2)])
        qty       = to_int(r.iloc[col_map.get('수량', 3)])
        price     = to_int(r.iloc[col_map.get('단가', 4)])
        trade_amt = to_int(r.iloc[col_map.get('거래금액', 5)])
        settle    = to_int(r.iloc[col_map.get('정산금액', 6)])
        comm      = to_int(r.iloc[col_map.get('수수료', 7)])
        tax_raw   = clean(r.iloc[col_map.get('제세금', 8)])
        tax       = to_int(tax_raw.replace(',', '') if tax_raw else '0')

        acct_abbrev = re.sub(r'[^\d-]', '', current_acct)
        acct_short  = acct_abbrev.replace('1020-', '교보')
        cp_sec      = f'교보증권({acct_abbrev})'
        sl = get_stock(stock) if stock else ''

        if '타사대체입고' in ttype and stock and qty > 0 and price > 0:
            cost = qty * price
            memo = f'{sl}({qty}주*@{price:,})입고#{acct_short}'
            rows += [row(month, day, '차변', '단매증', sl, memo, cost, 0),
                     row(month, day, '대변', '선급금', sl, memo, 0, cost)]
            cost_basis[(stock, current_acct)] = {'unit_price': price, 'qty': qty}

        elif '계좌대체입고' in ttype or '계좌대체출고' in ttype:
            pass

        elif '채권이자입금' in ttype and trade_amt > 0:
            national_tax = round(tax * 10 / 11)
            local_tax    = tax - national_tax
            memo = f'{sl} 채권이자입금' if sl else f'{stock} 채권이자입금'
            rows += [row(month, day, '차변', '예치금',   cp_sec,        memo, settle,       0),
                     row(month, day, '차변', '선납세금', '영등포세무서', memo, national_tax, 0),
                     row(month, day, '차변', '선납세금', '영등포구청',   memo, local_tax,    0),
                     row(month, day, '대변', '이자수익', '',             memo, 0, trade_amt)]

        elif '배당금입금' in ttype or '배당금입급' in ttype:
            # 분배금수익 분개 (예치금 / 분배금수익)
            memo = f'{sl} 배당금입금' if sl else f'{stock} 배당금입금'
            rows += [row(month, day, '차변', '예치금',     cp_sec, memo, settle, 0),
                     row(month, day, '대변', '분배금수익', cp_sec, memo, 0, settle)]

        elif '계좌대체출금' in ttype or '계좌대체입금' in ttype or '은행이체입금' in ttype:
            # 같은 회사 계좌간 이동 또는 IBK 짝 분개 - 무시
            pass

        elif '은행이체출금' in ttype and re.match(r'^O[하채권이][가-힣A-Za-z0-9]+납입$', stock):
            rows += [row(month, day, '차변', '보통예금', '',     stock, settle, 0),
                     row(month, day, '대변', '예치금',   cp_sec, stock, 0, settle)]

        elif '은행이체출금' in ttype and '수익금출금' in stock:
            rows += [row(month, day, '차변', '보통예금', '',     stock, settle, 0),
                     row(month, day, '대변', '예치금',   cp_sec, stock, 0, settle)]

        elif any(k in ttype for k in ['매도', '현금매도']) and qty > 0 and price > 0:
            # 채권은 (수량×단가)/10 형태라 settle 컬럼을 그대로 사용 (이미 정확)
            is_bond = '채권' in ttype
            memo     = f'{sl}({qty}주*@{price:,})매도#{acct_short}'
            cost_key = (stock, current_acct)
            rows += [row(month, day, '차변', '예치금',         cp_sec, memo, settle, 0),
                     row(month, day, '차변', '주식거래수수료', sl,     memo, comm,   0),
                     row(month, day, '차변', '세금과공과금',   sl,     memo, tax,    0)]
            if cost_key in cost_basis:
                acq_cost = cost_basis[cost_key]['unit_price'] * qty
                gain     = settle + comm + tax - acq_cost
                rows    += [row(month, day, '대변', '단매증', sl, memo, 0, acq_cost)]
                if gain > 0:
                    rows += [row(month, day, '대변', '처분이익', sl, memo, 0, gain)]
                elif gain < 0:
                    rows += [row(month, day, '차변', '처분손실', sl, memo, abs(gain), 0)]
                # 부분 매도 처리 (이동평균)
                cost_basis[cost_key]['qty'] -= qty
                if cost_basis[cost_key]['qty'] <= 0:
                    del cost_basis[cost_key]
            else:
                rows += [row(month, day, '대변', '단매증',   sl, memo, 0, 0, 'RED'),
                         row(month, day, '대변', '처분이익', sl, memo, 0, 0, 'RED')]
                unmapped.append({'날짜': f'{month}/{day}', '종목': stock,
                                 '수량': qty, '단가': price,
                                 '비고': f'교보 취득가 확인 필요 ({acct_short})'})

        elif '매수' in ttype and qty > 0 and price > 0:
            # 채권은 (수량×단가)/10 형태라 시트의 거래금액 컬럼을 직접 사용
            is_bond = '채권' in ttype
            cost = trade_amt if (is_bond and trade_amt > 0) else qty * price
            total_out = cost + comm
            memo      = f'{sl}({qty}주*@{price:,})매수#{acct_short}'
            rows += [row(month, day, '차변', '단매증',         sl,     memo, cost, 0),
                     row(month, day, '차변', '주식거래수수료', sl,    memo, comm, 0),
                     row(month, day, '대변', '예치금',         cp_sec, memo, 0, total_out)]
            # 취득가 등록 (이동평균)
            cost_key = (stock, current_acct)
            if cost_key in cost_basis:
                old = cost_basis[cost_key]
                new_qty = old['qty'] + qty
                # 채권이면 unit_price도 비율 조정
                eff_price = cost / qty if qty else price
                new_avg = (old['unit_price']*old['qty'] + eff_price*qty) / new_qty if new_qty else 0
                cost_basis[cost_key] = {'unit_price': new_avg, 'qty': new_qty}
            else:
                eff_price = cost / qty if qty else price
                cost_basis[cost_key] = {'unit_price': eff_price, 'qty': qty}

        elif ttype and ttype not in ('거래내역 없음',):
            if trade_amt > 0 or settle > 0:
                unmapped.append({'날짜': f'{month}/{day}', '거래유형': ttype,
                                 '종목': stock, '금액': settle or trade_amt,
                                 '비고': f'교보 미분류 ({acct_short})'})

    return rows, unmapped


# ── 한국투자증권 처리 ─────────────────────────────────────────────────────────
def parse_hantoo_sheet(df, account_id):
    trades, stock_by_date, wts_stocks = [], {}, set()

    hdr = None
    for i in range(min(15, len(df))):
        if any('거래일' in str(v) for v in df.iloc[i].astype(str).values):
            hdr = i; break
    if hdr is None: return trades, stock_by_date, wts_stocks

    i = hdr + 2  # 한투 항상 2행 포맷
    while i < len(df) - 1:
        try:
            r1 = df.iloc[i]
            r2 = df.iloc[i + 1] if i + 1 < len(df) else None

            date_val   = r1.iloc[0]
            trade_type = clean(r1.iloc[1])
            stock_name = clean(r1.iloc[2])
            qty        = to_int(r1.iloc[3])
            amount     = abs(to_int(r1.iloc[4]))
            commission = to_int(r1.iloc[5])
            net        = to_int(r1.iloc[7]) if len(r1) > 7 else 0
            unit_price = to_int(r2.iloc[3]) if r2 is not None and len(r2) > 3 else 0
            tax        = to_int(r2.iloc[5]) if r2 is not None and len(r2) > 5 else 0

            month, day = parse_date(date_val)
            if not month or not trade_type: i += 2; continue

            skip_kw = ['공모주입고','대여주식입고','대여주식출고','현금주식출고',
                       '출고수수료','HTS출고수수료','타사이체입금']
            if any(k in trade_type for k in skip_kw): i += 2; continue

            if 'WTS추납' in trade_type:
                sn = extract_stock_from_text(stock_name) or stock_name
                if get_stock(sn): wts_stocks.add(sn)
                i += 2; continue

            if 'HTS당사이체입고' in trade_type and not get_stock(stock_name):
                i += 2; continue

            trades.append({'month': month, 'day': day, 'type': trade_type,
                           'stock': stock_name, 'qty': qty, 'commission': commission,
                           'tax': tax, 'unit_price': unit_price, 'net': net,
                           'amount': amount, 'account_id': account_id})

            if stock_name and any(k in trade_type for k in ['입고','입금','매수','이체입고']):
                key = (month, day)
                if key not in stock_by_date: stock_by_date[key] = []
                if stock_name not in stock_by_date[key]:
                    stock_by_date[key].append(stock_name)
        except Exception:
            pass
        i += 2
    return trades, stock_by_date, wts_stocks


def process_hantoo_trades(all_trades, cost_basis):
    rows, unmapped = [], []
    for t in all_trades:
        m, d    = t['month'], t['day']
        ttype   = t['type']
        stock   = t['stock']
        qty     = t['qty']
        price   = t['unit_price']
        comm    = t['commission']
        tax     = t['tax']
        net     = t['net']
        acct_id = t['account_id']

        sl = get_stock(stock) if stock else ''
        aa = acct_id
        for old, new in [('81229969-01','한투9969'),('81163526-01','한투01'),
                         ('81163526-21','한투21'),('81247132-01','한투고유')]:
            aa = aa.replace(old, new)
        cp_sec = f'한국투자증권({acct_id})'

        if '매도' in ttype and qty > 0 and price > 0:
            memo     = f'{sl}({qty}주*@{price:,})매도#{aa}'
            cost_key = (stock, acct_id)
            cost     = cost_basis[cost_key]['unit_price'] * qty if cost_key in cost_basis else None
            gain     = net + comm + tax - cost if cost is not None else None

            rows += [row(m, d, '차변', '예치금',         cp_sec, memo, net,  0),
                     row(m, d, '차변', '주식거래수수료', sl,     memo, comm, 0),
                     row(m, d, '차변', '세금과공과금',   sl,     memo, tax,  0)]

            if cost is not None:
                rows += [row(m, d, '대변', '단매증', sl, memo, 0, cost)]
                if gain > 0:   rows += [row(m, d, '대변', '처분이익', sl, memo, 0, gain)]
                elif gain < 0: rows += [row(m, d, '차변', '단매증',   sl, memo, abs(gain), 0)]
            else:
                rows += [row(m, d, '대변', '단매증',   sl, memo, 0, 0, 'RED'),
                         row(m, d, '대변', '처분이익', sl, memo, 0, 0, 'RED')]
                unmapped.append({'날짜': f'{m}/{d}', '종목': stock, '수량': qty,
                                 '단가': price, '비고': '취득가 확인 필요'})
            if cost_key in cost_basis: del cost_basis[cost_key]

        elif any(k in ttype for k in ['입고','이체입고','이체입금']) \
             and stock and qty > 0 and price > 0:
            cost = qty * price
            memo = f'{sl}({qty}주*@{price:,})입고#{aa}'
            rows += [row(m, d, '차변', '단매증', sl, memo, cost, 0),
                     row(m, d, '대변', '선급금', sl, memo, 0, cost)]
            cost_basis[(stock, acct_id)] = {'unit_price': price, 'qty': qty}

        elif '매수' in ttype and qty > 0 and price > 0:
            cost      = qty * price
            total_out = cost + comm
            memo      = f'{sl}({qty}주*@{price:,})매수#{aa}'
            rows += [row(m, d, '차변', '단매증',         sl,     memo, cost, 0),
                     row(m, d, '차변', '주식거래수수료', sl,    memo, comm, 0),
                     row(m, d, '대변', '예치금',         cp_sec, memo, 0, total_out)]
            cost_basis[(stock, acct_id)] = {'unit_price': price, 'qty': qty}

        elif any(k in ttype for k in ['예탁금이용료','대여수수료']):
            amt = t.get('amount', 0) or abs(net)
            if amt > 0:
                rows += [row(m, d, '차변', '예치금',   cp_sec, ttype, amt, 0),
                         row(m, d, '대변', '이자수익', '',     ttype, 0,   amt)]

    return rows, unmapped


# ── 새 종목 감지 ──────────────────────────────────────────────────────────────
def detect_new_stocks(all_sheets):
    found, known = set(), set(STOCK_DB.keys())
    noise = {'nan','NaN','종목명(거래상대명)','적요명','거래상대명','종목명','',
             '보통예금','예치금','선급금','미수금','리버사이드파트너스'}
    for xl, sheet, fname in all_sheets:
        if not any(k in sheet for k in ['한국투자증권','한투','교보']): continue
        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            for i in range(len(df)):
                if any('거래일' in str(v) for v in df.iloc[i].astype(str).values):
                    for ri in range(i+2, len(df)):
                        try:
                            v = clean(df.iloc[ri, 2])
                            if v and v not in noise and v not in known \
                               and len(v) >= 2 and not v[0].isdigit():
                                found.add(v)
                        except: pass
                    break
        except: pass
    return found - known


# ── 엑셀 출력 (색상 적용) ─────────────────────────────────────────────────────
def create_excel(all_rows):
    if os.path.exists(TEMPLATE_PATH):
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for c in r: c.value = None
        for c in ws.iter_rows(min_row=1, max_row=1):
            for cell in c: cell.fill = FILL_YELLOW
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        headers = ['월','일','구분','계정과목코드','계정과목명','거래처코드',
                   '거래처명','적요명','차변(출금)','대변(입금)']
        ws.append(headers)
        for cell in ws[1]: cell.fill = FILL_YELLOW

    for i, rd in enumerate(all_rows, start=2):
        data = rd[:10]
        flag = rd[10] if len(rd) > 10 else ''
        for j, v in enumerate(data, start=1):
            cell = ws.cell(row=i, column=j, value=v if v != '' else None)
            if flag == 'RED':
                cell.fill = FILL_RED
                cell.font = Font(bold=True, color='FFFFFF')
            elif flag == 'ORANGE':
                cell.fill = FILL_ORANGE

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ── UI ────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎨 색상 범례")
    st.markdown("🔴 **빨간 행** — 취득가 확인 필요\n금액 0으로 처리됨. 직접 입력 필수.")
    st.markdown("🟠 **주황 행** — 미분류 거래\n분류 규칙 없음. 수동 확인 필요.")
    st.markdown("🟡 **노란 행** — 헤더")
    st.divider()
    st.markdown("### ➕ 새 종목 추가")
    st.markdown("변환 후 하단 **'🆕 신규 종목 감지'** 섹션의 코드를 Claude에게 붙여넣으시면 바로 반영됩니다.")

st.divider()
uploaded_files = st.file_uploader(
    "거래내역 파일 업로드 (.xlsx) — 여러 파일 동시 업로드 가능",
    type=['xlsx'], accept_multiple_files=True
)
st.divider()

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)}개 파일 선택됨: {', '.join(f.name for f in uploaded_files)}")
    if st.button("🔄 변환 시작", type="primary", use_container_width=True):
        with st.spinner("변환 중..."):
            try:
                all_xls = []
                for uploaded in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                        tmp.write(uploaded.read())
                        all_xls.append((uploaded.name, tmp.name))

                all_sheets = []
                for fname, tmp_path in all_xls:
                    xl = pd.ExcelFile(tmp_path)
                    for sheet in xl.sheet_names:
                        all_sheets.append((xl, sheet, fname))

                tmp_paths      = [p for _, p in all_xls]
                fname_combined = '_'.join(f.name.replace('.xlsx','') for f in uploaded_files[:2])
                if len(uploaded_files) > 2:
                    fname_combined += f'_외{len(uploaded_files)-2}개'

                # ── 회사 자동 감지 → AC, STOCK_DB 전환 ──────────────────────
                # globals()를 통해 모듈 레벨 변수 변경 (Streamlit이 코드를 함수로 감쌀 때 global 문법 충돌 방지)
                sheet_names_all = [s for _, s, _ in all_sheets]
                company = detect_company(sheet_names_all)
                if company == 'duri':
                    globals()['AC'] = DURI_AC
                    globals()['STOCK_DB'] = DURI_STOCK_DB
                    st.info(f"🏢 자동 감지: **두리인베스트먼트** (3자리 계정코드)")
                elif company == 'riverside':
                    globals()['AC'] = RIVERSIDE_AC
                    globals()['STOCK_DB'] = RIVERSIDE_STOCK_DB
                    st.info(f"🏢 자동 감지: **리버사이드파트너스** (5자리 계정코드)")
                else:
                    globals()['AC'] = RIVERSIDE_AC
                    globals()['STOCK_DB'] = RIVERSIDE_STOCK_DB
                    st.warning(f"⚠️ 회사 자동 감지 실패. 리버사이드로 처리합니다.")

                all_rows, all_unmapped = [], []
                cost_basis, stock_by_date, all_hantoo = {}, {}, []

                # 두리 전용 시트(아직 파서 미구현 - Phase 2에서 추가 예정)
                duri_only_sheets = ['IBK카드내역', '우리카드내역', '하나카드내역',
                                    '카드이용내역', '은행거래내역', '납입내역',
                                    '키움', '메리츠']

                # 두리에서 한투/교보 형식이 비슷한 시트만 기존 파서로 시도
                # (한국증권, 교보증권 시트 형식이 두리도 거의 동일)

                # 1) 한투 (또는 두리의 한국증권 시트)
                # 리버사이드: '한국투자증권' 시트 → process_hantoo
                # 두리:        '한국(80...)' 시트 → process_duri_hanguk
                for xl, sheet, fname in all_sheets:
                    is_hantoo = any(k in sheet for k in ['한국투자증권','한투'])
                    is_duri_hanguk = company == 'duri' and re.match(r'^한국\(\d', sheet)
                    if is_hantoo:
                        df = pd.read_excel(xl, sheet_name=sheet, header=None)
                        acct_id = sheet
                        for i in range(3):
                            cell = clean(df.iloc[i, 0]) if len(df) > i else ''
                            m = re.search(r'\d{5,}-\d{2}', cell)
                            if m: acct_id = m.group(); break
                        trades, sbd, wts = parse_hantoo_sheet(df, acct_id)
                        all_hantoo.extend(trades)
                        for k, v in sbd.items():
                            if k not in stock_by_date: stock_by_date[k] = []
                            for s in v:
                                if s not in stock_by_date[k]: stock_by_date[k].append(s)
                        stock_by_date.setdefault('__wts__', set()).update(wts)
                    elif is_duri_hanguk:
                        df = pd.read_excel(xl, sheet_name=sheet, header=None)
                        # 시트명에서 계좌번호 추출
                        m_acct = re.search(r'\((\d{5,}-\d+)\)', sheet)
                        acct_id = m_acct.group(1) if m_acct else sheet
                        hr, hu = process_duri_hanguk(df, acct_id, cost_basis)
                        all_rows.extend(hr)
                        all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in hu])

                sec_rows, sec_unmap = process_hantoo_trades(all_hantoo, cost_basis)
                all_rows.extend(sec_rows)
                all_unmapped.extend([{**u, '출처': '한투/한국'} for u in sec_unmap])

                # 2) 교보
                for xl, sheet, fname in all_sheets:
                    if '교보' in sheet:
                        df = pd.read_excel(xl, sheet_name=sheet, header=None)
                        kr, ku = process_kyobo(df, cost_basis)
                        all_rows.extend(kr)
                        all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in ku])

                # 2-B) 두리 키움증권
                if company == 'duri':
                    for xl, sheet, fname in all_sheets:
                        if re.match(r'^키움\(\d', sheet):
                            df = pd.read_excel(xl, sheet_name=sheet, header=None)
                            m_acct = re.search(r'\((\d+-\d+)\)', sheet)
                            acct_id = m_acct.group(1) if m_acct else sheet
                            kr, ku = process_duri_kiwoom(df, acct_id, cost_basis)
                            all_rows.extend(kr)
                            all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in ku])

                # 2-C) 두리 메리츠증권
                if company == 'duri':
                    for xl, sheet, fname in all_sheets:
                        if re.match(r'^메리츠\(\d', sheet):
                            df = pd.read_excel(xl, sheet_name=sheet, header=None)
                            m_acct = re.search(r'\((\d+-\d+-\d+)\)', sheet)
                            acct_id = m_acct.group(1) if m_acct else sheet
                            mr, mu = process_duri_meritz(df, acct_id, cost_basis)
                            all_rows.extend(mr)
                            all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in mu])

                # 3) IBK 기업은행 (리버사이드의 'IBK 기업은행' / 두리의 'IBK기업은행 내역')
                for xl, sheet, fname in all_sheets:
                    is_ibk = any(k in sheet for k in ['IBK기업은행', 'IBK 기업은행'])
                    is_riverside_ibk = company == 'riverside' and any(k in sheet for k in ['IBK','기업은행'])
                    if is_ibk or is_riverside_ibk:
                        # 단, IBK카드내역은 제외 (카드 파서로 가야 함)
                        if 'IBK카드' in sheet or '카드내역' in sheet:
                            continue
                        df = pd.read_excel(xl, sheet_name=sheet, header=None)
                        # 회사별로 다른 파서 호출
                        if company == 'duri':
                            ir, iu = process_duri_ibk(df, stock_by_date)
                        else:
                            ir, iu = process_ibk(df, stock_by_date)
                        all_rows.extend(ir)
                        all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in iu])

                # 4) 비씨카드 (리버사이드 전용 - 두리는 다른 카드 시트)
                if company != 'duri':
                    for xl, sheet, fname in all_sheets:
                        if any(k in sheet for k in ['비씨','세부']):
                            df = pd.read_excel(xl, sheet_name=sheet, header=None)
                            cr, cu = process_card(df)
                            all_rows.extend(cr)
                            all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in cu])

                # 4-B) 두리 카드이용내역 (회사 정리 통합본)
                if company == 'duri':
                    for xl, sheet, fname in all_sheets:
                        if sheet == '카드이용내역':
                            df = pd.read_excel(xl, sheet_name=sheet, header=None)
                            cr, cu = process_duri_card_usage(df)
                            all_rows.extend(cr)
                            all_unmapped.extend([{**u, '출처': f'{fname}>{sheet}'} for u in cu])

                # 5) 두리 전용 시트들은 미처리 시트로 보고 (Phase 2에서 파서 추가 예정)
                if company == 'duri':
                    duri_unprocessed = []
                    for xl, sheet, fname in all_sheets:
                        # 이미 처리한 시트(한국, 교보, IBK기업은행 내역) 제외
                        if re.match(r'^한국\(\d', sheet) or '교보(' in sheet:
                            continue
                        if 'IBK기업은행' in sheet:
                            continue
                        # 두리 전용 시트
                        if any(k in sheet for k in ['IBK카드','우리카드','하나카드',
                                                     '카드이용내역','은행거래내역','납입내역',
                                                     '키움','메리츠','KB(']):
                            duri_unprocessed.append(sheet)
                    if duri_unprocessed:
                        st.warning(f"⏳ 미처리 두리 시트 {len(duri_unprocessed)}개 (Phase 2에서 파서 추가 예정): {', '.join(duri_unprocessed)}")

                # 새 종목 감지
                new_stocks = detect_new_stocks(all_sheets)

                for p in tmp_paths:
                    try: os.unlink(p)
                    except: pass

                if not all_rows:
                    st.error("변환된 데이터가 없습니다.")
                    st.info("인식 가능한 시트명: '한국투자증권'/'한투', '교보', 'IBK'/'기업은행'/'은행', '비씨'/'카드'/'세부'")
                else:
                    excel_out = create_excel(all_rows)
                    dr       = sum(r[8] for r in all_rows if r[8] != '')
                    cr_      = sum(r[9] for r in all_rows if r[9] != '')
                    red_rows = sum(1 for r in all_rows if len(r) > 10 and r[10] == 'RED')

                    st.success("✅ 변환 완료!")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("전표 행 수",   f"{len(all_rows):,}행")
                    c2.metric("🔴 확인 필요", f"{red_rows:,}행")
                    c3.metric("🟠 미분류",    f"{len(all_unmapped):,}건")
                    c4.metric("차대변",       "✅ 일치" if dr == cr_ else f"⚠️ {abs(dr-cr_):,.0f}원 차이")
                    st.info(f"차변 합계: {dr:,.0f}원  |  대변 합계: {cr_:,.0f}원")

                    if red_rows > 0:
                        st.error(f"🔴 {red_rows}행: 취득가 확인 필요 — 엑셀에서 빨간 행 찾아 금액 직접 입력하세요.")

                    if all_unmapped:
                        with st.expander(f"🟠 미분류 {len(all_unmapped)}건 상세"):
                            st.dataframe(pd.DataFrame(all_unmapped), use_container_width=True)

                    st.download_button(
                        "📥 변환 파일 다운로드",
                        data=excel_out,
                        file_name=f"더존업로드_{fname_combined}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, type="primary"
                    )

                    # 새 종목 섹션
                    if new_stocks:
                        st.divider()
                        st.warning(f"🆕 {len(new_stocks)}개 신규 종목 감지 — 아래 코드를 Claude에게 붙여넣으세요.")
                        lines = [f"    '{s}': '주식#코스닥#{s}',  # ← 코스피/코스닥/채권#상장 확인 필요"
                                 for s in sorted(new_stocks)]
                        st.code("# STOCK_DB 추가분:\n" + "\n".join(lines), language='python')

            except Exception as e:
                st.error(f"오류: {e}")
                import traceback; st.code(traceback.format_exc())
else:
    st.info("거래내역 파일을 업로드하면 변환 버튼이 활성화됩니다.")

st.divider()
st.caption("종목 추가, 거래처 수정, 미분류 처리 등은 Claude에게 요청해 주세요.")
