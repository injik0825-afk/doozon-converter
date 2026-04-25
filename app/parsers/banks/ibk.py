"""
parsers/banks/ibk.py
IBK 기업은행 거래내역 파서 (두리인베스트먼트 기준)

컬럼 구조 (row 2가 헤더, row 3부터 데이터):
  col[0]=번호, col[1]=거래일시, col[2]=출금, col[3]=입금,
  col[4]=거래후잔액, col[5]=거래내용, col[6]=상대계좌번호, col[7]=상대은행

핵심 패턴 (분개장 분석 결과):
  자사이체:         두리인베스트(* + 국민은행       → 보통예금 ↔ 보통예금(타계좌)
  증권사이체(입):    키움두리/KB증권두리 + 국민은행    → 보통예금(차) / 예치금(대)
  증권사이체(출):    두리인베스트(주 (상대은행 없음)   → 예치금(차) / 보통예금(대)
  4대보험:          국민연금/고용/산재/국민건강 포함  → 예수금+비용 / 보통예금
  국세납부:          국세조회납부                    → 예수금 / 보통예금
  지방세납부:        지방세납부                      → 예수금 / 보통예금
  이체수수료:        이체수수료                      → 지급수수료(판) / 보통예금
  카드결제:          우리카드결제/비씨카드출금          → 미지급금 / 보통예금
  일임청약출금:      종목명-일임/고유/고위험           → 예수금(차) / 보통예금(대)  거래처:수탁운용
  일임청약환급입금:  웨스트종목명 / ㈜비엔케이증권     → 보통예금(차) / 예수금(대)  거래처:수탁운용
  르네상스청약출금:  종목명-르네상스                  → 선급금(차) / 보통예금(대)
  급여지급:         이름(개인) + 은행               → 미지급금(차) / 보통예금(대)
  퇴직연금DC:       퇴직연금DC                      → 퇴직급여(차) / 보통예금(대)
  퇴직연금부담금:    퇴직연금부담금                   → 퇴직급여(판)(차) / 보통예금(대)
  퇴직연금수수료:    퇴직연금수수료                   → 지급수수료(판)(차) / 보통예금(대)
  통신비:           SKT/KT+번호                    → 복리후생비(판)(차) / 보통예금(대)
  복지비:           조의금/경조금                   → 접대비(판) or 복리후생비(판) / 보통예금(대)
  은행결산이자:      2026년결산                      → 이자수익(금융) / 보통예금(차)
  임대료:           임대료                          → 미지급금(차) / 보통예금(대)
  에이치엔/수탁:     에이치엔인베스트/수탁관련          → 보통예금(차) / 예수금(대)
"""
import re
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class IBKBankParser(BaseParser):
    name = "IBK기업은행"

    # 상대은행별 증권사 거래처 매핑 (입금 시 키움/KB 판단용)
    SECURITIES_BY_CONTENT = {
        '키움': '키움_이체',
        'KB증권': 'KB_이체',
        '교보증권': '교보_이체',
        '메리츠': '메리츠_이체',
        '한국투자': '한투_이체',
    }

    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        account_id = kwargs.get('account_id', 'IBK_메인')
        events = []

        for idx in range(3, len(df)):
            row = df.iloc[idx]
            try:
                event = self._parse_row(row, idx, account_id)
                if event:
                    events.append(event)
            except Exception:
                continue

        return events

    def _parse_row(self, row, idx: int, account_id: str):
        일시_str = self.clean_text(row[1]) if len(row) > 1 else ''
        if not 일시_str or '거래일시' in 일시_str or '계좌' in 일시_str:
            return None

        날짜 = self.parse_date(일시_str)
        if 날짜 is None:
            return None

        출금 = self.safe_int(row[2]) if len(row) > 2 else 0
        입금 = self.safe_int(row[3]) if len(row) > 3 else 0
        거래내용 = self.clean_text(row[5]) if len(row) > 5 else ''
        상대계좌 = self.clean_text(row[6]) if len(row) > 6 else ''
        상대은행 = self.clean_text(row[7]) if len(row) > 7 else ''

        if 출금 == 0 and 입금 == 0:
            return None

        event_type = self._classify(거래내용, 상대은행, 출금 > 0)

        event = RawEvent(
            날짜=날짜,
            event_type=event_type,
            원천=account_id,
            원천행번호=idx,
            총액=출금 if 출금 > 0 else 입금,
            상대계좌=상대계좌,
            상대은행=상대은행,
            적요원본=거래내용,
        )
        event.extra['출금'] = 출금
        event.extra['입금'] = 입금

        return event

    def _classify(self, 내용: str, 상대은행: str, is_출금: bool) -> str:
        n = 내용

        # ── 4대보험 ──
        if '국민연금' in n:     return '국민연금'
        if '국민건강' in n or '건강보험' in n: return '건강보험'
        if '고용보험' in n:     return '고용보험'
        if '산재보험' in n:     return '산재보험'

        # ── 세금 납부 ──
        if '국세조회납부' in n or '국세납부' in n: return '국세납부'
        if '지방세납부' in n or '지방소득세납부' in n: return '지방세납부'

        # ── 카드 결제 ──
        if '우리카드결제' in n: return '카드결제_우리'
        if '하나카드결제' in n: return '카드결제_하나'
        if '비씨카드' in n:     return '카드결제_BC'
        if 'IBK카드' in n or '기업카드결제' in n: return '카드결제_IBK'

        # ── 이체수수료 ──
        if '이체수수료' in n:   return '이체수수료'

        # ── 퇴직연금 ──
        if '퇴직연금수수료' in n: return '퇴직연금수수료'
        if '퇴직연금DC' in n:    return '퇴직연금DC'
        if '퇴직연금부담금' in n or '퇴직연금부담' in n: return '퇴직연금부담금'

        # ── 임대료 ──
        if '임대료' in n:       return '임대료'

        # ── 통신비 (SKT/KT 휴대폰) ──
        if re.match(r'^\d{11}SKT', n) or re.match(r'^01\d{9}', n): return '통신비'
        if re.match(r'^KT\d{10}', n): return '통신비KT'

        # ── 복지비 ──
        if '조의금' in n or '경조' in n: return '접대비_기타'

        # ── 결산이자 ──
        if '결산' in n and is_출금 is False: return '은행이자'

        # ── 일임/수탁 청약 납입/환급 ──
        # 일임/고위험/채권 → 예수금(수탁운용 대신 납입)
        # 고유/르네상스/르네예약 → 선급금(자사 납입)
        ILIM_KEYS = ['일임', '고위험', '채권']
        GOYU_KEYS = ['고유', '르네상스', '르네', '르네예약']
        WEST_KEYS = ['웨스트']
        
        # 내용에 –(em dash) 또는 -(hyphen) 뒤에 위 키워드 포함
        def has_suffix(text, keys):
            import re
            # "종목명－키워드" 또는 "종목명고위험" 등 붙은 형태 모두 허용
            for k in keys:
                if re.search(rf'[-－]?{k}', text):
                    return True
            return False
        
        if any(n.endswith(k) or f'－{k}' in n or f'-{k}' in n for k in ILIM_KEYS):
            return '일임청약출금'
        if any(n.endswith(k) or f'－{k}' in n or f'-{k}' in n for k in GOYU_KEYS):
            return '청약납입_은행'
        if n.startswith('웨스트'):
            return '일임청약환급입금'
        if '비엔케이증권' in n or '㈜비엔케이' in n:
            return '일임청약환급입금'
        if '에이치엔인베스트' in n:
            return '에이치엔입금'

        # ── 미지급금 결제 (관리비 등 계상 후 현금결제) ──
        if re.match(r'^관리\d{4}', n):
            return '미지급금결제'

        # ── 급여 지급 (미지급금 → 보통예금) ──
        if is_출금 and 상대은행 and len(n) <= 5 and not any(
            x in n for x in ['증권', '은행', '보험', '카드', '연금', '세금', '이체']
        ):
            return '급여지급'

        # ── 증권사 이체 입금 ──
        if not is_출금:
            if '키움' in n:   return '증권사이체입금'
            if 'KB증권' in n: return '증권사이체입금'
            if '교보' in n:   return '증권사이체입금'
            if '한국투자' in n or '한투' in n: return '증권사이체입금'
            if '에이치엔인베스트' in n: return '에이치엔입금'

        # ── 자사이체 (두리인베스트먼트 국민은행) ──
        if '두리인베스트' in n and '국민은행' in 상대은행:
            return '자사이체입금' if not is_출금 else '자사이체출금'

        # ── 내부 자금 이동 (증권사로 송금, 상대은행 없음) ──
        if is_출금 and not 상대은행 and '두리인베스트' in n:
            return '증권사이체출금'

        return '기타출금' if is_출금 else '기타입금'
