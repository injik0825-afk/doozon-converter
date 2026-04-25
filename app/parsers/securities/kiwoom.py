"""
parsers/securities/kiwoom.py
키움증권 거래내역 파서

구조가 특이함: 2행씩 한 거래 (첫 행에 금액, 다음 행에 통화/종목 보조)
row[0]=거래일자(홀수행)/통화(짝수행)
row[1]=거래소 (짝수행에 종목명)
row[2]=적요명
row[3]=수량/좌수 (짝수행은 단가)
row[4]=거래금액
row[5]=수수료
row[6]=거래세
row[7]=정산금액
row[9]=예수금잔고
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class KiwoomParser(BaseParser):
    name = "키움증권"
    
    TYPE_MAP = {
        '예탁금이용료(이자)입금': '예탁금이자',
        '대체출금': '대체출금',
        '대체입금': '대체입금',
        '이체입금(연계은행)': '이체입금',
        '이체출금(연계은행)': '이체출금',
        '타사대체입고': '타사대체입고',
        '타사대체출고': '타사대체출고',
        '장내매수': '매수',
        '장내매도': '매도',
        '배당금입금': '배당금',
    }
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        account_id = kwargs.get('account_id', 'unknown')
        events = []
        
        i = 2  # 헤더 2줄 건너뛰기
        while i < len(df):
            row = df.iloc[i]
            next_row = df.iloc[i+1] if i+1 < len(df) else None
            
            try:
                event = self._parse_pair(row, next_row, i, account_id)
                if event:
                    events.append(event)
            except Exception as e:
                print(f"[{self.name}] {account_id} 행 {i} 파싱 실패: {e}")
            
            i += 2  # 키움은 2행이 한 쌍
        
        return events
    
    def _parse_pair(self, row, next_row, idx: int, account_id: str):
        날짜 = self.parse_date(row[0])
        if 날짜 is None:
            return None
        
        적요 = self.clean_text(row[2])
        event_type = self.TYPE_MAP.get(적요, 적요)
        
        수량 = self.safe_float(row[3])
        거래금액 = self.safe_int(row[4])
        수수료 = self.safe_int(row[5])
        거래세 = self.safe_int(row[6])
        정산금액 = self.safe_int(row[7])
        
        # 종목명/단가 (next_row에서)
        종목명 = ''
        단가 = 0
        if next_row is not None:
            종목명 = self.clean_text(next_row[2]) if len(next_row) > 2 else ''
            단가 = self.safe_float(next_row[3]) if len(next_row) > 3 else 0
        
        event = RawEvent(
            날짜=날짜,
            event_type=event_type,
            원천=account_id,
            원천행번호=idx,
            총액=거래금액 if 거래금액 > 0 else 정산금액,
            종목명=종목명,
            수량=수량,
            단가=단가,
            수수료=수수료,
            거래세=거래세,
            적요원본=적요,
        )
        
        # 정산금액 처리: 배치 합산 이상값 검사
        # 매수: 정산금액이 (거래금액+수수료)의 1.1배 이상이면 배치합산값 → 무시
        # 매도: 정산금액이 0이면 무시 (단순)
        if 정산금액 and 거래금액 > 0:
            expected = 거래금액 + (수수료 or 0) + (거래세 or 0)
            if event_type == '매수' and 정산금액 > expected * 1.1:
                pass  # 배치합산 정산금액 → extra에 넣지 않음
            elif event_type == '매도' and 정산금액 < 거래금액 * 0.3:
                pass  # 이상값 무시
            else:
                if 정산금액 != 거래금액:
                    event.extra['정산금액'] = 정산금액
        
        if 정산금액 and 정산금액 != 거래금액 and '정산금액' not in event.extra:
            event.extra['정산금액'] = 정산금액
        
        # 종목유형 추정
        if event_type in ('매수', '매도', '타사대체입고', '타사대체출고', '배당금'):
            # 주식이 대부분이나 정확한 구분은 회사 설정/포트폴리오 참조 필요
            event.종목유형 = '주식'
        
        return event
