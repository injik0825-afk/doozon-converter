"""
parsers/securities/kyobo.py
교보증권 거래내역 파서

컬럼 구조:
row[0]=거래일자, row[1]=번호, row[2]=적요명, row[3]=종목명,
row[4]=수량, row[5]=단가, row[6]=거래금액, row[7]=정산금액,
row[8]=수수료, row[9]=제세금, row[10]=예수금잔고, row[11]=유가증권잔고
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class KyoboParser(BaseParser):
    name = "교보증권"
    
    TYPE_MAP = {
        '은행이체입금': '이체입금',
        '은행이체송금': '이체출금',
        '은행이체출금': '이체출금',
        '계좌대체입금': '당사이체입금',
        '계좌대체출금': '당사이체출금',
        '배당금입금': '배당금',
        '채권장내당일매수': '매수',
        '채권장내당일매도': '매도',
        '주식당일매수': '매수',
        '주식당일매도': '매도',
        '사채이자입금': '채권이자',
        '예탁금이용료': '예탁금이자',
    }
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        account_id = kwargs.get('account_id', 'unknown')
        events = []
        
        for idx in range(1, len(df)):  # 헤더 1행
            row = df.iloc[idx]
            try:
                event = self._parse_row(row, idx, account_id)
                if event:
                    events.append(event)
            except Exception as e:
                print(f"[{self.name}] {account_id} 행 {idx} 파싱 실패: {e}")
                continue
        
        return events
    
    def _parse_row(self, row, idx: int, account_id: str):
        날짜 = self.parse_date(row[0])
        if 날짜 is None:
            return None
        
        적요 = self.clean_text(row[2])
        event_type = self.TYPE_MAP.get(적요, 적요)
        
        종목명 = self.clean_text(row[3]) if len(row) > 3 else ''
        수량 = self.safe_float(row[4]) if len(row) > 4 else 0
        단가 = self.safe_float(row[5]) if len(row) > 5 else 0
        거래금액 = self.safe_int(row[6]) if len(row) > 6 else 0
        정산금액 = self.safe_int(row[7]) if len(row) > 7 else 0
        수수료 = self.safe_int(row[8]) if len(row) > 8 else 0
        제세금 = self.safe_int(row[9]) if len(row) > 9 else 0
        
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
            세금=제세금,
            적요원본=적요,
        )
        
        if 정산금액 and 정산금액 != 거래금액:
            event.extra['정산금액'] = 정산금액
        
        # 종목유형 추정 - 교보는 채권 거래가 많음
        if '채권' in 적요 or '채권' in 종목명:
            event.종목유형 = '채권'
            if '국민주택' in 종목명:
                event.시장구분 = '국민주택1종채권'
            else:
                event.시장구분 = '회사채'
        elif event_type in ('매수', '매도', '배당금'):
            event.종목유형 = '주식'
        
        return event
