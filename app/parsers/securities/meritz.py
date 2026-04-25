"""
parsers/securities/meritz.py
메리츠증권 거래내역 파서

구조가 독특: 2행씩 한 거래
- 홀수행: 거래일자, 수수료, 세전이자, 거래금액, 예수금잔고 등
- 짝수행: 거래적요(예탁금이용료), 종목명, 거래번호

두리인베스트먼트의 메리츠는 이자 계좌만 3개 있음 (5320, 5325, 5327)
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class MeritzParser(BaseParser):
    name = "메리츠증권"
    
    TYPE_MAP = {
        '예탁금이용료': '예탁금이자',
        '은행이체출금': '이체출금',
        '은행이체입금': '이체입금',
    }
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        account_id = kwargs.get('account_id', 'unknown')
        events = []
        
        i = 2  # 헤더 2줄
        while i < len(df):
            row = df.iloc[i]
            next_row = df.iloc[i+1] if i+1 < len(df) else None
            
            try:
                event = self._parse_pair(row, next_row, i, account_id)
                if event:
                    events.append(event)
            except Exception as e:
                print(f"[{self.name}] {account_id} 행 {i} 파싱 실패: {e}")
            
            i += 2
        
        return events
    
    def _parse_pair(self, row, next_row, idx: int, account_id: str):
        날짜 = self.parse_date(row[0])
        if 날짜 is None:
            return None
        
        # 다음 행에서 적요 (row[1]이 거래적요 코드)
        적요 = ''
        if next_row is not None:
            적요 = self.clean_text(next_row[0])
        
        event_type = self.TYPE_MAP.get(적요, 적요)
        
        세전이자 = self.safe_int(row[4]) if len(row) > 4 else 0
        거래금액 = self.safe_int(row[5]) if len(row) > 5 else 0
        수수료 = self.safe_int(row[3]) if len(row) > 3 else 0
        예수금잔고 = self.safe_int(row[10]) if len(row) > 10 else 0
        
        event = RawEvent(
            날짜=날짜,
            event_type=event_type,
            원천=account_id,
            원천행번호=idx,
            총액=세전이자 if event_type == '예탁금이자' else 거래금액,
            수수료=수수료,
            적요원본=적요,
        )
        
        if event_type == '예탁금이자':
            event.세금 = 0  # 메리츠 예탁금이자는 보통 세금 없음
        
        return event
