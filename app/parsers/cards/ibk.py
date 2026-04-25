"""
parsers/cards/ibk.py
IBK 법인카드 거래내역 파서

구조:
row[2]=이용구분(일시불/할부), row[3]=승인일시, row[4]=카드번호,
row[5]=카드별명, row[6]=이용가맹점명, row[7]=승인금액
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class IBKCardParser(BaseParser):
    name = "IBK법인카드"
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        account_id = kwargs.get('account_id', 'IBK카드')
        events = []
        
        # row 3부터 데이터
        for idx in range(3, len(df)):
            row = df.iloc[idx]
            try:
                event = self._parse_row(row, idx, account_id)
                if event:
                    events.append(event)
            except Exception as e:
                continue
        
        return events
    
    def _parse_row(self, row, idx: int, account_id: str):
        승인일시 = self.clean_text(row[3]) if len(row) > 3 else ''
        if not 승인일시 or '승인일시' in 승인일시:
            return None
        
        날짜 = self.parse_date(승인일시)
        if 날짜 is None:
            return None
        
        가맹점명 = self.clean_text(row[6]) if len(row) > 6 else ''
        승인금액 = self.safe_int(row[7]) if len(row) > 7 else 0
        
        if 승인금액 == 0:
            return None
        
        event = RawEvent(
            날짜=날짜,
            event_type='카드승인',
            원천=account_id,
            원천행번호=idx,
            총액=승인금액,
            적요원본=가맹점명,
        )
        event.extra['가맹점'] = 가맹점명
        event.extra['카드'] = 'IBK기업카드'
        
        return event
