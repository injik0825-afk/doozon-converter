"""
parsers/cards/summary.py
카드이용내역 통합 시트 파서
(두리의 경우 '카드이용내역' 시트에 분개 힌트인 '비고' 컬럼이 있음)

컬럼: 날짜, 내역, 금액, 비고(용도), 카드사
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class CardSummaryParser(BaseParser):
    name = "카드이용내역통합"
    
    # 비고(용도) → 회계처리 힌트
    USAGE_MAP = {
        '주차비': '여비교통비',
        '식대': '복리후생비',
        '접대': '접대비',
        '골프라운딩': '접대비',
        '회식': '접대비',
        '소모품': '소모품비',
        '커피캡슐': '복리후생비',
        '부식비': '복리후생비',
        '등기부등본': '지급수수료',
        '인감증명': '지급수수료',
        '주유': '차량유지비',
        '유류': '차량유지비',
        '통신': '통신비',
        '도서': '도서인쇄비',
        '교육': '교육훈련비',
        '여비': '여비교통비',
        '출장': '여비교통비',
        '경조': '복리후생비',
        '연회비': '지급수수료',
    }
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        events = []
        
        # row 2부터 데이터 (헤더 2줄)
        for idx in range(2, len(df)):
            row = df.iloc[idx]
            try:
                event = self._parse_row(row, idx)
                if event:
                    events.append(event)
            except Exception:
                continue
        
        return events
    
    def _parse_row(self, row, idx: int):
        날짜 = self.parse_date(row[0])
        if 날짜 is None:
            return None
        
        내역 = self.clean_text(row[1]) if len(row) > 1 else ''
        금액 = self.safe_int(row[2]) if len(row) > 2 else 0
        비고 = self.clean_text(row[3]) if len(row) > 3 else ''
        카드사 = self.clean_text(row[4]) if len(row) > 4 else ''
        
        if 금액 == 0:
            return None
        
        # 비고에서 계정과목 힌트 찾기
        계정과목힌트 = ''
        for keyword, 계정 in self.USAGE_MAP.items():
            if keyword in 비고:
                계정과목힌트 = 계정
                break
        
        event = RawEvent(
            날짜=날짜,
            event_type='카드사용',
            원천='카드이용내역',
            원천행번호=idx,
            총액=금액,
            적요원본=비고,
        )
        event.extra['가맹점'] = 내역
        event.extra['비고'] = 비고
        event.extra['카드사'] = 카드사
        event.extra['계정과목힌트'] = 계정과목힌트
        
        return event
