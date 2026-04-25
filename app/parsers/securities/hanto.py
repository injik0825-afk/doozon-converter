"""
parsers/securities/hanto.py
한국투자증권 거래내역 파서

한투는 계좌가 4개:
- 80777717-01: 공모주 청약 계좌 (공모주 입고/출고)
- 80777717-21: 자금 관리 계좌 (이체)
- 80895963-01: 일반 거래 (사채이자 등)
- 80895963-21: 일반 거래 (이체)

컬럼 구조 (80895963-01 기준):
row[0]=거래일, row[2]=거래종류, row[3]=종목명,
row[5]=거래수량, row[6]=거래단가,
row[8]=거래금액(세전), row[9]=정산금액(세후),
row[10]=수수료, row[11]=거래세, row[13]=세금(원천),
row[15]=예수금잔액
"""
from typing import List
import pandas as pd
from ..base import BaseParser, RawEvent


class HantoParser(BaseParser):
    """한국투자증권 파서"""
    name = "한국투자증권"
    
    # 거래종류별 event_type 매핑
    TYPE_MAP = {
        '사채이자입금': '채권이자',
        '배당금입금': '배당금',
        '예탁금이용료': '예탁금이자',
        '타사이체입금': '이체입금',
        '타사이체출금': '이체출금',
        'HTS당사이체입금': '당사이체입금',
        'HTS당사이체출금': '당사이체출금',
        'HTS당사이체입고': 'HTS당사이체입고',
        'HTS당사이체출고': 'HTS당사이체출고',
        'HTS타사이체입금': '이체입금',
        'HTS타사이체출금': '이체출금',
        'HTS타사이체입고': 'HTS타사이체입고',
        'HTS타사이체출고': 'HTS타사이체출고',
        '공모주입고': '공모주입고',
        '공모주출고': '공모주출고',
        'WTS추납대체청약': '청약납입',
        '이체입금': '이체입금',
        '이체출금': '이체출금',
        '장내매수': '매수',
        '장내매도': '매도',
        'KOSDAQ매도': '매도',
        'KOSDAQ매수': '매수',
        'HTS코스닥주식매도': '매도',
        'HTS코스닥주식매수': '매수',
        'HTS거래소주식매도': '매도',
        'HTS거래소주식매수': '매수',
        '채권만기상환': '채권만기상환',
        '배당세금추징': '배당세금추징',
        '제비용출금': '제비용출금',
    }
    
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        """
        한투 거래내역 파싱
        kwargs:
            - account_id: 계좌 별칭 (예: '한투_95963_01')
        """
        account_id = kwargs.get('account_id', 'unknown')
        events = []
        
        # 헤더 2행 건너뛰기
        for idx in range(1, len(df)):
            row = df.iloc[idx]
            try:
                event = self._parse_row(row, idx, account_id)
                if event:
                    events.append(event)
            except Exception as e:
                # 파싱 실패 행은 로그만 남기고 계속
                print(f"[{self.name}] {account_id} 행 {idx} 파싱 실패: {e}")
                continue
        
        return events
    
    def _parse_row(self, row, idx: int, account_id: str):
        날짜 = self.parse_date(row[0])
        if 날짜 is None:
            return None
        
        거래종류 = self.clean_text(row[2])
        event_type = self.TYPE_MAP.get(거래종류, 거래종류)
        
        종목명_raw = self.clean_text(row[3]) if len(row) > 3 else ''
        # 종목명 정규화 (회사 설정의 함수가 있으면 사용)
        종목명 = self._normalize_name(종목명_raw)
        수량 = self.safe_float(row[5]) if len(row) > 5 else 0
        단가 = self.safe_float(row[6]) if len(row) > 6 else 0
        거래금액 = self.safe_int(row[8]) if len(row) > 8 else 0
        정산금액 = self.safe_int(row[9]) if len(row) > 9 else 0
        수수료 = self.safe_int(row[10]) if len(row) > 10 else 0
        거래세 = self.safe_int(row[11]) if len(row) > 11 else 0
        
        # 원천징수세금
        세금 = 0
        for col_idx in (13, 14):
            if len(row) > col_idx:
                v = self.safe_int(row[col_idx])
                if v > 0 and v < 거래금액:
                    세금 = v
                    break
        
        event = RawEvent(
            날짜=날짜,
            event_type=event_type,
            원천=account_id,
            원천행번호=idx,
            # 거래금액/정산금액이 0이면 수량×단가로 계산 (타사이체입고 등)
            총액=(거래금액 if 거래금액 > 0 else
                  정산금액 if 정산금액 > 0 else
                  int(round(수량 * 단가))),
            종목명=종목명,
            수량=수량,
            단가=단가,
            수수료=수수료,
            거래세=거래세,
            세금=세금,
            적요원본=거래종류,
        )
        
        # 종목유형 추정
        if event_type in ('채권이자', '채권만기상환'):
            event.종목유형 = '채권'
            event.시장구분 = '회사채'  # 80895963-01의 채권은 회사채
        elif event_type == '배당금':
            event.종목유형 = '주식'
        elif event_type in ('공모주입고', '공모주출고'):
            event.종목유형 = '주식'
        elif event_type == '매수' or event_type == '매도':
            # 거래종류 원본에 코스닥/거래소 힌트
            if 'KOSDAQ' in 거래종류 or '코스닥' in 거래종류:
                event.종목유형 = '주식'
                event.시장구분 = '코스닥'
            elif 'KOSPI' in 거래종류 or '거래소' in 거래종류:
                event.종목유형 = '주식'
                event.시장구분 = '코스피'
            else:
                event.종목유형 = '주식'
        
        if 정산금액 and 정산금액 != 거래금액:
            event.extra['정산금액'] = 정산금액
        
        return event

    @staticmethod
    def _normalize_name(name: str) -> str:
        """종목명 정규화 - 분개장 형식과 일치시키기"""
        if not name:
            return ''
        import re
        name = name.strip()
        name = re.sub(r'^주식회사\s+', '', name)
        name = re.sub(r'^\(주\)\s*', '', name)
        name = re.sub(r'\s*\(주\)$', '', name)
        # "이마트24신종자본증권 37" → "이마트24신종자본증권37"
        name = re.sub(r'(\D)\s+(\d+)$', r'\1\2', name)
        return name
