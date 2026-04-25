"""
parsers/base.py
거래내역 파서 베이스 클래스

각 파서는 특정 증권사/은행/카드의 거래내역 엑셀 시트를 읽고,
표준 '거래 이벤트'로 변환하는 역할을 담당.
회사별 계정 규칙은 여기서 몰라도 되고, Converter가 처리.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import date
from typing import List, Optional, Any
import pandas as pd


@dataclass
class RawEvent:
    """
    거래내역에서 추출한 원시 이벤트 (계정 규칙 적용 전)
    
    예:
    - RawEvent(date=4/15, event_type='주식매수', 종목='케이뱅크', 수량=500, 단가=8300, 수수료=0, 총액=4150000)
    - RawEvent(date=4/1, event_type='예탁금이자', 총액=1666, 세금=0)
    - RawEvent(date=4/10, event_type='채권이자입금', 종목='두산310-2', 총액=196312, 세금=30220)
    """
    날짜: date
    event_type: str              # 'buy', 'sell', '채권이자', '예탁금이자', '이체입금', 등
    원천: str                    # 데이터 출처 (예: '키움_6340')
    원천행번호: int = -1
    
    # 거래 공통
    총액: int = 0                 # 대표 금액 (세전 or 실입금, event_type에 따라 해석)
    
    # 증권 거래용 필드
    종목명: str = ''
    종목유형: str = ''            # '주식' / '채권' / '펀드'
    시장구분: str = ''
    수량: float = 0
    단가: float = 0
    수수료: int = 0
    세금: int = 0                 # 원천징수세금 (이자/배당)
    거래세: int = 0               # 증권거래세
    
    # 이체/결제용
    상대계좌: str = ''
    상대은행: str = ''
    적요원본: str = ''            # 원본 거래내용
    메모: str = ''                # 처리 힌트
    
    # 기타
    extra: dict = field(default_factory=dict)


class BaseParser(ABC):
    """파서 베이스 클래스"""
    
    name: str = "BaseParser"
    
    @abstractmethod
    def parse(self, df: pd.DataFrame, **kwargs) -> List[RawEvent]:
        """
        DataFrame을 받아 RawEvent 리스트 반환
        kwargs에는 계좌 정보 등이 전달됨
        """
        pass
    
    # 공통 헬퍼
    @staticmethod
    def safe_int(value: Any, default: int = 0) -> int:
        """안전한 정수 변환"""
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.replace(',', '').strip()
            if not value or value == '-':
                return default
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def safe_float(value: Any, default: float = 0.0) -> float:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.replace(',', '').strip()
            if not value or value == '-':
                return default
        try:
            return float(value)
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def parse_date(value: Any) -> Optional[date]:
        """다양한 형식의 날짜를 date 객체로"""
        if pd.isna(value):
            return None
        try:
            return pd.to_datetime(value).date()
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def clean_text(value: Any) -> str:
        """공백 정리"""
        if pd.isna(value):
            return ''
        return str(value).strip()
