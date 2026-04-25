"""
companies/base.py
회사별 설정 베이스 클래스

각 회사는 이 클래스를 상속받아 다음을 정의:
- 계정과목 코드표 (ACCOUNTS)
- 거래처명 규칙
- 은행/증권사/카드 매핑
- 특수 분개 규칙 (예: 평가상계 방식)
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Dict, List, Optional


@dataclass
class BankAccount:
    """은행 계좌 정보"""
    별칭: str           # 앱 내 식별자 (예: "IBK_메인")
    계좌번호: str        # 예: "208-147831-04-011"
    은행명: str          # 예: "IBK기업은행"
    거래처명: str        # 분개장에 쓸 이름 (예: "기업은행 208-147831-04-011")


@dataclass
class SecuritiesAccount:
    """증권사 계좌 정보"""
    별칭: str           # 앱 내 식별자 (예: "키움_6340")
    계좌번호: str        # 예: "2770-6340"
    증권사: str          # 예: "키움증권"
    거래처명: str        # 분개장에 쓸 이름 (예: "키움증권 2770-6340")
    시트명패턴: str      # 엑셀 시트명 검색 패턴 (예: "키움(2770-6340)")


@dataclass
class Card:
    """카드 정보"""
    별칭: str            # 앱 내 식별자
    카드사: str          # 예: "IBK기업카드", "우리카드"
    카드번호끝자리: str  # 예: "5318"
    거래처명: str        # 분개장용 (예: "기업카드#5318")


class CompanyConfig(ABC):
    """회사별 설정 추상 클래스"""
    
    # 기본 정보
    회사명: str = ""
    사업연도: int = 0  # 예: 2026
    
    # ========================================
    # 계정과목 코드표 (회사별로 다름)
    # ========================================
    ACCOUNTS: Dict[str, int] = {}
    
    # ========================================
    # 특수 거래처 (세무서, 공단, 임대인 등)
    # ========================================
    PARTNERS: Dict[str, str] = {}  # 별칭 → 실제 거래처명
    
    # ========================================
    # 자금 계좌들
    # ========================================
    BANK_ACCOUNTS: List[BankAccount] = []
    SECURITIES_ACCOUNTS: List[SecuritiesAccount] = []
    CARDS: List[Card] = []
    
    # ========================================
    # 특수 처리 옵션
    # ========================================
    평가상계_방식: str = 'normal'  # 'normal' / 'signed' (음수로 되돌림)
    
    # ========================================
    # 헬퍼 메소드
    # ========================================
    def get_account_code(self, 계정과목: str) -> int:
        """계정과목명 → 코드"""
        if 계정과목 not in self.ACCOUNTS:
            raise KeyError(
                f"[{self.회사명}] 계정과목 코드 없음: '{계정과목}'")
        return self.ACCOUNTS[계정과목]
    
    def get_bank_by_account(self, 계좌번호: str) -> Optional[BankAccount]:
        for b in self.BANK_ACCOUNTS:
            if b.계좌번호 == 계좌번호:
                return b
        return None
    
    def get_securities_by_pattern(self, 시트명: str) -> Optional[SecuritiesAccount]:
        for s in self.SECURITIES_ACCOUNTS:
            if s.시트명패턴 in 시트명 or 시트명 in s.시트명패턴:
                return s
        return None
    
    def get_card_by_number(self, 끝자리: str) -> Optional[Card]:
        for c in self.CARDS:
            if c.카드번호끝자리 == 끝자리:
                return c
        return None
    
    @abstractmethod
    def 분개규칙_평가상계(self, 포지션_전월말_평가금액: int, 포지션_장부가: int,
                       종목_거래처명: str, 월: int, 일: int) -> List[dict]:
        """
        월초 평가상계 전표 생성 규칙
        회사마다 방식이 다를 수 있음 (두리는 음수 사용, 리버사이드는 정상 차/대변)
        Returns: 분개 라인들의 dict 리스트
        """
        pass
