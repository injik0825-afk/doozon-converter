"""
core/month_end.py
월말평가 + 월초 평가상계 분개 생성

[월초 평가상계]
전월말에 평가손익이 있었다면, 월초에 되돌리는 전표를 먼저 생성.
회사별로 방식이 다름:
  - 'signed' (두리인베스트먼트): 음수 대체전표 사용
  - 'normal' (리버사이드파트너스 등): 정상 차/대변

[월말평가]
포트폴리오의 모든 종목에 대해 현재가 × 수량 - 장부가 = 평가손익 계산.

[자동 플래그]
월초/월말 자동 생성 분개는 모두 AUTO_GENERATED 플래그 부여 (회색).
신규 종목이면 NEW_SECURITY 플래그 추가.
"""
from datetime import date
from typing import Dict
from calendar import monthrange

from .journal import JournalBook, Transaction
from .portfolio import Portfolio
from ..companies.base import CompanyConfig
from ..utils.visual_flags import Flag


class MonthEndProcessor:
    """월초 평가상계 + 월말평가 처리"""

    def __init__(self, company: CompanyConfig, portfolio: Portfolio, book: JournalBook):
        self.company = company
        self.portfolio = portfolio
        self.book = book

    def generate_opening_reversal(self, year: int, month: int):
        """월초 평가상계 (전월 평가손익 되돌림)"""
        if self.company.평가상계_방식 == 'signed':
            self._generate_signed_reversal(year, month)
        else:
            self._generate_normal_reversal(year, month)

    def _generate_signed_reversal(self, year: int, month: int):
        """두리인베스트먼트 방식: 같은 구분에 음수/양수 대체전표"""
        적요 = f"{month:02d}월말 평가상계"

        for 종목명, pos in self.portfolio.positions.items():
            # 학습 모드: 수량 0이어도 평가손익이 있으면 진행 (장부가 모름)
            if pos.전월말_평가손익 == 0:
                continue

            tx = Transaction(date(year, month, 1), '평가상계', '월초자동', -1)
            평가손익 = pos.전월말_평가손익

            # 신규 종목이면 거래처 등록 안되어 있을 수 있음
            flags = [Flag.AUTO_GENERATED]
            if pos.is_new and not pos.거래처코드:
                flags.append(Flag.NEW_SECURITY)

            if 평가손익 > 0:
                # 전월 평가이익 → 되돌림
                tx.add('차변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', -평가손익,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
                tx.add('차변', self.company.get_account_code('단기매매증권평가이익'),
                       '단기매매증권평가이익', 평가손익,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
            else:
                # 전월 평가손실 → 되돌림
                손실 = -평가손익
                tx.add('대변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', -손실,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
                tx.add('대변', self.company.get_account_code('단기매매증권평가손실'),
                       '단기매매증권평가손실', 손실,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)

            self.book.add_transaction(tx)

    def _generate_normal_reversal(self, year: int, month: int):
        """일반 방식 (리버사이드파트너스 스타일)"""
        적요 = f"{month:02d}월말 평가상계"

        for 종목명, pos in self.portfolio.positions.items():
            if pos.전월말_평가손익 == 0:
                continue

            tx = Transaction(date(year, month, 1), '평가상계', '월초자동', -1)
            평가손익 = pos.전월말_평가손익

            flags = [Flag.AUTO_GENERATED]
            if pos.is_new and not pos.거래처코드:
                flags.append(Flag.NEW_SECURITY)

            if 평가손익 > 0:
                tx.add('차변', self.company.get_account_code('단기매매증권평가이익'),
                       '단기매매증권평가이익', 평가손익,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
                tx.add('대변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', 평가손익,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
            else:
                손실 = -평가손익
                tx.add('차변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', 손실,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)
                tx.add('대변', self.company.get_account_code('단기매매증권평가손실'),
                       '단기매매증권평가손실', 손실,
                       거래처명=pos.거래처명, 거래처코드=pos.거래처코드,
                       적요=적요, flags=flags)

            self.book.add_transaction(tx)

    def generate_month_end_valuation(self, year: int, month: int,
                                      current_prices: Dict[str, float]):
        """월말평가 분개 생성"""
        last_day = monthrange(year, month)[1]
        평가일 = date(year, month, last_day)

        results = self.portfolio.calculate_month_end_valuation(current_prices)

        for r in results:
            평가손익 = r['평가손익']
            if 평가손익 == 0:
                continue

            수량 = r['수량']
            장부가 = r['장부가액']
            평가금액 = r['평가금액']
            단가 = r['현재가']
            거래처 = r['거래처명']
            거래처코드 = r.get('거래처코드', '')

            적요 = f"월말평가{'이익' if 평가손익 > 0 else '손실'}({int(수량):,}주*@{단가:,.0f})"
            적요_풀 = f"{적요}_{r['종목명']}_원금:{장부가:,}원"

            flags = [Flag.AUTO_GENERATED]
            if r.get('is_new'):
                flags.append(Flag.NEW_SECURITY)

            tx = Transaction(평가일, '월말평가', '월말자동', -1)

            if 평가손익 > 0:
                tx.add('차변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', 평가손익,
                       거래처명=거래처, 거래처코드=거래처코드,
                       적요=적요_풀, flags=flags)
                tx.add('대변', self.company.get_account_code('단기매매증권평가이익'),
                       '단기매매증권평가이익', 평가손익,
                       거래처명=거래처, 거래처코드=거래처코드,
                       적요=적요, flags=flags)
            else:
                손실 = -평가손익
                tx.add('차변', self.company.get_account_code('단기매매증권평가손실'),
                       '단기매매증권평가손실', 손실,
                       거래처명=거래처, 거래처코드=거래처코드,
                       적요=적요, flags=flags)
                tx.add('대변', self.company.get_account_code('단기매매증권'),
                       '단기매매증권', 손실,
                       거래처명=거래처, 거래처코드=거래처코드,
                       적요=적요_풀, flags=flags)

            self.book.add_transaction(tx)

            # 다음 달 평가상계용으로 저장
            pos = self.portfolio.positions[r['종목명']]
            pos.전월말_평가금액 = 평가금액
            pos.전월말_평가손익 = 평가손익
