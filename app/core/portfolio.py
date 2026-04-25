"""
core/portfolio.py
포트폴리오 관리 - 종목별 보유수량과 평균취득단가 추적
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import pandas as pd


@dataclass
class Position:
    """한 종목의 포지션"""
    종목명: str
    종목유형: str              # '주식' / '채권' / '펀드'
    시장구분: str = ''         # '코스피' / '코스닥' / '회사채' 등
    수량: float = 0
    평균단가: float = 0       # 이동평균법 기준
    장부가액: int = 0         # 수량 × 평균단가 (반올림)

    # 전월말 평가 정보 (평가상계 시 사용)
    전월말_평가금액: int = 0
    전월말_평가손익: int = 0

    # 신규 여부 (이번 결산기에 처음 등장)
    is_new: bool = False
    # 거래처코드 (신규면 비어있음 → 더존에 등록 필요)
    거래처코드: str = ''
    # 거래처명 (회사 설정으로 미리 계산해둘 수 있음 - format_partner 사용)
    _custom_거래처명: str = ''

    @property
    def 거래처명(self) -> str:
        """더존 거래처명 포맷 (커스텀 있으면 우선)"""
        if self._custom_거래처명:
            return self._custom_거래처명
        if self.종목유형 == '주식':
            return f"주식#{self.시장구분}#{self.종목명}"
        elif self.종목유형 == '채권':
            # 국민주택1종/국고채는 #구분# 없이 종목명만
            if self.종목명.startswith('국민주택') or self.종목명.startswith('국고채'):
                return f"채권#{self.종목명}"
            return f"채권#{self.시장구분 or '회사채'}#{self.종목명}"
        elif self.종목유형 == '펀드':
            return f"펀드#{self.종목명}"
        return self.종목명

    def set_거래처명(self, name: str):
        """회사 설정으로 계산된 거래처명을 미리 지정"""
        self._custom_거래처명 = name

    def buy(self, 수량: float, 단가: float) -> None:
        """매수 (이동평균법으로 단가 재계산)"""
        new_수량 = self.수량 + 수량
        new_장부가 = int(round(self.장부가액 + 수량 * 단가))
        new_평균 = new_장부가 / new_수량 if new_수량 > 0 else 0
        self.수량 = new_수량
        self.평균단가 = new_평균
        self.장부가액 = new_장부가

    def sell(self, 수량: float, 매도단가: float) -> tuple[int, int, bool]:
        """
        매도 (이동평균법)
        Returns: (처분대가, 처분손익, 취득가액부족여부)
            처분대가 = 수량 × 매도단가
            처분손익 = (매도단가 - 평균단가) × 수량
            취득가액부족여부 = 보유수량 0이거나 평균단가 0인 경우 True
        """
        취득가부족 = False
        # 보유 부족 케이스
        if self.수량 < 0.001 or self.평균단가 == 0:
            취득가부족 = True
            # 처분손익을 0으로 처리 (장부가 = 매도가)
            처분대가 = int(round(수량 * 매도단가))
            return 처분대가, 0, 취득가부족

        # 정상: 매도수량이 보유수량 초과면 보유분만 정리
        실제매도수량 = min(수량, self.수량)
        장부가_차감 = int(round(실제매도수량 * self.평균단가))
        처분대가 = int(round(수량 * 매도단가))
        # 보유 수량 한도 내 부분만 매도→처분손익 계산, 초과분은 매도가만큼 부족 처리
        if 수량 > self.수량 + 0.001:
            취득가부족 = True
            # 부족분은 장부가 = 매도가로 처리
            장부가_차감 += int(round((수량 - self.수량) * 매도단가))
        처분손익 = 처분대가 - 장부가_차감

        # 포지션 차감
        self.수량 -= 수량
        self.장부가액 -= 장부가_차감
        if self.수량 < 0.001:
            self.수량 = 0
            self.장부가액 = 0
            self.평균단가 = 0

        return 처분대가, 처분손익, 취득가부족


class Portfolio:
    """
    전체 포트폴리오 관리
    """
    def __init__(self):
        self.positions: Dict[str, Position] = {}  # key: 종목명

    def load_opening_positions(self, data: List[dict]):
        """
        기초 포지션 로드 (전월말 기준)
        학습 모드: 수량/단가 없이 전월말_평가손익만 있어도 동작
        """
        for item in data:
            name = item.get('종목명', '').strip()
            if not name:
                continue
            수량 = float(item.get('수량', 0) or 0)
            평균단가 = float(item.get('평균단가', 0) or 0)
            장부가 = int(item.get('장부가액', None) or round(수량 * 평균단가))
            전월말평가금액 = int(item.get('전월말_평가금액', 0) or 0)

            pos = Position(
                종목명=name,
                종목유형=item.get('종목유형', '주식'),
                시장구분=item.get('시장구분', ''),
                수량=수량,
                평균단가=평균단가,
                장부가액=장부가,
                전월말_평가금액=전월말평가금액,
                is_new=False,
                거래처코드=str(item.get('거래처코드', '') or ''),
            )

            # 전월말_평가손익 직접 제공 시 우선 사용 (학습 모드 지원)
            if item.get('전월말_평가손익') is not None:
                pos.전월말_평가손익 = int(item['전월말_평가손익'])
            else:
                pos.전월말_평가손익 = 전월말평가금액 - 장부가

            # 거래처명 커스텀 (분개장에서 역산한 형식)
            if item.get('거래처명'):
                pos.set_거래처명(item['거래처명'])

            self.positions[name] = pos

    def get_or_create(self, 종목명: str, 종목유형: str = '주식',
                       시장구분: str = '') -> Position:
        """포지션 조회, 없으면 신규로 생성 (is_new=True)"""
        if 종목명 not in self.positions:
            self.positions[종목명] = Position(
                종목명=종목명, 종목유형=종목유형, 시장구분=시장구분,
                is_new=True,  # 신규 종목 마킹
            )
        return self.positions[종목명]

    def is_new_security(self, 종목명: str) -> bool:
        """신규 종목 여부 (거래처 등록 안 된 상태)"""
        if 종목명 not in self.positions:
            return True
        pos = self.positions[종목명]
        return pos.is_new and not pos.거래처코드

    def has_cost_basis(self, 종목명: str) -> bool:
        """취득가액 정보 있는지 (매도 가능한지)"""
        if 종목명 not in self.positions:
            return False
        pos = self.positions[종목명]
        return pos.수량 > 0 and pos.평균단가 > 0

    def buy(self, 종목명: str, 수량: float, 단가: float,
             종목유형: str = '주식', 시장구분: str = '') -> tuple[Position, bool]:
        """
        매수 처리
        Returns: (포지션, is_new_flag)
            is_new_flag: 이번에 처음 보는 종목이면 True
        """
        is_new = 종목명 not in self.positions
        pos = self.get_or_create(종목명, 종목유형, 시장구분)
        # 유형/시장 보강
        if not pos.시장구분 and 시장구분:
            pos.시장구분 = 시장구분
        if not pos.종목유형 and 종목유형:
            pos.종목유형 = 종목유형
        pos.buy(수량, 단가)
        return pos, (is_new or pos.is_new)

    def sell(self, 종목명: str, 수량: float, 매도단가: float
              ) -> tuple[int, int, Position, bool, bool]:
        """
        매도 처리
        Returns: (처분대가, 처분손익, 포지션, 취득가액부족여부, 신규종목여부)
        """
        # 보유 안 한 종목 매도 시도 (드물지만 가능)
        if 종목명 not in self.positions:
            # 임시 포지션 생성 (취득가액 없음 상태)
            pos = self.get_or_create(종목명, '주식', '')
            처분대가 = int(round(수량 * 매도단가))
            return 처분대가, 0, pos, True, True

        pos = self.positions[종목명]
        is_new = pos.is_new and not pos.거래처코드
        처분대가, 처분손익, 취득가부족 = pos.sell(수량, 매도단가)
        return 처분대가, 처분손익, pos, 취득가부족, is_new

    def snapshot(self) -> pd.DataFrame:
        """현재 포트폴리오 스냅샷"""
        rows = []
        for pos in self.positions.values():
            if pos.수량 <= 0:
                continue
            rows.append({
                '종목명': pos.종목명,
                '종목유형': pos.종목유형,
                '시장구분': pos.시장구분,
                '수량': pos.수량,
                '평균단가': round(pos.평균단가, 2),
                '장부가액': pos.장부가액,
                '거래처코드': pos.거래처코드,
                '신규': pos.is_new,
            })
        return pd.DataFrame(rows)

    def calculate_month_end_valuation(
        self, current_prices: Dict[str, float]
    ) -> List[dict]:
        """
        월말평가 계산
        current_prices: {종목명: 현재단가}
        """
        results = []
        for 종목명, pos in self.positions.items():
            if pos.수량 <= 0:
                continue
            if 종목명 not in current_prices:
                continue
            현재가 = current_prices[종목명]
            평가금액 = int(round(pos.수량 * 현재가))
            평가손익 = 평가금액 - pos.장부가액
            results.append({
                '종목명': 종목명,
                '종목유형': pos.종목유형,
                '시장구분': pos.시장구분,
                '수량': pos.수량,
                '장부가액': pos.장부가액,
                '현재가': 현재가,
                '평가금액': 평가금액,
                '평가손익': 평가손익,
                '거래처명': pos.거래처명,
                '거래처코드': pos.거래처코드,
                'is_new': pos.is_new and not pos.거래처코드,
            })
        return results
