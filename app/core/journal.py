"""
core/journal.py
분개 엔트리 모델 - 모든 거래의 최종 출력 형태

각 분개 라인은 시각적 경고 플래그를 가질 수 있어,
엑셀/UI에서 색상으로 구분되어 표시된다 (utils.visual_flags 참조).
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Optional, List
import pandas as pd

from ..utils.visual_flags import Flag, top_flag, labels_for


@dataclass
class JournalEntry:
    """
    하나의 분개 라인 (차변 또는 대변 한 줄)

    더존 위하고 일반전표 업로드 양식:
    1.월 | 2.일 | 3.구분 | 4.계정과목코드 | 5.계정과목명 |
    6.거래처코드 | 7.거래처명 | 8.적요명 | 9.차변(출금) | 10.대변(입금)
    """
    월: int
    일: int
    구분: str              # '차변' 또는 '대변'
    계정코드: int
    계정과목: str
    금액: int
    거래처명: str = ''
    거래처코드: str = ''
    적요: str = ''
    전표그룹: str = ''      # 같은 전표에 속하는 라인들의 식별자
    flags: List[Flag] = field(default_factory=list)  # 시각적 경고 플래그
    메모: str = ''          # 사람이 보는 부가 설명

    def __post_init__(self):
        if self.구분 not in ('차변', '대변'):
            raise ValueError(f"구분은 '차변' 또는 '대변'이어야 합니다: {self.구분}")
        if self.금액 < 0:
            raise ValueError(f"금액은 양수여야 합니다. 음수인 경우 구분을 반대로: {self.금액}")

    @property
    def 차변금액(self) -> int:
        return self.금액 if self.구분 == '차변' else 0

    @property
    def 대변금액(self) -> int:
        return self.금액 if self.구분 == '대변' else 0

    def add_flag(self, flag: Flag):
        """플래그 중복 없이 추가"""
        if flag not in self.flags:
            self.flags.append(flag)

    @property
    def top_flag(self) -> Optional[Flag]:
        return top_flag(self.flags)

    @property
    def flag_labels(self) -> List[str]:
        return labels_for(self.flags)

    def to_dict(self, include_flags: bool = False) -> dict:
        """더존 업로드 양식 dict로 변환"""
        d = {
            '1.월': self.월,
            '2.일': self.일,
            '3.구분': self.구분,
            '4.계정과목코드': self.계정코드,
            '5.계정과목명': self.계정과목,
            '6.거래처코드': self.거래처코드,
            '7.거래처명': self.거래처명,
            '8.적요명': self.적요,
            '9.차변(출금)': self.차변금액 if self.구분 == '차변' else '',
            '10.대변(입금)': self.대변금액 if self.구분 == '대변' else '',
        }
        if include_flags:
            d['_경고'] = ', '.join(self.flag_labels) if self.flags else ''
            d['_메모'] = self.메모
        return d


@dataclass
class Transaction:
    """
    하나의 거래 (여러 분개 라인의 묶음)
    """
    날짜: date
    거래유형: str
    원천: str
    원천행번호: int = -1
    entries: List[JournalEntry] = field(default_factory=list)
    메모: str = ''

    def add(self, 구분: str, 계정코드: int, 계정과목: str, 금액: int,
            거래처명: str = '', 적요: str = '',
            flags: Optional[List[Flag]] = None,
            메모: str = '',
            거래처코드: str = ''):
        """
        분개 라인 추가 (플래그 함께 전달 가능)
        Returns: 생성된 JournalEntry (또는 0원이면 None)
        """
        if 금액 == 0:
            return None
        if 금액 < 0:
            구분 = '대변' if 구분 == '차변' else '차변'
            금액 = -금액
        entry = JournalEntry(
            월=self.날짜.month, 일=self.날짜.day, 구분=구분,
            계정코드=계정코드, 계정과목=계정과목, 금액=금액,
            거래처명=거래처명, 거래처코드=거래처코드, 적요=적요,
            전표그룹=f"{self.날짜}_{self.거래유형}_{self.원천행번호}",
            flags=list(flags) if flags else [],
            메모=메모,
        )
        self.entries.append(entry)
        return entry

    def add_flag_to_all(self, flag: Flag):
        """이 거래의 모든 라인에 플래그 추가"""
        for e in self.entries:
            e.add_flag(flag)

    def validate(self) -> tuple[bool, str]:
        차 = sum(e.차변금액 for e in self.entries)
        대 = sum(e.대변금액 for e in self.entries)
        if 차 == 대 and 차 > 0:
            return True, "OK"
        return False, f"차변 {차:,} ≠ 대변 {대:,} (차이 {차-대:,}) | {self.거래유형} {self.날짜}"

    def __repr__(self):
        return f"Transaction({self.날짜}, {self.거래유형}, {len(self.entries)} lines)"


class JournalBook:
    """
    분개장 전체 관리
    """
    def __init__(self):
        self.transactions: List[Transaction] = []
        self.warnings: List[str] = []

    def add_transaction(self, tx: Transaction):
        ok, msg = tx.validate()
        if not ok:
            self.warnings.append(f"⚠️ {msg}")
            tx.add_flag_to_all(Flag.BALANCE_MISMATCH)
        self.transactions.append(tx)

    def all_entries(self) -> List[JournalEntry]:
        entries = []
        for tx in self.transactions:
            entries.extend(tx.entries)
        entries.sort(key=lambda e: (e.월, e.일))
        return entries

    def to_dataframe(self, include_flags: bool = False) -> pd.DataFrame:
        rows = [e.to_dict(include_flags=include_flags) for e in self.all_entries()]
        return pd.DataFrame(rows)

    def flag_counts(self) -> dict:
        """플래그별 라인 수 집계"""
        counts = {}
        for e in self.all_entries():
            for f in e.flags:
                counts[f] = counts.get(f, 0) + 1
        return counts

    def lines_with_flag(self, flag: Flag) -> List[JournalEntry]:
        return [e for e in self.all_entries() if flag in e.flags]

    def summary(self) -> dict:
        entries = self.all_entries()
        total_dr = sum(e.차변금액 for e in entries)
        total_cr = sum(e.대변금액 for e in entries)
        by_type = {}
        for tx in self.transactions:
            by_type[tx.거래유형] = by_type.get(tx.거래유형, 0) + 1
        flag_counts = self.flag_counts()
        return {
            '총 거래 수': len(self.transactions),
            '총 분개 라인': len(entries),
            '차변 합계': total_dr,
            '대변 합계': total_cr,
            '차/대변 일치': total_dr == total_cr,
            '거래 유형별 건수': by_type,
            '경고 수': len(self.warnings),
            '플래그별 라인수': {f.value: c for f, c in flag_counts.items()},
        }
