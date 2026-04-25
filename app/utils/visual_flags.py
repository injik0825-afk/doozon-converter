"""
utils/visual_flags.py
시각적 경고 표시 시스템 - 모든 회사 공통

분개 라인에 플래그를 달면 엑셀 출력 시 자동으로 색상이 입혀지고,
Streamlit UI에서도 색상으로 구분되어 표시됨.

사용 방법:
    entry.flags.append(Flag.NEW_SECURITY)
    entry.flags.append(Flag.MISSING_COST_BASIS)
"""
from enum import Enum


class Flag(str, Enum):
    """
    분개 라인 경고 플래그 (회사 무관)

    우선순위(높→낮): UNHANDLED > MISSING_COST_BASIS > NEW_SECURITY > INFERRED_ACCOUNT
    한 라인에 여러 플래그 가능, 화면에는 가장 높은 우선순위 색상 표시.
    """
    # 🔴 빨강 - 수동 검토 필수
    UNHANDLED = 'UNHANDLED'                  # 분개 자동생성 실패한 라인 (placeholder)
    BALANCE_MISMATCH = 'BALANCE_MISMATCH'    # 거래 단위 차/대변 불일치

    # 🟡 노랑 - 처분손익 계산 부정확 (취득가액 정보 부족)
    MISSING_COST_BASIS = 'MISSING_COST_BASIS'  # 매도했는데 보유수량 0 또는 평균단가 0

    # 🔵 파랑 - 신규 종목 (거래처 코드 미등록)
    NEW_SECURITY = 'NEW_SECURITY'            # 이번 달 처음 등장한 종목

    # 🟠 주황 - 추정/유추된 정보
    INFERRED_ACCOUNT = 'INFERRED_ACCOUNT'    # 계정과목을 비고/적요로 유추함
    INFERRED_PARTNER = 'INFERRED_PARTNER'    # 거래처를 추정함 (이체 상대 등)

    # ⚪ 회색 - 참고용
    AUTO_GENERATED = 'AUTO_GENERATED'        # 월말평가 등 자동 생성 (오류 아님, 단순 표시)


# 색상 우선순위 (위가 우선)
PRIORITY = [
    Flag.UNHANDLED,
    Flag.BALANCE_MISMATCH,
    Flag.MISSING_COST_BASIS,
    Flag.NEW_SECURITY,
    Flag.INFERRED_ACCOUNT,
    Flag.INFERRED_PARTNER,
    Flag.AUTO_GENERATED,
]


# 엑셀 셀 채움 색상 (ARGB 또는 RGB hex)
EXCEL_FILL_COLOR = {
    Flag.UNHANDLED:           'FFC7CE',  # 연빨강
    Flag.BALANCE_MISMATCH:    'FFC7CE',  # 연빨강
    Flag.MISSING_COST_BASIS:  'FFEB9C',  # 노랑
    Flag.NEW_SECURITY:        'BDD7EE',  # 파랑
    Flag.INFERRED_ACCOUNT:    'FFD966',  # 주황(연한)
    Flag.INFERRED_PARTNER:    'FCE4D6',  # 살구
    Flag.AUTO_GENERATED:      'EDEDED',  # 연회색
}

# Streamlit DataFrame 표시용 (CSS color)
ST_BG_COLOR = {
    Flag.UNHANDLED:           '#FFC7CE',
    Flag.BALANCE_MISMATCH:    '#FFC7CE',
    Flag.MISSING_COST_BASIS:  '#FFEB9C',
    Flag.NEW_SECURITY:        '#BDD7EE',
    Flag.INFERRED_ACCOUNT:    '#FFD966',
    Flag.INFERRED_PARTNER:    '#FCE4D6',
    Flag.AUTO_GENERATED:      '#EDEDED',
}

# 사람이 읽을 라벨 (UI용)
LABEL_KO = {
    Flag.UNHANDLED:           '🔴 자동분개 실패 (수동입력)',
    Flag.BALANCE_MISMATCH:    '🔴 차/대변 불일치',
    Flag.MISSING_COST_BASIS:  '🟡 취득가액 없음 (처분손익 0)',
    Flag.NEW_SECURITY:        '🔵 신규 종목 (거래처코드 미등록)',
    Flag.INFERRED_ACCOUNT:    '🟠 계정과목 추정',
    Flag.INFERRED_PARTNER:    '🟠 거래처 추정',
    Flag.AUTO_GENERATED:      '⚪ 자동생성 (월말평가 등)',
}


def top_flag(flags) -> Flag | None:
    """플래그 리스트에서 가장 우선순위 높은 플래그 반환"""
    if not flags:
        return None
    flag_set = {f if isinstance(f, Flag) else Flag(f) for f in flags}
    for p in PRIORITY:
        if p in flag_set:
            return p
    return None


def excel_fill_for(flags) -> str | None:
    """플래그 → 엑셀 색상 hex (없으면 None)"""
    f = top_flag(flags)
    return EXCEL_FILL_COLOR.get(f) if f else None


def st_bg_for(flags) -> str | None:
    """플래그 → Streamlit 배경색 (없으면 None)"""
    f = top_flag(flags)
    return ST_BG_COLOR.get(f) if f else None


def labels_for(flags) -> list[str]:
    """플래그 리스트 → 사람이 읽을 라벨"""
    out = []
    seen = set()
    for f in flags:
        f_enum = f if isinstance(f, Flag) else Flag(f)
        if f_enum in seen:
            continue
        seen.add(f_enum)
        out.append(LABEL_KO.get(f_enum, str(f_enum.value)))
    return out
