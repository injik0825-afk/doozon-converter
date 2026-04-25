"""
companies/__init__.py
회사 레지스트리 - 새 회사 추가 시 여기에 등록
"""
from typing import Dict, Type
from .base import CompanyConfig
from .duri_investment import DuriInvestmentConfig
# from .riverside_partners import RiversidePartnersConfig  # 기존 앱의 리버사이드 설정


# 회사 코드 → 설정 클래스 매핑
REGISTRY: Dict[str, Type[CompanyConfig]] = {
    'duri': DuriInvestmentConfig,
    # 'riverside': RiversidePartnersConfig,
}


def get_company_config(code: str) -> CompanyConfig:
    """회사 코드로 설정 인스턴스 반환"""
    if code not in REGISTRY:
        raise ValueError(
            f"등록되지 않은 회사 코드: {code}. 사용 가능: {list(REGISTRY.keys())}"
        )
    return REGISTRY[code]()


def list_companies() -> Dict[str, str]:
    """등록된 회사 목록 (code → 회사명)"""
    return {code: cls.회사명 for code, cls in REGISTRY.items()}
