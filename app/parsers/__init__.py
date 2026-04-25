"""
parsers/__init__.py
파서 레지스트리
"""
from .base import BaseParser, RawEvent
from .securities.hanto import HantoParser
from .securities.kiwoom import KiwoomParser
from .securities.kyobo import KyoboParser
from .securities.meritz import MeritzParser
from .banks.ibk import IBKBankParser
from .cards.ibk import IBKCardParser
from .cards.summary import CardSummaryParser


# 파서 매핑 (회사별로 다른 파서를 쓸 수도 있으므로 회사 설정에서 지정)
PARSERS = {
    'hanto': HantoParser,
    'kiwoom': KiwoomParser,
    'kyobo': KyoboParser,
    'meritz': MeritzParser,
    'ibk_bank': IBKBankParser,
    'ibk_card': IBKCardParser,
    'card_summary': CardSummaryParser,
}


def get_parser(name: str) -> BaseParser:
    if name not in PARSERS:
        raise ValueError(f"알 수 없는 파서: {name}")
    return PARSERS[name]()
