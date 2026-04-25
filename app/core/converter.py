"""
core/converter.py
메인 변환 엔진

RawEvent 리스트를 받아 회사 설정에 따라 JournalBook으로 변환.
시각적 경고 플래그(노랑/파랑/주황 등)를 자동으로 부여한다.
"""
from typing import List, Optional
from datetime import date

from .journal import JournalBook, Transaction
from .portfolio import Portfolio
from ..parsers.base import RawEvent
from ..companies.base import CompanyConfig
from ..utils.visual_flags import Flag


class Converter:
    """거래 이벤트 → 분개장 변환기"""

    def __init__(self, company: CompanyConfig, portfolio: Optional[Portfolio] = None):
        self.company = company
        self.portfolio = portfolio or Portfolio()
        self.book = JournalBook()
        self.unhandled_events: List[RawEvent] = []

    def convert(self, events: List[RawEvent]) -> JournalBook:
        """
        이벤트 리스트를 분개장으로 변환
        매수/매도 순서가 처분이익에 영향을 주므로 날짜순 정렬 후 처리
        """
        events_sorted = sorted(events, key=lambda e: (e.날짜, e.원천, e.원천행번호))

        for event in events_sorted:
            try:
                self._process_event(event)
            except Exception as e:
                self.unhandled_events.append(event)
                self.book.warnings.append(
                    f"❌ 이벤트 처리 실패 [{event.날짜} {event.event_type} {event.종목명}]: {e}"
                )

        return self.book

    def _process_event(self, event: RawEvent):
        """이벤트 유형별 라우팅"""
        handler_map = {
            '예탁금이자': self._handle_예탁금이자,
            '채권이자': self._handle_채권이자,
            '채권이자입금': self._handle_채권이자,  # alias
            '배당금': self._handle_배당금,
            '매수': self._handle_매수,
            '매도': self._handle_매도,
            'KOSDAQ매도': self._handle_매도,        # alias
            'KOSDAQ매수': self._handle_매수,
            'HTS코스닥주식매도': self._handle_매도,
            'HTS코스닥주식매수': self._handle_매수,
            'HTS거래소주식매도': self._handle_매도,
            'HTS거래소주식매수': self._handle_매수,
            '채권만기상환': self._handle_채권만기상환,
            '타사대체입고': self._handle_공모주_타사입고,
            '타사이체입고': self._handle_공모주_타사입고,   # alias
            '타사대체출고': self._handle_공모주_타사출고,
            'HTS타사이체출고': self._handle_공모주_타사출고,
            'HTS타사이체입고': self._handle_공모주_타사입고,
            'HTS당사이체입고': self._handle_공모주_타사입고,
            'HTS당사이체출고': self._handle_공모주_타사출고,
            '공모주입고': self._handle_공모주입고,
            '공모주출고': self._handle_공모주_타사출고,
            '청약납입': self._handle_청약납입,
            'WTS추납대체청약': self._handle_청약납입,
            '이체입금': self._handle_이체입금,
            '이체출금': self._handle_이체출금,
            '당사이체입금': self._handle_당사이체입금,
            '당사이체출금': self._handle_당사이체출금,
            '대체입금': self._handle_대체입금,
            '대체출금': self._handle_대체출금,
            # 은행
            '국민연금': self._handle_국민연금,
            '건강보험': self._handle_건강보험,
            '고용보험': self._handle_고용보험,
            '산재보험': self._handle_산재보험,
            '소득세': self._handle_소득세,
            '지방소득세': self._handle_지방소득세,
            '국세납부': self._handle_국세납부,
            '지방세납부': self._handle_지방세납부,
            '은행이자': self._handle_은행이자,
            '이체수수료': self._handle_이체수수료,
            '증권사이체출금': self._handle_증권사이체출금,
            '증권사이체입금': self._handle_증권사이체입금,
            '자사이체출금': self._handle_자사이체출금,
            '자사이체입금': self._handle_자사이체입금,
            '카드결제_우리': self._handle_카드결제_우리,
            '카드결제_하나': self._handle_카드결제_하나,
            '카드결제_IBK': self._handle_카드결제_IBK,
            '카드결제_BC': self._handle_카드결제_BC,
            # 일임계좌 청약 관련
            '일임청약출금': self._handle_일임청약출금,
            '일임청약환급입금': self._handle_일임청약환급입금,
            '에이치엔입금': self._handle_에이치엔입금,
            '청약납입_은행': self._handle_청약납입_은행,
            '미지급금결제': self._handle_미지급금결제,
            # 급여/퇴직
            '급여지급': self._handle_급여지급,
            '퇴직연금DC': self._handle_퇴직연금DC,
            '퇴직연금부담금': self._handle_퇴직연금부담금,
            '퇴직연금수수료': self._handle_퇴직연금수수료,
            # 기타 은행 지출
            '임대료': self._handle_임대료,
            '통신비': self._handle_통신비,
            '통신비KT': self._handle_통신비,
            '접대비_기타': self._handle_접대비_기타,
            # 카드
            '카드사용': self._handle_카드사용,
            '카드승인': self._handle_카드승인,
        }

        handler = handler_map.get(event.event_type)
        if handler is None:
            self.unhandled_events.append(event)
            # 기본 자리표시자 분개 (모든 라인에 UNHANDLED 플래그)
            self._make_unhandled_placeholder(event)
            return

        tx = handler(event)
        if tx is not None:
            self.book.add_transaction(tx)

    def _make_unhandled_placeholder(self, event: RawEvent):
        """미처리 이벤트도 분개장 시각화에서 보이도록 자리표시자 생성"""
        if event.총액 == 0:
            return  # 0원은 무시
        tx = Transaction(event.날짜, f'미처리_{event.event_type}',
                          event.원천, event.원천행번호)
        # 차/대변 의도를 알 수 없으므로 양쪽에 동일하게 미수금/잡손익 자리표시자
        # 실제로는 사용자가 수정해야 함
        try:
            tx.add('차변', 999, '?미정?', event.총액,
                   거래처명=event.원천,
                   적요=f"[자동분개실패] {event.적요원본}",
                   flags=[Flag.UNHANDLED],
                   메모=f"파서 출력: {event.event_type} | 종목: {event.종목명}")
            tx.add('대변', 999, '?미정?', event.총액,
                   거래처명=event.원천,
                   적요=f"[자동분개실패] {event.적요원본}",
                   flags=[Flag.UNHANDLED])
            self.book.add_transaction(tx)
        except Exception:
            pass

    # ========================================
    # 증권사 거래 핸들러
    # ========================================
    def _handle_예탁금이자(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '예탁금이자', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        금액 = event.총액
        tx.add('차변', self._code('예치금'), '예치금', 금액,
               거래처명=acct, 적요='예탁금이용료(이자)입금')
        tx.add('대변', self._code('이자수익(금융)'), '이자수익(금융)', 금액,
               거래처명=acct, 적요='예탁금이용료(이자)입금')
        return tx

    def _handle_채권이자(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '채권이자', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)

        세전 = event.총액
        세금 = event.세금
        세후 = event.extra.get('정산금액', 세전 - 세금)

        # 국세/지방세 9:1 분리
        if 세금 > 0:
            국세 = round(세금 * 10 / 11)
            지방세 = 세금 - 국세
        else:
            국세 = 지방세 = 0

        # 신규 종목 체크 (채권 거래처도 신규일 수 있음)
        flags = []
        if event.종목명 and event.종목명 not in self.portfolio.positions:
            flags.append(Flag.NEW_SECURITY)

        채권거래처 = self._format_security_partner(event)
        적요 = f"{event.종목명} 채권이자입금" if event.종목명 else '채권이자입금'

        tx.add('차변', self._code('예치금'), '예치금', 세후,
               거래처명=acct, 적요=적요)
        if 국세 > 0:
            tx.add('차변', self._code('선납세금'), '선납세금', 국세,
                   거래처명=self.company.PARTNERS.get('국세', '세무서'),
                   적요=f"{event.종목명} 선납법인세")
        if 지방세 > 0:
            tx.add('차변', self._code('선납세금'), '선납세금', 지방세,
                   거래처명=self.company.PARTNERS.get('지방세', '구청'),
                   적요=f"{event.종목명} 선납법인세 지방세분")
        tx.add('대변', self._code('이자수익(금융)'), '이자수익(금융)', 세전,
               거래처명=채권거래처, 적요=적요,
               flags=flags,
               메모='신규 채권 - 거래처 등록 필요' if Flag.NEW_SECURITY in flags else '')
        return tx

    def _handle_배당금(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '배당금', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        금액 = event.총액

        flags = []
        if event.종목명 and event.종목명 not in self.portfolio.positions:
            flags.append(Flag.NEW_SECURITY)

        tx.add('차변', self._code('예치금'), '예치금', 금액,
               거래처명=acct, 적요=f"{event.종목명} 배당금입금")
        tx.add('대변', self._code('분배금수익'), '분배금수익', 금액,
               거래처명=acct, 적요=f"{event.종목명} 배당금입금",
               flags=flags)
        return tx

    def _handle_매수(self, event: RawEvent) -> Transaction:
        """
        주식/채권 매수:
          차변 단기매매증권    (원금: 거래금액 우선, 없으면 수량×단가)
          차변 주식거래수수료  (수수료)
          대변 예치금          (정산금액)
        포트폴리오: buy(종목, 수량, 실효단가=원금/수량)
        """
        tx = Transaction(event.날짜, '매수', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        종목거래처 = self._format_security_partner(event)

        # 원금: 거래금액(총액)이 있으면 우선 사용 (교보 채권 ×10 문제 회피)
        if event.총액 > 0:
            원금 = event.총액
        else:
            원금 = int(round(event.수량 * event.단가))
        수수료 = event.수수료
        정산 = event.extra.get('정산금액', 원금 + 수수료)

        # 적요용 표시단가 (교보 채권에서 파일단가는 ×10이라 거래처명용)
        표시단가 = event.단가

        # 포트폴리오: 실효단가 = 원금 / 수량 (이동평균 정확하게)
        실효단가 = (원금 / event.수량) if event.수량 > 0 else event.단가
        pos, is_new = self.portfolio.buy(
            event.종목명, event.수량, 실효단가,
            종목유형=event.종목유형 or '주식',
            시장구분=event.시장구분,
        )
        # 회사 설정 거래처명 우선
        format_fn = getattr(self.company, 'format_partner', None)
        if format_fn and pos:
            종목거래처 = format_fn(pos.종목유형, event.종목명, pos.시장구분)
        종목거래처코드 = pos.거래처코드 if pos else ''

        flags = []
        메모 = ''
        if is_new:
            flags.append(Flag.NEW_SECURITY)
            메모 = '신규 종목 - 더존에 거래처 등록 필요'

        계좌접미 = getattr(self.company, 'ACCOUNT_SUFFIX_FOR_REMARK', {}).get(event.원천, event.원천)
        적요 = f"{event.종목명}({int(event.수량):,}주*@{표시단가:,.0f})매수#{계좌접미}"

        tx.add('차변', self._code('단기매매증권'), '단기매매증권', 원금,
               거래처명=종목거래처, 거래처코드=종목거래처코드,
               적요=적요, flags=flags, 메모=메모)
        if 수수료 > 0:
            tx.add('차변', self._code('주식거래수수료(판)'), '주식거래수수료(판)',
                   수수료, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요='매수수수료', flags=flags)
        tx.add('대변', self._code('예치금'), '예치금', 정산,
               거래처명=acct, 적요=적요)
        return tx

    def _handle_매도(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '매도', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)

        수량 = event.수량
        표시단가 = event.단가
        수수료 = event.수수료
        거래세 = event.거래세

        # 매도대금: 거래금액(총액) 우선 (교보 채권 ×10 방지)
        if event.총액 > 0:
            매도대금 = event.총액
        else:
            매도대금 = int(round(수량 * 표시단가))

        # 실효단가 = 매도대금 / 수량 (포트폴리오 처분이익 계산용)
        실효단가 = (매도대금 / 수량) if 수량 > 0 else 표시단가
        정산 = event.extra.get('정산금액', 매도대금 - 수수료 - 거래세)

        처분대가, 처분손익, pos, 취득가부족, is_new = self.portfolio.sell(
            event.종목명, 수량, 실효단가
        )
        장부가_차감 = 처분대가 - 처분손익

        if pos:
            if not pos.시장구분 and event.시장구분: pos.시장구분 = event.시장구분
            if not pos.종목유형 and event.종목유형: pos.종목유형 = event.종목유형

        format_fn = getattr(self.company, 'format_partner', None)
        if format_fn and pos:
            종목거래처 = format_fn(pos.종목유형, event.종목명, pos.시장구분)
        else:
            종목거래처 = self._format_security_partner(event)
        종목거래처코드 = pos.거래처코드 if pos else ''

        flags = []
        메모 = ''
        if 취득가부족:
            flags.append(Flag.MISSING_COST_BASIS)
            메모 = (f'취득가액 정보 없음 - 처분손익 0원으로 처리됨. '
                  f'분개 검토 후 수동 보정 필요')
            self.book.warnings.append(
                f"⚠️ {event.종목명} 매도 시 취득가액 부족 ({event.날짜})")
        if is_new:
            flags.append(Flag.NEW_SECURITY)
            if 메모:
                메모 += ' | 신규 종목 - 거래처 등록 필요'
            else:
                메모 = '신규 종목 - 거래처 등록 필요'

        계좌접미 = getattr(self.company, 'ACCOUNT_SUFFIX_FOR_REMARK', {}).get(event.원천, event.원천)
        적요 = f"{event.종목명}({int(수량):,}주*@{표시단가:,.0f})매도#{계좌접미}"

        # 차변
        tx.add('차변', self._code('예치금'), '예치금', 정산,
               거래처명=acct, 적요=적요)
        if 수수료 > 0:
            tx.add('차변', self._code('주식거래수수료(판)'), '주식거래수수료(판)',
                   수수료, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요='매도수수료', flags=flags)
        if 거래세 > 0:
            tx.add('차변', self._code('세금과공과금(판)'), '세금과공과금(판)',
                   거래세, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요='증권거래세', flags=flags)

        # 대변 - 단기매매증권 차감
        if 장부가_차감 > 0:
            tx.add('대변', self._code('단기매매증권'), '단기매매증권', 장부가_차감,
                   거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=적요, flags=flags, 메모=메모)

        if 처분손익 > 0:
            tx.add('대변', self._code('단기매매증권처분이익'), '단기매매증권처분이익',
                   처분손익, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=f"{event.종목명}({int(수량):,}주*@{표시단가:,.0f})매도처분이익",
                   flags=flags)
        elif 처분손익 < 0:
            tx.add('차변', self._code('단기매매증권처분손실'), '단기매매증권처분손실',
                   -처분손익, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=f"{event.종목명}({int(수량):,}주*@{표시단가:,.0f})매도처분손실",
                   flags=flags)
        return tx

    # ========================================
    # 공모주 관련
    # ========================================
    def _handle_공모주_타사입고(self, event: RawEvent) -> Transaction:
        """
        타사대체입고 (공모주 배정 or 이관입고):
          차변 단기매매증권 / 대변 선급금 (청약했던 선급금 차감)
        """
        tx = Transaction(event.날짜, '공모주입고', event.원천, event.원천행번호)

        # 금액: 총액이 있으면 우선, 없으면 수량×단가
        금액 = event.총액
        if 금액 == 0 and event.수량 > 0 and event.단가 > 0:
            금액 = int(round(event.수량 * event.단가))
        if 금액 == 0:
            return None  # 정보 부족

        is_new = event.종목명 not in self.portfolio.positions
        flags = [Flag.NEW_SECURITY] if is_new else []
        메모 = '신규 종목 - 거래처 등록 필요' if is_new else ''

        pos, _ = self.portfolio.buy(event.종목명, event.수량, event.단가,
                                     종목유형=event.종목유형 or '주식',
                                     시장구분=event.시장구분 or '코스닥')
        종목거래처 = pos.거래처명
        종목거래처코드 = pos.거래처코드

        # 적요 포맷 (분개장 형식: 종목명(N주*@단가)입고#계좌별칭)
        계좌별칭 = getattr(self.company, 'ACCOUNT_SUFFIX_FOR_REMARK', {}).get(event.원천, event.원천)
        if event.수량 > 0 and event.단가 > 0:
            적요 = f"{event.종목명}({int(event.수량):,}주*@{event.단가:,.0f})입고#{계좌별칭}"
        else:
            적요 = f"{event.종목명} 입고#{계좌별칭}"

        tx.add('차변', self._code('단기매매증권'), '단기매매증권', 금액,
               거래처명=종목거래처, 거래처코드=종목거래처코드,
               적요=적요, flags=flags, 메모=메모)
        tx.add('대변', self._code('선급금'), '선급금', 금액,
               거래처명=종목거래처, 거래처코드=종목거래처코드,
               적요=적요, flags=flags)
        return tx

    def _handle_공모주_타사출고(self, event: RawEvent) -> None:
        return None  # 상대 증권사에서 입고로 처리

    def _handle_공모주입고(self, event: RawEvent) -> Transaction:
        return self._handle_공모주_타사입고(event)

    # ========================================
    # 계좌간 이체
    # ========================================
    def _handle_이체입금(self, event: RawEvent) -> Transaction:
        """증권사 이체입금(은행→증권사): 예치금(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '이체입금', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        금액 = event.총액
        bank = self._infer_counterparty_bank()
        tx.add('차변', self._code('예치금'), '예치금', 금액,
               거래처명=acct, 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 금액,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_이체출금(self, event: RawEvent) -> Transaction:
        """증권사 이체출금(증권사→은행): 보통예금(차) / 예치금(대)"""
        tx = Transaction(event.날짜, '이체출금', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        금액 = event.총액
        bank = self._infer_counterparty_bank()
        tx.add('차변', self._code('보통예금'), '보통예금', 금액,
               거래처명=bank, 적요=event.적요원본)
        tx.add('대변', self._code('예치금'), '예치금', 금액,
               거래처명=acct, 적요=event.적요원본)
        return tx

    def _handle_당사이체입금(self, event: RawEvent) -> None:
        """당사이체입금: 출금쪽에서 처리 → skip (이중 방지)"""
        return None

    def _handle_당사이체출금(self, event: RawEvent) -> Transaction:
        """당사이체출금: 예치금↔예치금 (같은 증권사 내 계좌간 이체)"""
        tx = Transaction(event.날짜, '당사이체출금', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)
        금액 = event.총액
        tx.add('차변', self._code('예치금'), '예치금', 금액,
               거래처명=acct + '(대체처)', 적요='HTS당사이체출금',
               flags=[Flag.INFERRED_PARTNER],
               메모='대체처 계좌 확인 필요')
        tx.add('대변', self._code('예치금'), '예치금', 금액,
               거래처명=acct, 적요='HTS당사이체출금')
        return tx

    def _handle_대체입금(self, event: RawEvent) -> None:
        return None  # skip

    def _handle_대체출금(self, event: RawEvent) -> Transaction:
        return self._handle_당사이체출금(event)

    # ========================================
    # 은행 핸들러
    # ========================================
    def _handle_4대보험_공통(self, event: RawEvent, 보험종류: str,
                            partner_key: str) -> Transaction:
        """
        4대보험 공통 핸들러
        회사 설정의 INSURANCE_DEDUCTION에서 직원공제분(예수금) 가져오고,
        나머지를 회사부담분(설정된 계정)에 할당.
        """
        tx = Transaction(event.날짜, 보험종류, event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        partner = self.company.PARTNERS.get(partner_key, '')

        # 회사 설정에서 직원공제분 가져오기 (없으면 50:50 추정)
        deduction = getattr(self.company, 'INSURANCE_DEDUCTION', {}).get(보험종류)
        company_acct_map = getattr(self.company, 'INSURANCE_COMPANY_ACCT', {})
        company_acct = company_acct_map.get(보험종류, '세금과공과금(판)')

        if deduction is None:
            # 설정 없음 → 절반 추정
            예수금 = 출금 // 2
            flags_dedu = [Flag.INFERRED_ACCOUNT]
            메모_dedu = '직원공제분/회사부담분 비율 50:50 추정 - 회사 설정 권장'
        else:
            예수금 = deduction
            flags_dedu = []
            메모_dedu = ''

        회사부담 = 출금 - 예수금

        if 예수금 > 0:
            tx.add('차변', self._code('예수금'), '예수금', 예수금,
                   거래처명=partner, 적요=event.적요원본,
                   flags=flags_dedu, 메모=메모_dedu)
        if 회사부담 > 0:
            tx.add('차변', self._code(company_acct), company_acct, 회사부담,
                   거래처명=partner, 적요=event.적요원본,
                   flags=flags_dedu)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_국민연금(self, event: RawEvent) -> Transaction:
        return self._handle_4대보험_공통(event, '국민연금', '국민연금')

    def _handle_건강보험(self, event: RawEvent) -> Transaction:
        return self._handle_4대보험_공통(event, '건강보험', '건강보험')

    def _handle_고용보험(self, event: RawEvent) -> Transaction:
        return self._handle_4대보험_공통(event, '고용보험', '고용보험')

    def _handle_산재보험(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '산재보험', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('보험료(판)'), '보험료(판)', 출금,
               거래처명='근로복지공단', 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_소득세(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '소득세', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('예수금'), '예수금', 출금,
               거래처명=self.company.PARTNERS.get('국세', '마포세무서'),
               적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_지방소득세(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '지방소득세', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('예수금'), '예수금', 출금,
               거래처명=self.company.PARTNERS.get('지방세', '마포구청'),
               적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_국세납부(self, event: RawEvent) -> Transaction:
        """국세 납부 (예수금 → 보통예금)"""
        tx = Transaction(event.날짜, '국세납부', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('예수금'), '예수금', 출금,
               거래처명=self.company.PARTNERS.get('국세', '마포세무서'),
               적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_지방세납부(self, event: RawEvent) -> Transaction:
        """지방세 납부"""
        tx = Transaction(event.날짜, '지방세납부', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('예수금'), '예수금', 출금,
               거래처명=self.company.PARTNERS.get('지방세', '마포구청'),
               적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_채권만기상환(self, event: RawEvent) -> Transaction:
        """
        채권 만기 상환:
          차변 예치금 (상환금액)
          대변 단기매매증권 (장부가액)
          대변/차변 처분이익/손실 (차액)
        """
        tx = Transaction(event.날짜, '채권만기상환', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)

        상환금액 = event.총액
        # 포트폴리오에서 만기 처리 (매도와 동일 로직, 단가는 상환금액 / 수량)
        if event.수량 > 0:
            상환단가 = 상환금액 / event.수량
        else:
            # 수량 미지정 시 보유 전량 상환으로 처리
            pos = self.portfolio.positions.get(event.종목명)
            if pos:
                event.수량 = pos.수량
                상환단가 = 상환금액 / event.수량 if event.수량 > 0 else 0
            else:
                상환단가 = 0

        처분대가, 처분손익, pos, 취득가부족, is_new = self.portfolio.sell(
            event.종목명, event.수량, 상환단가
        )
        장부가_차감 = 처분대가 - 처분손익

        종목거래처 = pos.거래처명 if pos else self._format_security_partner(event)
        종목거래처코드 = pos.거래처코드 if pos else ''

        flags = []
        메모 = ''
        if 취득가부족:
            flags.append(Flag.MISSING_COST_BASIS)
            메모 = '취득가액 정보 없음 - 처분손익 0원 처리. 정확한 보정 필요'
        if is_new:
            flags.append(Flag.NEW_SECURITY)

        적요 = f"{event.종목명} 만기상환"
        tx.add('차변', self._code('예치금'), '예치금', 상환금액,
               거래처명=acct, 적요=적요)
        if 장부가_차감 > 0:
            tx.add('대변', self._code('단기매매증권'), '단기매매증권', 장부가_차감,
                   거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=적요, flags=flags, 메모=메모)
        if 처분손익 > 0:
            tx.add('대변', self._code('단기매매증권처분이익'), '단기매매증권처분이익',
                   처분손익, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=f"{event.종목명} 만기상환처분이익", flags=flags)
        elif 처분손익 < 0:
            tx.add('차변', self._code('단기매매증권처분손실'), '단기매매증권처분손실',
                   -처분손익, 거래처명=종목거래처, 거래처코드=종목거래처코드,
                   적요=f"{event.종목명} 만기상환처분손실", flags=flags)
        return tx

    def _handle_청약납입(self, event: RawEvent) -> Transaction:
        """
        공모주 청약 납입:
          차변 선급금 (청약금)
          차변 주식거래수수료 (청약수수료)
          대변 예치금 (정산금액)
        """
        tx = Transaction(event.날짜, '청약납입', event.원천, event.원천행번호)
        acct = self._get_securities_account_name(event.원천)

        # event.총액 = 거래금액(청약금), 정산금액 = extra
        청약금 = event.총액
        수수료 = event.수수료
        정산 = event.extra.get('정산금액', 청약금 + 수수료)

        # 종목명 정규화
        normalized = event.종목명
        normalize_fn = getattr(self.company, 'normalize_security_name', None)
        if normalize_fn and normalized:
            normalized = normalize_fn(normalized)

        # 신규 종목 체크
        is_new = normalized and normalized not in self.portfolio.positions
        flags = [Flag.NEW_SECURITY] if is_new else []

        # 거래처 (회사 설정 format 사용)
        format_fn = getattr(self.company, 'format_partner', None)
        if format_fn and normalized:
            종목거래처 = format_fn('주식', normalized, '코스닥')
        else:
            종목거래처 = f"주식#코스닥#{normalized}"

        적요 = f"{normalized}({int(event.수량) if event.수량 else ''}주*@{event.단가:,.0f})납입" if event.단가 else f"{normalized} 청약납입"
        if not event.단가 or event.수량 == 0:
            적요 = f"{normalized} 청약납입"

        tx.add('차변', self._code('선급금'), '선급금', 청약금,
               거래처명=종목거래처, 적요=적요, flags=flags,
               메모='신규 청약종목 - 거래처 등록 필요' if is_new else '')
        if 수수료 > 0:
            tx.add('차변', self._code('주식거래수수료(판)'), '주식거래수수료(판)',
                   수수료, 거래처명=종목거래처, 적요='청약수수료', flags=flags)
        tx.add('대변', self._code('예치금'), '예치금', 정산,
               거래처명=acct, 적요=event.적요원본)
        return tx

    def _handle_은행이자(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '은행이자', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', 0)
        입금 = event.extra.get('입금', 0)
        bank = self._get_bank_account_name(event.원천)
        if 출금 > 0:
            tx.add('차변', self._code('선납세금'), '선납세금', 출금,
                   거래처명=self.company.PARTNERS.get('국세', '마포세무서'),
                   적요='결산이자세금')
            tx.add('대변', self._code('보통예금'), '보통예금', 출금,
                   거래처명=bank, 적요='결산이자세금')
        else:
            tx.add('차변', self._code('보통예금'), '보통예금', 입금,
                   거래처명=bank, 적요='결산이자')
            tx.add('대변', self._code('이자수익(금융)'), '이자수익(금융)', 입금,
                   거래처명=bank, 적요='결산이자')
        return tx

    def _handle_이체수수료(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '이체수수료', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('지급수수료(판)'), '지급수수료(판)', 출금,
               거래처명=bank, 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_증권사이체출금(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '증권사이체출금', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        증권사 = self._guess_securities_from_content(event.적요원본)
        # 증권사 이름 추정
        flags = [Flag.INFERRED_PARTNER] if 증권사 == '증권사(불명)' else []
        tx.add('차변', self._code('예치금'), '예치금', 출금,
               거래처명=증권사, 적요=event.적요원본, flags=flags)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_증권사이체입금(self, event: RawEvent) -> None:
        """IBK의 증권사이체입금: 증권사 이체출금에서 이미 처리 → skip"""
        return None

    def _handle_자사이체출금(self, event: RawEvent) -> None:
        """자사이체(은행↔은행 내부): 분개장에 없음 → skip"""
        return None

    def _handle_자사이체입금(self, event: RawEvent) -> None:
        """자사이체 입금: 분개장에 없음 → skip"""
        return None

    def _handle_카드결제_우리(self, event: RawEvent) -> Transaction:
        return self._handle_카드결제(event, '우리카드')

    def _handle_카드결제_하나(self, event: RawEvent) -> Transaction:
        return self._handle_카드결제(event, '하나카드')

    def _handle_카드결제_IBK(self, event: RawEvent) -> Transaction:
        return self._handle_카드결제(event, 'IBK기업카드')

    def _handle_카드결제_BC(self, event: RawEvent) -> Transaction:
        return self._handle_카드결제(event, 'BC카드')

    def _handle_카드결제(self, event: RawEvent, 카드사: str) -> Transaction:
        tx = Transaction(event.날짜, f'{카드사}결제', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        card_partner = bank
        for c in self.company.CARDS:
            if c.카드사 == 카드사:
                card_partner = c.거래처명
                break
        tx.add('차변', self._code('미지급금'), '미지급금', 출금,
               거래처명=card_partner, 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    # ========================================
    # 카드 사용
    # ========================================
    def _handle_카드사용(self, event: RawEvent) -> Transaction:
        tx = Transaction(event.날짜, '카드사용', event.원천, event.원천행번호)
        금액 = event.총액
        가맹점 = event.extra.get('가맹점', '')
        비고 = event.extra.get('비고', '')
        카드사 = event.extra.get('카드사', '')
        힌트 = event.extra.get('계정과목힌트', '')

        # 계정과목 결정 + 추정 여부
        flags = []
        메모 = ''
        if 힌트:
            계정명 = f"{힌트}(판)"
            if 계정명 not in self.company.ACCOUNTS:
                # 힌트 매칭 실패 → 검색
                매칭 = [k for k in self.company.ACCOUNTS if 힌트 in k]
                if 매칭:
                    계정명 = 매칭[0]
                else:
                    계정명 = '복리후생비(판)'
                    flags.append(Flag.INFERRED_ACCOUNT)
                    메모 = f"비고({비고})에서 계정 매칭 실패 - 기본값 적용"
        else:
            계정명 = '복리후생비(판)'
            flags.append(Flag.INFERRED_ACCOUNT)
            메모 = "비고 없음 - 기본값(복리후생비) 적용 - 검토 필요"

        # 카드 거래처
        card_partner = 가맹점
        for c in self.company.CARDS:
            if 카드사 and (c.카드사.startswith(카드사) or 카드사 in c.카드사):
                card_partner = c.거래처명
                break

        적요 = 비고 if 비고 else 가맹점

        tx.add('차변', self._code(계정명), 계정명, 금액,
               거래처명=가맹점, 적요=적요, flags=flags, 메모=메모)
        tx.add('대변', self._code('미지급금'), '미지급금', 금액,
               거래처명=card_partner, 적요=적요, flags=flags)
        return tx

    def _handle_카드승인(self, event: RawEvent) -> None:
        return None

    # ========================================
    # 일임계좌 청약 관련 (IBK 은행)
    # ========================================
    def _handle_일임청약출금(self, event: RawEvent) -> Transaction:
        """일임/고유/고위험 청약 납입: 예수금(차) / 보통예금(대)  거래처: 수탁운용"""
        tx = Transaction(event.날짜, '일임청약출금', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('예수금'), '예수금', 출금,
               거래처명='수탁운용', 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_일임청약환급입금(self, event: RawEvent) -> Transaction:
        """청약 환급입금(웨스트/비엔케이): 보통예금(차) / 예수금(대)  거래처: 수탁운용"""
        tx = Transaction(event.날짜, '일임청약환급입금', event.원천, event.원천행번호)
        입금 = event.extra.get('입금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('보통예금'), '보통예금', 입금,
               거래처명=bank, 적요=event.적요원본)
        tx.add('대변', self._code('예수금'), '예수금', 입금,
               거래처명='수탁운용', 적요=event.적요원본)
        return tx

    def _handle_에이치엔입금(self, event: RawEvent) -> Transaction:
        """에이치엔인베스트 입금 (수탁운용): 보통예금(차) / 예수금(대)"""
        tx = Transaction(event.날짜, '에이치엔입금', event.원천, event.원천행번호)
        입금 = event.extra.get('입금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('보통예금'), '보통예금', 입금,
               거래처명=bank, 적요=event.적요원본)
        tx.add('대변', self._code('예수금'), '예수금', 입금,
               거래처명='수탁운용', 적요=event.적요원본)
        return tx

    def _handle_청약납입_은행(self, event: RawEvent) -> Transaction:
        """
        고유/르네상스 청약납입 (IBK 은행 출금):
        선급금(차) + 주식거래수수료(판)(차) / 보통예금(대)
        
        수수료 = 청약금 × 1% (분개장 패턴: 총액 / 1.01 = 청약금)
        종목명은 적요에서 추출 시도
        """
        tx = Transaction(event.날짜, '청약납입_은행', event.원천, event.원천행번호)
        총액 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)

        # 수수료 1% 추정 (총액 = 청약금 × 1.01)
        청약금 = int(round(총액 / 1.01))
        수수료 = 총액 - 청약금
        
        # 종목명 추출 (적요에서 '-' 또는 '－' 앞부분)
        내용 = event.적요원본
        import re
        종목명_raw = re.split(r'[-－]', 내용)[0].strip()
        # 정규화
        normalize_fn = getattr(self.company, 'normalize_security_name', None)
        종목명 = normalize_fn(종목명_raw) if normalize_fn else 종목명_raw
        
        # 거래처명
        format_fn = getattr(self.company, 'format_partner', None)
        종목거래처 = format_fn('주식', 종목명, '코스닥') if format_fn and 종목명 else f'주식#코스닥#{종목명}'
        
        flags = [Flag.INFERRED_ACCOUNT]
        메모 = f'수수료 1% 추정. 정확한 청약금/수수료는 증권사 거래내역 참조'
        
        tx.add('차변', self._code('선급금'), '선급금', 청약금,
               거래처명=종목거래처, 적요=f"{종목명} 납입",
               flags=flags, 메모=메모)
        if 수수료 > 0:
            tx.add('차변', self._code('주식거래수수료(판)'), '주식거래수수료(판)',
                   수수료, 거래처명=종목거래처, 적요='청약수수료', flags=flags)
        tx.add('대변', self._code('보통예금'), '보통예금', 총액,
               거래처명=bank, 적요=내용)
        return tx

    def _handle_미지급금결제(self, event: RawEvent) -> Transaction:
        """관리비 등 기계상된 미지급금 현금 결제: 미지급금(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '미지급금결제', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('미지급금'), '미지급금', 출금,
               거래처명='LG팰리스빌딩관리단', 적요=event.적요원본,
               flags=[Flag.INFERRED_PARTNER], 메모='거래처 확인 필요')
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    # ========================================
    # 급여 / 퇴직연금
    # ========================================
    def _handle_급여지급(self, event: RawEvent) -> Transaction:
        """급여 지급: 미지급금(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '급여지급', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        수취인 = event.적요원본
        tx.add('차변', self._code('미지급금'), '미지급금', 출금,
               거래처명='직원급여', 적요=수취인)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=수취인)
        return tx

    def _handle_퇴직연금DC(self, event: RawEvent) -> Transaction:
        """퇴직연금 DC형 납입: 임원/직원퇴직급여(판)(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '퇴직연금DC', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        # 분개장에서 임원/직원 각 50:50
        half = 출금 // 2
        flags = [Flag.INFERRED_ACCOUNT]
        메모 = '임원/직원 퇴직연금 비율 50:50 추정 - 실제 계약 확인'
        tx.add('차변', self._code('임원퇴직급여(DC)(판)'), '임원퇴직급여(DC)(판)',
               half, 거래처명='', 적요=event.적요원본, flags=flags, 메모=메모)
        tx.add('차변', self._code('직원퇴직급여(DC)(판)'), '직원퇴직급여(DC)(판)',
               출금 - half, 거래처명='', 적요=event.적요원본, flags=flags)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_퇴직연금부담금(self, event: RawEvent) -> Transaction:
        """퇴직연금 부담금: 임원/직원퇴직급여(DC)(판)(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '퇴직연금부담금', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        # 분개장: 833334 임원 / 416666 직원 (2:1 비율)
        임원분 = int(round(출금 * 2 / 3))
        직원분 = 출금 - 임원분
        flags = [Flag.INFERRED_ACCOUNT]
        tx.add('차변', self._code('임원퇴직급여(DC)(판)'), '임원퇴직급여(DC)(판)',
               임원분, 거래처명='', 적요=event.적요원본, flags=flags,
               메모='임원:직원 약 2:1 추정')
        tx.add('차변', self._code('직원퇴직급여(DC)(판)'), '직원퇴직급여(DC)(판)',
               직원분, 거래처명='', 적요=event.적요원본, flags=flags)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_퇴직연금수수료(self, event: RawEvent) -> Transaction:
        """퇴직연금 수수료: 지급수수료(판)(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '퇴직연금수수료', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        # 분개장에서는 운용관리/자산관리 두 줄로 나뉘지만 총액 하나로 처리
        flags = [Flag.INFERRED_ACCOUNT]
        tx.add('차변', self._code('지급수수료(판)'), '지급수수료(판)', 출금,
               거래처명='한국투자증권 80931954-29', 적요=event.적요원본,
               flags=flags, 메모='운용관리/자산관리 수수료 합산 - 세부 분리 필요')
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    # ========================================
    # 기타 은행 지출
    # ========================================
    def _handle_임대료(self, event: RawEvent) -> Transaction:
        """임대료: 미지급금(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '임대료', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('미지급금'), '미지급금', 출금,
               거래처명='김민영', 적요=event.적요원본,
               flags=[Flag.INFERRED_PARTNER], 메모='임대인 이름 확인 필요')
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_통신비(self, event: RawEvent) -> Transaction:
        """통신비 (SKT/KT): 복리후생비(판)(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '통신비', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('복리후생비(판)'), '복리후생비(판)', 출금,
               거래처명='', 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    def _handle_접대비_기타(self, event: RawEvent) -> Transaction:
        """조의금/경조 등: 접대비(기업업무추진비)(판)(차) / 보통예금(대)"""
        tx = Transaction(event.날짜, '접대비_기타', event.원천, event.원천행번호)
        출금 = event.extra.get('출금', event.총액)
        bank = self._get_bank_account_name(event.원천)
        tx.add('차변', self._code('접대비(기업업무추진비)(판)'),
               '접대비(기업업무추진비)(판)', 출금,
               거래처명='', 적요=event.적요원본)
        tx.add('대변', self._code('보통예금'), '보통예금', 출금,
               거래처명=bank, 적요=event.적요원본)
        return tx

    # ========================================
    # 헬퍼 메소드
    # ========================================
    def _code(self, 계정과목: str) -> int:
        return self.company.get_account_code(계정과목)

    def _get_securities_account_name(self, 원천: str) -> str:
        for s in self.company.SECURITIES_ACCOUNTS:
            if s.별칭 == 원천 or 원천 in s.별칭:
                return s.거래처명
        return 원천

    def _get_bank_account_name(self, 원천: str) -> str:
        for b in self.company.BANK_ACCOUNTS:
            if b.별칭 == 원천 or 원천 in b.별칭:
                return b.거래처명
        if self.company.BANK_ACCOUNTS:
            return self.company.BANK_ACCOUNTS[0].거래처명
        return '보통예금'

    def _infer_counterparty_bank(self) -> str:
        if self.company.BANK_ACCOUNTS:
            return self.company.BANK_ACCOUNTS[0].거래처명
        return ''

    def _format_security_partner(self, event: RawEvent) -> str:
        """종목 → 거래처명 (회사 설정의 format_partner 사용)"""
        # 회사 설정에 format_partner가 있으면 그것 사용 (회사별 형식 차이 처리)
        format_fn = getattr(self.company, 'format_partner', None)
        if format_fn:
            return format_fn(
                event.종목유형 or '주식',
                event.종목명,
                event.시장구분,
            )
        # fallback
        if event.종목유형 == '주식':
            시장 = event.시장구분 or '코스닥'
            return f"주식#{시장}#{event.종목명}"
        elif event.종목유형 == '채권':
            종류 = event.시장구분 or '회사채'
            return f"채권#{종류}#{event.종목명}"
        elif event.종목유형 == '펀드':
            return f"펀드#{event.종목명}"
        return event.종목명

    def _guess_securities_from_content(self, content: str) -> str:
        for s in self.company.SECURITIES_ACCOUNTS:
            if s.증권사.replace('증권', '') in content:
                return s.거래처명
        return '증권사(불명)'
