# 📊 더존 위하고 전표 변환기

거래내역 엑셀 파일을 더존 위하고 일반전표 업로드 형식으로 자동 변환하는 Streamlit 앱.

---

## 🏢 지원 회사 (9개)

| # | 회사 | 자동 감지 키워드 | 계정코드 | 상태 |
|---|---|---|---|---|
| 1 | **리버사이드파트너스** | `IBK 기업은행`, `한국투자증권`, `교보증권`, `비씨카드` | 5자리 | ✅ 11개 일치 |
| 2 | **두리인베스트먼트** | `키움(...)`, `메리츠(...)`, `한국(80...)`, `IBK기업은행 내역` | 3자리 | ✅ 3개 일치 + 🟡 3개 근접 |
| 3 | **엔라이튼자산운용** | `하나은행거래내역_*`, `법인카드승인내역`, `공모주실현손익` | 5자리 | ✅ 3개 일치 + 🟡 3개 근접 |
| 4 | **케이피자산운용** (구 한국연금투자자문) | `미래006/220/520/710/720/730`, `단기매매증권_거래_내역` | 5자리 (KS/KQ 분리) | 🟡 5개 근접 |
| 5 | **어라운드자산운용** | `_거래내역서_` + 증권사명 (db/im/kb/nh/교보/신영/신한/키움/한양/한투) | 5자리 | 🟡 3개 근접 |
| 6 | **밸류어블파트너스** | `밸류어블`, `고객납입용`, `보험납입`, `하이아웃`, `하이일드채권`, `2771-XXXX` | 3자리 | ✅ |
| 7 | **로만자산운용** | `계좌별_거래내역_050_`, `계좌별_거래내역_704_`, `법인이용대금명세서_20260406` | 3자리 | ✅ |
| 8 | **아스투자일임** | `법인이용내역(전체)_회계양식`, `140-011-650180`, `140-013-357558`, `LS증권` | 3자리 | ✅ |
| 9 | **센트럴밸류파트너스** | `센트럴밸류`, `_거래내역_N(하이일드\|고유\|투자일임)`, `NH농협_은행계좌`, `신한은행_은행계좌`, iM/하나/한화투자/신영/메리츠/교보증권 거래내역 | 3자리 ⚠️추정치 | ✅ |

> ⚠️ **센트럴밸류파트너스**: `CENTRAL_AC` 계정코드는 추정치입니다. 처음 사용 시 실제 더존 설정과 대조 후 사용하세요.

업로드된 파일의 시트명 + 파일명을 보고 회사를 **자동 판단**합니다.

---

## 🚀 사용 방법

1. 거래내역 엑셀 파일들 다운로드 → Streamlit 앱에 업로드 (여러 파일 동시 가능)
2. **🔄 변환 시작** 버튼 클릭
3. 화면 표시:
   - 🏢 자동 감지된 회사명
   - 📊 전표 행 수, 미분류 건수, 차/대변 합계
   - 🆕 새 종목 감지 시 추가 코드
4. **📥 변환 파일 다운로드** → 색상 적용된 엑셀 파일 받기
5. 더존 위하고 업로드 전 **🔴 빨간 행은 직접 수정** (취득가 등)

---

## 🎨 변환 결과 색상

| 색상 | 의미 | 조치 |
|---|---|---|
| 🔴 **빨간 행** | 취득가 확인 필요 (금액 0) | 매수 단가 직접 입력 |
| 🟠 **주황 행** | 미분류 거래 | 수동 확인 후 계정 입력 |
| 🟡 **노란 행** | 헤더 | - |

---

## ⚠️ 자동변환 영역 외 (수동 입력 필요)

거래내역 파일에 원천 데이터가 없는 **결산성 분개**는 자동변환 안 됨.

### 모든 회사 공통
- 월말 평가상계 / 평가이익 / 평가손실
- 손익계정 대체 / 이월이익잉여금 / 미처분이익잉여금
- 직원/임원/퇴직급여, 상여금, 법인세

### 회사별 추가 결산성 항목

- **두리**: 임원/직원 퇴직급여(DC) 281M, 상여금 120M, 미처분/이월이익잉여금 76억
- **엔라이튼**: 임원/직원급여 78M, 감가상각/리스부채/임차보증금 25M+, 교육세(금융) 차액 잡손실 수동 분리
- **케이피**: 이연법인세자산 528M, 사용권자산/감가상각/퇴직연금
- **어라운드**: 운용보수 미수계상 24M+, 선물거래손실, 매도가능증권
- **밸류어블**: 월초 평가상계, 월말 평가손익, 감가상각, 법인세, 급여·세금 결산 분개
- **로만**: 월초/월말 매도가능증권 평가상계, 공모주 배정주식 입고, 리스부채/이자비용, 연말정산 이월차감
- **센트럴밸류**: 월초/월말 단매증 평가, 손익계정 대체, 퇴직급여, 법인세

---

## 📊 회사별 지원 파일 목록

### 엔라이튼자산운용

| 파일 | 파서 | 비고 |
|---|---|---|
| `하나은행거래내역_02504.xls` | `process_enra_hana('02504')` | 운용보수 계좌 |
| `하나은행거래내역_18404.xls` | `process_enra_hana('18404')` | 메인 경영관리 |
| `하나은행거래내역_67204.xls` | `process_enra_hana('67204')` | 고유자산 |
| `법인카드승인내역(전체).xlsx` | `process_enra_corp_card` | 업종별 자동 분류 |
| `공모주실현손익.xlsx` (26.03 시트 등) | `process_enra_ipo_pnl` | 월 시트명 자동 감지 |

**4월 기준**: 전표 1,537행, 🟠 ORANGE 1건 (교육세 차액 수동 분리 필요), 미분류 0건

### 밸류어블파트너스

| 파일 | 파서 | 비고 |
|---|---|---|
| `비용지출_*.xlsx` (2026.*월분 시트) | `process_valuable_expenses` | 월 자동 감지 |
| `기업카드이용내역.xlsx` | `process_valuable_expenses` | `인쇄및파일저장` 시트 |
| `국민은행_*납입*.xls` | `process_valuable_kbbank` | 시트명에 `납입` 포함 |
| `기업은행_*.xls` | `process_valuable_ibk` | |
| `키움_*.xlsx` | `process_valuable_kiwoom` | |
| `미래에셋_*.xlsx` | `process_valuable_mirae_haoet` / `_mirae_acct` | 하이아웃/계좌번호 자동 분기 |
| `NH증권_*.xlsx` | `process_valuable_nh` | |
| `KB증권_*.xlsx` | `process_valuable_kb_sec` | |
| `삼성증권_*.xlsx` | `process_valuable_samsung` | |
| `신한증권_*.xlsx` | `process_valuable_shinhan` | |
| `한국투자증권_*.xlsx` | `process_valuable_hantoo_val` | |
| `교보증권_*.xlsx` | `process_valuable_kyobo` | Sheet1, 3행쌍 |
| `대신증권_*.xls` | `process_valuable_daeshin` | `거래내역_전체` 시트 |
| `한화증권_*.xlsx` | `process_valuable_hanhwa` | `myTable` 시트 |

**4월 기준**: 전표 367행, 🟠 ORANGE 18건, 미분류 0건

### 센트럴밸류파트너스

| 파일 | 파서 | 비고 |
|---|---|---|
| `신한은행_은행계좌_거래내역.xlsx` | `process_central_shinhan_bank` | `sheet` 시트 |
| `NH농협_은행계좌_거래내역.xls` | `process_central_nh_bank` | |
| `미래에셋증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_mirae` | |
| `삼성증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_samsung` | |
| `신한투자증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_shinhan_sec` | X 접미사도 감지 |
| `NH투자증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_nh_sec` | `8203` 시트 |
| `KB증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_kb` | |
| `한국투자증권_거래내역_N(계좌유형)_O.xls` | `process_central_hantoo` | |
| `키움증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_kiwoom` | |
| `유진투자증권_거래내역_N(계좌유형).xlsx` | `process_central_yujin` | `_O` 없어도 감지 |
| `iM증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_im` | 🆕 4월 |
| `하나증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_hana` | 🆕 4월 |
| `한화투자증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_hanhwa` | 🆕 4월 |
| `신영증권_거래내역_N(계좌유형)_O.xls` | `process_central_shinyoung` | 🆕 4월 |
| `메리츠증권_거래내역_N(계좌유형)_O.xls` | `process_central_meritz` | 🆕 4월 |
| `교보증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_kyobo` | 🆕 4월 |
| `DB금융투자_거래내역_N(계좌유형)_O.xlsx` | `process_central_db` | |
| `BNK증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_bnk` | |
| `유안타증권_거래내역_N(계좌유형)_O.xlsx` | `process_central_yuanta` | 신탁 |
| `대신증권_거래내역_N(계좌유형)_O.xls` | `process_central_daeshin` | HTML xls |

> `N(계좌유형)` = `1(하이일드)`, `2(고유)`, `3(투자일임)` 등, 파일명에서 자동 추출

**4월 기준**: 전표 997행, 미분류 9건 (유진 합병출고 5건 + 기타 0원 자금이동 4건)

---

## 🛠️ 기술 정보

**총 약 7,730줄** | 9개 회사 자동 감지 | 280개+ 계정 매핑 | 280개+ 종목

### 파서 구조

| 회사 | 핵심 파서 |
|---|---|
| 리버사이드 | `process_hantoo_sheet`, `process_kyobo`, `process_ibk`, `process_card` |
| 두리 | `process_duri_ibk`, `process_duri_kiwoom`, `process_duri_meritz`, `process_duri_hanguk`, `process_duri_card_usage` |
| 엔라이튼 | `process_enra_hana` + `classify_enra_hana`, `process_enra_corp_card`, `process_enra_ipo_pnl` |
| 케이피 | `process_kp_securities` (KOSPI/KOSDAQ), `process_kp_card`, `process_kp_bank` |
| 어라운드 | `process_around_card`, `process_around_bank`, `process_around_securities` (10개 증권사) |
| 밸류어블 | `process_valuable_expenses`, `process_valuable_kbbank`, `process_valuable_ibk`, `process_valuable_kiwoom`, `process_valuable_mirae_haoet`, `process_valuable_mirae_acct`, `process_valuable_nh`, `process_valuable_kb_sec`, `process_valuable_samsung`, `process_valuable_shinhan`, `process_valuable_hantoo_val`, `process_valuable_kyobo`, `process_valuable_daeshin`, `process_valuable_hanhwa` |
| 로만 | `process_roman_bank050`, `process_roman_bank704`, `process_roman_card` |
| 아스 | `process_ast_bank650`, `process_ast_bank357`, `process_ast_card`, `process_ast_meritz`, `process_ast_hantoo`, `process_ast_ls` |
| 센트럴밸류 | `_central_sec_row` (공통), `_is_central_ignore` (공백 포함 무시 패턴), `process_central_*` (20개) |

---

## 🆕 새 종목 추가 방법

매월 새 주식/채권/펀드에 투자하면 변환 후 화면에 **"🆕 신규 종목 감지"** 박스가 표시됩니다. 코드를 복사해 Claude에게 보내주세요.

```python
# 예시 - STOCK_DB 추가분:
'새종목A': '주식#코스피#새종목A',  # ← 코스피/코스닥/회사채 구분 필요
```

---

## 🔄 GitHub 업로드 방법

### ⚠️ 복사/붙여넣기 절대 금지

GitHub 웹 편집기 복붙 시 한글·이모지·정규식이 깨짐. **파일 자체를 업로드**하세요.

1. GitHub `doozon-converter` → `app/streamlit_app.py` 클릭
2. `[...]` → **Delete file** → Commit changes
3. `app/` 폴더 → **Add file → Upload files**
4. 다운받은 `streamlit_app.py` 드래그 앤 드롭 → **Commit changes**
5. 1~2분 대기 → Streamlit Cloud 자동 재배포

README도 루트 경로에 동일 방식으로 업로드하세요.

---

## 📝 작업 이력

| 차수 | 내용 |
|---|---|
| 1차 | 리버사이드파트너스 — 색상 표시, 비씨카드, 매도/매수/청약 수수료 |
| 2차 | 두리인베스트먼트 — 3자리 계정, IBK/키움/메리츠/한국증권/카드 파서 |
| 3차 | 엔라이튼자산운용 — 하나은행 3계좌, 법인카드 부가세 분리, 공모주실현손익 |
| 4차 | 케이피자산운용 — KS/KQ 분리, 미래에셋 6계좌, 하나카드 12장 |
| 5차 | 어라운드자산운용 — 증권사 10개 파서 |
| 6차 | 밸류어블파트너스 — 증권사 10개 파서, 비용지출 카드 |
| 7차 | 로만자산운용 — 신한은행 2계좌, 신한카드 3종 |
| 8차 | 아스투자일임 — 신한카드 XLS 계정명 자동분류, LS증권 |
| 9차 | 센트럴밸류파트너스 신규 + 밸류어블 3월 — 파서 20개, 교보/대신/한화 추가 |
| **10차** | **4월 2026 대규모 업데이트** |

### 10차 상세 (4월 2026)

**엔라이튼**
- 신규 종목: 키움히어로제2호스팩, 채비, 코스모로보틱스, 신한제18호스팩
- 신규 계정: 미지급세금(26100), 청약예치금(12500), 부가세예수금(25500)
- 신규 처리: 교육세(금융), 1기예정부가세, 명함제작비, 메이븐 자문수수료, 스팩 매도대금, NH투자증권 운용보수, 일임납입
- 법인카드 업종 추가: 치킨전문점, 기타서비스, 스포츠용품점, 일반의류, 기타운송기구, 기타레저스포츠, 약국, 지하철, 우편, 공공편의서비스

**밸류어블**
- 미래에셋 시트명이 계좌번호로 변경 → 자동 분기 처리
- 신한제18호스팩 파서 신규

**센트럴밸류**
- 월 필터 제거: `m != 3` 조건 12건 전부 제거 (매월 처리 가능)
- 신규 파서 7개: iM증권, 하나증권, 한화투자증권, 신영증권, 메리츠증권, 교보증권(센트럴), 유진투자증권 전면 교체
- `_central_sec_row` 강화: 매도 25종, 매수 10종, ignore 18종 패턴 추가
- `_is_central_ignore` 헬퍼: 공백 포함 거래종류 정규화
- 한투 파서 컬럼 인덱스 수정
- 신규 종목: 코스모로보틱스, 신한제18호스팩, 키움히어로제2호스팩, 채비

---

## ❓ 문제 해결

| 증상 | 원인 및 해결 |
|---|---|
| "변환된 데이터가 없습니다" | 파일명/시트명 패턴 미일치 → Claude에게 파일명 + 헤더 알려주세요 |
| 자동 감지 오류 | 시트명에 식별 키워드 부족 → 패턴 추가 요청 |
| 🟠 교육세 ORANGE (엔라이튼) | `미지급세금 + 잡손실`로 더존에서 수동 분리 필요 |
| 센트럴밸류 계정코드 불일치 | `CENTRAL_AC`는 추정치 → 실제 코드로 수정 요청 |
| 대신증권 HTML xls | 거래 없는 월은 빈 HTML → 자동 스킵 (정상) |

---

## 📞 문의

종목 추가, 거래 분류 변경, 새 회사 지원 등은 **Claude에게 요청**해 주세요.

다음 달 분개장 + 거래내역 + 현재 `streamlit_app.py`를 같이 올려주시면 이어서 작업할 수 있어요.
