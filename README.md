# ALIO Crawler

알리오(www.alio.go.kr) 공공기관 경영정보 공개시스템 항목별 공시 일괄 수집기.

> 알리오 사이트에서 기관마다 일일이 클릭해서 받아야 하는 PDF·HWP·엑셀 첨부파일을 자동으로 폴더에 정리해 받는 도구.

### 지금 바로 시작하기

1. **[Download ZIP](https://github.com/giovinazo/alio-crawler/archive/refs/heads/main.zip)** 클릭
2. 압축 풀고 `alio_crawler_v5.4.py` 실행 (`python3 alio_crawler_v5.4.py`)

> 설치가 어려우면 — ZIP을 [Claude Desktop](https://claude.ai/download) 대화에 올리고 **"이거 설치 도와줘"** 입력

## 주요 기능

- 알리오 항목별공시 **전 항목(92개, 2026-05 기준) 자동 수집** — `formList.json` API 활용
- **다중 기관 × 다중 항목 일괄 다운로드** — UI 트리뷰 체크박스
- 첨부파일 다운로드 엔드포인트 통합 (PDF / file / dfile / rule + 게시판형 2종)
- 게시판형 항목(감사원 지적사항, 국회지적사항, 임원 모집공고 등) 첨부 자동 수집
- 내부규정(rule): findRuleList → findRuleDtl → bFiles 파싱 → rulefiledown 다운로드
- 자체감사(audit) / 경영실적 평가(mgmt_eval): itemReportListSusi → 보고서별 첨부 다운로드
- 입찰공고 외부 링크(g2b.go.kr) URL 자료_목록.txt에 정리
- 폴더 자동 분류: `ALIO_타임스탬프 / 항목명 / 기관명 / 파일들`
- 기관 단위 ThreadPoolExecutor 병렬 처리(envlaw 기준 약 5배 속도)
- 알리오 항목 수 변경 시 자동 캐시 갱신 (`alio_items.json`)

## 환경

- Python 3.10+ (개발: macOS 12, Python 3.14)
- GUI: 순수 tkinter + ttk 기본 테마
- 의존: `requests`, `openpyxl`, `beautifulsoup4`

## 공유 라이브러리 (`alio_core.py`)

알리오 API 호출·HTML 파싱·다운로드 코어 600라인은 **`alio_core.py`** 한 파일에 모여 있다. 이 파일은 [alio-mcp](https://github.com/giovinazo/alio-mcp) 레포의 정본을 **sync한 사본**이며, 두 프로젝트가 동일 코어를 공유한다.

- **직접 수정 금지** — 알리오 사이트 패턴이 바뀌면 alio-mcp 정본을 수정한 뒤 sync한다.
- **동기화 절차** (alio-mcp 폴더에 형제로 위치한 경우):
  ```bash
  cd ../alio-mcp && ./sync_to_crawler.sh
  ```
  다른 위치라면 `CRAWLER_DIR=/path/to/this ./sync_to_crawler.sh`.
- alio-mcp 없이도 본 레포 단독으로 `python3 alio_crawler_v5.4.py` 실행 가능.

## 실행

```bash
python3 alio_crawler_v5.4.py
```

GUI 흐름:

1. **기관 필터** 선택 (기관유형/주무부처/지역)
2. **공시항목 선택…** 클릭 → 트리뷰에서 다중 선택
3. **1.기관목록수집** → 결과 트리뷰에서 기관 체크
4. **2.공시내용수집** → 폴더에 자동 저장
5. (선택) **3.엑셀저장** — 항목별 시트로 데이터 추출

## 검증 스크립트

| 파일 | 용도 |
|---|---|
| `self_check_v5_4_1.py` | 회귀 점검 (20건 PASS/FAIL) |
| `full_audit_v5_4_1.py` | 83개 항목 매핑·API 응답 분석 |
| `full_download_audit.py` | 항목별 다운로드 가능 여부 빠른 점검 |
| `precise_audit.py` | type별 정확한 흐름 재현 정밀 점검 |

```bash
python3 self_check_v5_4_1.py    # 약 30초
python3 precise_audit.py        # 약 5분
```

## 폴더 구조 예시

```
저장경로/
└─ ALIO_20260427_192556/
   ├─ 내부규정/
   │  ├─ 한국산업단지공단/
   │  │  ├─ 정관.pdf
   │  │  └─ 인사규정.hwp
   │  └─ ...
   └─ 감사원 지적사항/
      └─ 한국무역보험공사/
         ├─ 자료_목록.txt
         └─ 01_2025년도_감사원_규제개선/
            └─ 감사보고서.pdf
```

## 버전 히스토리

| 버전 | 주요 변경 |
|---|---|
| v5.4.2 | 알리오 항목 92개로 확장 (ESG 운영·AI 활용 카테고리 신설) + precise_audit.py 정확화 (rule/audit/mgmt_eval 실제 흐름 반영) |
| v5.4.1 | 게시판형 첨부 2개 패턴 통합 + 콤마 fallback + 외부 링크 + envlaw 병렬화 |
| v5.4 | 83개 항목 자동 수집 + UI 다중 선택 + 폴더 구조 재편 |
| v5.3 | 경영실적 평가결과 추가 (Susi API) |
| v5.2.x | 자체감사·수의계약·사망자수·환경법규 등 점진 추가 |

## 저작권

© 2025-2026 허재영. 자세한 내용은 [저작권 설명서](저작권_설명서_ALIO_Crawler.md).
