# -*- coding: utf-8 -*-
"""
ALIO 크롤러 v5.4.2 자체점검 스크립트
- GUI 없이 핵심 기능을 자동 검증
- 한국산업단지공단(C0208) 데이터로 실제 API 호출·다운로드 검증
- 결과: PASS / FAIL / SKIP 표 출력

실행: python3 self_check_v5_4_1.py
"""
from __future__ import annotations

import importlib.util
import os
import shutil
import sys
import tempfile
import time
import urllib3
from datetime import datetime

urllib3.disable_warnings()

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET_FILE = os.path.join(HERE, "alio_crawler_v5.4.py")
TEST_APBA_ID = "C0208"  # 한국산업단지공단
TEST_APBA_NM = "한국산업단지공단"


# ── 결과 누적기 ────────────────────────────────────
class CheckResults:
    def __init__(self):
        self.results = []

    def add(self, name, status, detail=""):
        self.results.append((name, status, detail))
        # 즉시 출력
        symbol = {"PASS": "✓", "FAIL": "✗", "SKIP": "—"}.get(status, "?")
        color = {"PASS": "\033[92m", "FAIL": "\033[91m", "SKIP": "\033[93m"}.get(status, "")
        reset = "\033[0m"
        print(f"  {color}[{symbol}] {status:4s}{reset} {name}")
        if detail:
            print(f"           └ {detail}")

    def summary(self):
        total = len(self.results)
        passes = sum(1 for _, s, _ in self.results if s == "PASS")
        fails = sum(1 for _, s, _ in self.results if s == "FAIL")
        skips = sum(1 for _, s, _ in self.results if s == "SKIP")
        return passes, fails, skips, total


# ── 모듈 로더 ──────────────────────────────────────
def load_module():
    spec = importlib.util.spec_from_file_location("alio_v54", TARGET_FILE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


# ── 점검 실행 ──────────────────────────────────────
def main():
    print("=" * 70)
    print(f"ALIO 크롤러 v5.4.2 자체점검 ({datetime.now():%Y-%m-%d %H:%M:%S})")
    print(f"대상: {TARGET_FILE}")
    print(f"테스트 기관: {TEST_APBA_NM} (apbaId={TEST_APBA_ID})")
    print("=" * 70)

    R = CheckResults()
    tmp_root = tempfile.mkdtemp(prefix="alio_self_check_")
    try:
        # ─── 1. 환경 ─────────────────────────────
        print("\n[1] 환경·모듈")
        try:
            m = load_module()
            R.add("모듈 import", "PASS", f"{TARGET_FILE.split('/')[-1]}")
        except Exception as e:
            R.add("모듈 import", "FAIL", str(e))
            return _final_report(R)

        # 필수 함수 존재
        required = [
            "fetch_alio_items", "get_alio_items", "load_alio_items_cache",
            "save_alio_items_cache", "build_item_display_name", "build_item_root_no",
            "group_items_by_category", "ENDPOINT_REGISTRY", "detect_endpoint_kind",
            "build_save_path", "download_file_to_path", "download_attachment",
            "fetch_board_attachment_list", "download_board_attachment",
            "ItemSelectorDialog", "ALIOCrawlerGUI",
        ]
        missing = [name for name in required if not hasattr(m, name)]
        if missing:
            R.add("필수 함수·클래스 정의", "FAIL", f"누락: {missing}")
        else:
            R.add("필수 함수·클래스 정의", "PASS", f"{len(required)}/{len(required)}")

        # 캐시 파일
        cache_path = m.get_alio_items_cache_path()
        if os.path.exists(cache_path):
            R.add("캐시 파일 존재", "PASS", f"{os.path.getsize(cache_path):,} bytes")
        else:
            R.add("캐시 파일 존재", "SKIP", "갱신 시 자동 생성")

        # ─── 2. API 연결 ────────────────────────
        print("\n[2] API 연결")
        sess = m.create_session(verify_ssl=False)
        sess.headers.update(m.HEADERS)

        # formList.json
        try:
            items = m.fetch_alio_items()
            if len(items) >= 80:
                R.add("formList.json 호출", "PASS", f"{len(items)}개 항목 수집")
            else:
                R.add("formList.json 호출", "FAIL", f"항목 수 부족: {len(items)}개")
        except Exception as e:
            R.add("formList.json 호출", "FAIL", str(e))
            items = []

        # itemOrganListJung.json (보고서형 - 일반현황)
        try:
            url = f"{m.BASE_URL}/item/itemOrganListJung.json"
            body = {
                "reportFormRootNo": "10105",
                "apbaType": [], "jidtDptm": [], "area": [],
                "apba_id": "", "pageNo": 1,
            }
            r = m.retry_request(sess, "POST", url, json=body,
                                headers={"Content-Type": "application/json;charset=UTF-8"},
                                timeout=15)
            organs = r.json().get("data", {}).get("organList", [])
            if isinstance(organs, dict):
                organs = organs.get("result", [])
            if len(organs) >= 100:
                R.add("itemOrganListJung 호출 (보고서형)", "PASS", f"{len(organs)}개 기관")
            else:
                R.add("itemOrganListJung 호출 (보고서형)", "FAIL", f"기관 수 부족: {len(organs)}개")
        except Exception as e:
            R.add("itemOrganListJung 호출 (보고서형)", "FAIL", str(e))

        # itemReportListSusi.json (게시판형 - 감사원 지적사항)
        try:
            url = f"{m.BASE_URL}/item/itemReportListSusi.json"
            body = {
                "pageNo": 1, "apbaId": TEST_APBA_ID, "apbaType": "",
                "reportFormRootNo": "B1220-P2200",
                "search_word": "", "search_flag": "title",
                "bid_type": "", "enfc_istt": "",
            }
            r = m.retry_request(sess, "POST", url, json=body,
                                headers={"Content-Type": "application/json;charset=UTF-8"},
                                timeout=15)
            result = r.json().get("data", {}).get("result", [])
            if len(result) > 0:
                R.add("itemReportListSusi 호출 (게시판형)", "PASS", f"감사원 지적사항 {len(result)}건")
            else:
                R.add("itemReportListSusi 호출 (게시판형)", "SKIP", "한산공 자료 없음")
        except Exception as e:
            R.add("itemReportListSusi 호출 (게시판형)", "FAIL", str(e))
            result = []

        # ─── 3. 매칭 로직 ───────────────────────
        print("\n[3] 매칭 로직")
        if items:
            to_legacy = m.ALIOCrawlerGUI._item_meta_to_legacy.__get__(type("X", (), {})())
            v53_names = list(m.DISCLOSURE_ITEMS.keys())
            matched = set()
            for it in items:
                legacy = to_legacy(it)
                nm = legacy.get("name")
                if nm in v53_names:
                    matched.add(nm)
            if len(matched) == len(v53_names):
                R.add("v5.3 11개 항목 매칭", "PASS", f"{len(matched)}/{len(v53_names)}")
            else:
                missing = set(v53_names) - matched
                R.add("v5.3 11개 항목 매칭", "FAIL", f"누락: {missing}")

            # 신규 항목 type 분포
            from collections import Counter
            new_types = Counter()
            for it in items:
                legacy = to_legacy(it)
                if legacy.get("name") not in v53_names:
                    new_types[legacy.get("type", "?")] += 1
            if new_types:
                detail = ", ".join(f"{k}:{v}" for k, v in new_types.most_common())
                R.add("신규 항목 type 분포", "PASS", detail)
            else:
                R.add("신규 항목 type 분포", "SKIP", "신규 항목 없음")
        else:
            R.add("v5.3 11개 항목 매칭", "SKIP", "items 없음")

        # ─── 4. 유틸 함수 ───────────────────────
        print("\n[4] 유틸 함수")
        # build_save_path
        try:
            p = m.build_save_path(tmp_root, "테스트 항목/특수<>:?", "한산공\\테스트")
            if os.path.exists(p):
                R.add("build_save_path", "PASS", os.path.relpath(p, tmp_root))
            else:
                R.add("build_save_path", "FAIL", "폴더 미생성")
        except Exception as e:
            R.add("build_save_path", "FAIL", str(e))

        # _resolve_collision_path
        try:
            test_dir = m.build_save_path(tmp_root, "충돌테스트", "기관")
            f1 = os.path.join(test_dir, "x.pdf")
            open(f1, "w").close()
            f2 = m._resolve_collision_path(f1)
            if f2 != f1 and "(1)" in f2:
                R.add("_resolve_collision_path", "PASS", os.path.basename(f2))
            else:
                R.add("_resolve_collision_path", "FAIL", f2)
        except Exception as e:
            R.add("_resolve_collision_path", "FAIL", str(e))

        # detect_endpoint_kind 매트릭스
        try:
            cases = [
                ({"mcd": "21110", "reportYn": "N", "reportNos": "21110"}, "rule"),
                ({"mcd": "70400", "reportYn": "Y", "reportNos": "70401"}, "pdf+file+dfile"),
                ({"mcd": "10105", "reportYn": "Y", "reportNos": "10105"}, "pdf+file"),
                ({"mcd": "B1210", "reportYn": "N", "reportNos": "B1210"}, "file"),
            ]
            errors = []
            for meta, expected in cases:
                got = m.detect_endpoint_kind(meta)
                if got != expected:
                    errors.append(f"{meta.get('mcd')}: {got}≠{expected}")
            if not errors:
                R.add("detect_endpoint_kind", "PASS", f"{len(cases)}/{len(cases)} 케이스")
            else:
                R.add("detect_endpoint_kind", "FAIL", "; ".join(errors))
        except Exception as e:
            R.add("detect_endpoint_kind", "FAIL", str(e))

        # ─── 5. 다운로드 모듈 (실제 파일 다운로드) ───
        print("\n[5] 다운로드 모듈 (실제 파일)")

        # 5-1. 게시판형 첨부파일 (한산공 감사원 지적사항)
        if result:
            first = result[0]
            try:
                violation_meta = {
                    "report_form_no": first.get("reportFormNo", ""),
                    "table_name": first.get("tableName", ""),
                    "idx_name": first.get("idxName", ""),
                    "idx": first.get("idx", ""),
                    "submission_no": first.get("submissionNo", ""),
                    "bid_type": first.get("bidType", "") or "",
                    "disclosure_no": first.get("disclosureNo", ""),
                    "title": first.get("title", ""),
                }
                attachments = m.fetch_board_attachment_list(sess, TEST_APBA_ID, violation_meta)
                if attachments:
                    R.add("게시판형 첨부 메타 추출", "PASS",
                          f"{len(attachments)}개 ({first.get('title', '')[:30]}...)")
                    # 첫 첨부 다운로드
                    test_dir = m.build_save_path(tmp_root, "감사원 지적사항", TEST_APBA_NM)
                    ok, path, msg = m.download_board_attachment(sess, attachments[0], test_dir)
                    if ok:
                        size = os.path.getsize(path)
                        with open(path, "rb") as f:
                            head = f.read(8)
                        is_valid = head.startswith(b"%PDF") or head.startswith(b"PK") or size > 5000
                        if is_valid:
                            R.add("게시판형 첨부 다운로드", "PASS",
                                  f"{os.path.basename(path)} ({size:,} bytes, head={head[:4]})")
                        else:
                            R.add("게시판형 첨부 다운로드", "FAIL",
                                  f"파일 헤더 비정상: {head}")
                    else:
                        R.add("게시판형 첨부 다운로드", "FAIL", msg)
                else:
                    R.add("게시판형 첨부 메타 추출", "SKIP", "첨부 없음")
            except Exception as e:
                R.add("게시판형 첨부 메타 추출", "FAIL", str(e))
        else:
            R.add("게시판형 첨부 메타 추출", "SKIP", "자료 목록 없음")
            R.add("게시판형 첨부 다운로드", "SKIP", "자료 목록 없음")

        # 5-2. PDF 다운로드 엔드포인트 직접 호출 (보고서형)
        try:
            assert m.ENDPOINT_REGISTRY["pdf"] == "/download/pdf.json"
            assert m.ENDPOINT_REGISTRY["file"] == "/download/file.json"
            assert m.ENDPOINT_REGISTRY["dfile"] == "/download/dfile.json"
            assert m.ENDPOINT_REGISTRY["rule"] == "/download/rulefiledown.json"
            R.add("ENDPOINT_REGISTRY 정의", "PASS", "4종 엔드포인트 일치")
        except (AssertionError, KeyError) as e:
            R.add("ENDPOINT_REGISTRY 정의", "FAIL", str(e))

        # 5-2-2. 신규 jung 항목 실제 PDF 다운로드 (보고서형 신규 항목 동작 보증)
        try:
            # 임원연봉(20501) — 한산공 disclosureNo 안정적
            jung_url = f"{m.BASE_URL}/item/itemOrganListJung.json"
            jung_body = {
                "reportFormRootNo": "20501",
                "apbaType": [], "jidtDptm": [], "area": [],
                "apba_id": "", "pageNo": 1,
            }
            jr = m.retry_request(sess, "POST", jung_url, json=jung_body,
                                 headers={"Content-Type": "application/json;charset=UTF-8"},
                                 timeout=15)
            organs = jr.json().get("data", {}).get("organList", [])
            if isinstance(organs, dict):
                organs = organs.get("result", [])
            c0208 = next((o for o in organs if o.get("apbaId") == TEST_APBA_ID), None)
            if c0208 and c0208.get("disclosureNo"):
                pdf_url = f"{m.BASE_URL}/download/pdf.json?disclosureNo={c0208['disclosureNo']}"
                pr = sess.get(pdf_url, timeout=30, stream=True)
                content = pr.content
                if content[:4] == b"%PDF" and len(content) > 5000:
                    R.add("신규 jung PDF 다운로드", "PASS",
                          f"임원연봉 PDF {len(content):,} bytes, %PDF 헤더")
                else:
                    R.add("신규 jung PDF 다운로드", "FAIL",
                          f"PDF 비정상 (size={len(content)}, head={content[:4]})")
            else:
                R.add("신규 jung PDF 다운로드", "SKIP", "disclosureNo 없음")
        except Exception as e:
            R.add("신규 jung PDF 다운로드", "FAIL", str(e))

        # 5-3. 내부규정 API 응답 (다운로드 흐름의 첫 단계)
        try:
            url = f"{m.BASE_URL}/occasional/findRuleList.json"
            r = sess.get(url, params={
                "type": "apbaNa", "word": TEST_APBA_NM,
                "pageNo": 1, "divis": "K1500"  # 정관
            }, timeout=15)
            d = r.json()
            result_list = d.get("data", {}).get("result", [])
            if result_list:
                R.add("내부규정 findRuleList 호출", "PASS",
                      f"한산공 정관 {len(result_list)}건")
            else:
                R.add("내부규정 findRuleList 호출", "SKIP", "정관 없음")
        except Exception as e:
            R.add("내부규정 findRuleList 호출", "FAIL", str(e))

        # ─── 6. 헤더 버전 일치 ─────────────────
        print("\n[6] 버전 표기")
        with open(TARGET_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        version_checks = [
            ("# 버전: 5.4.2", "헤더 버전 표기"),
            ("ALIO 항목별 공시 크롤링 시스템 v5.4.2", "GUI 타이틀"),
            ("v5.4.2 사용법", "도움말 헤더"),
            ("(v5.4.2)", "로그 메시지"),
        ]
        for needle, label in version_checks:
            if needle in content:
                R.add(label, "PASS", needle)
            else:
                R.add(label, "FAIL", f"미발견: {needle}")

    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    return _final_report(R)


def _final_report(R):
    print("\n" + "=" * 70)
    p, f, s, t = R.summary()
    print(f"결과: PASS={p}, FAIL={f}, SKIP={s} / 총 {t}건")
    print("=" * 70)
    if f > 0:
        print("\n[실패 항목]")
        for name, status, detail in R.results:
            if status == "FAIL":
                print(f"  ✗ {name}: {detail}")
    sys.exit(0 if f == 0 else 1)


if __name__ == "__main__":
    main()
