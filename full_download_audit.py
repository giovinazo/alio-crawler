# -*- coding: utf-8 -*-
"""
v5.4.1 전 항목 실제 다운로드 시뮬레이션 — 한국산업단지공단(C0208) 기준.

각 공시항목마다 다음을 자동으로 시도:
  1. (보고서형) itemOrganListJung.json → 한산공 행 추출 → PDF/files 다운로드
  2. (게시판형) itemReportListSusi.json → 첫 자료 → fetch_board_attachment_list → 다운로드
  3. (특수형) rule, audit, safety 등 — 핵심 API만 호출

결과 표:
  - 항목명, type, kind, 자료 수, PDF 수신, 첨부 수신, 비고
  - 다운로드 0건 항목 별도 리스트

실행: python3 full_download_audit.py
약 5~10분 소요 (83개 × 1~2회 API 호출)
"""
from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import shutil
import urllib3
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

urllib3.disable_warnings()

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET_FILE = os.path.join(HERE, "alio_crawler_v5.4.py")
TEST_APBA_ID = "C0208"
TEST_APBA_NAME = "한국산업단지공단"


def load_module():
    spec = importlib.util.spec_from_file_location("alio_v54", TARGET_FILE)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def probe_jung_item(m, sess, item_meta, root_no, tmpdir):
    """보고서형(jung+Y) — 한산공 행에서 PDF·files 다운로드 시도."""
    url = f"{m.BASE_URL}/item/itemOrganListJung.json"
    body = {
        "reportFormRootNo": root_no,
        "apbaType": [], "jidtDptm": [], "area": [],
        "apba_id": "", "pageNo": 1,
    }
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        d = r.json()
        if d.get("status") != "success":
            return {"records": 0, "pdf": 0, "files": 0, "err": "API status err"}
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}

    organs = d.get("data", {}).get("organList", [])
    if isinstance(organs, dict):
        organs = organs.get("result", [])
    matches = [o for o in organs if o.get("apbaId") == TEST_APBA_ID]
    if not matches:
        return {"records": 0, "pdf": 0, "files": 0, "err": "한산공 매칭 없음"}

    # 첫 매칭 행에서 시도
    pdf_count = 0
    file_count = 0
    for case in matches[:1]:
        dno = case.get("disclosureNo", "")
        # PDF
        if dno:
            try:
                pr = sess.get(f"{m.BASE_URL}/download/pdf.json?disclosureNo={dno}",
                              timeout=20)
                if pr.status_code == 200 and pr.content[:4] == b"%PDF" and len(pr.content) > 1000:
                    pdf_count += 1
            except Exception:
                pass
        # files
        files_str = case.get("files", "")
        files_parsed = m.parse_files_field(files_str) if files_str else []
        for fp in files_parsed[:3]:  # 최대 3개만
            fid = fp.get("id", "")
            if not fid:
                continue
            try:
                fr = sess.get(f"{m.BASE_URL}/download/file.json?f={fid}&d={dno}",
                              timeout=20)
                if fr.status_code == 200 and len(fr.content) > 100:
                    ct = (fr.headers.get("Content-Type") or "").lower()
                    if "html" not in ct:
                        file_count += 1
            except Exception:
                pass

    return {"records": len(matches), "pdf": pdf_count, "files": file_count, "err": ""}


def probe_envlaw_item(m, sess, item_meta, root_no, tmpdir):
    """게시판형(envlaw) — 한산공 첫 자료에서 fetch_board_attachment_list + 다운로드."""
    url = f"{m.BASE_URL}/item/itemReportListSusi.json"
    body = {"pageNo": 1, "apbaId": TEST_APBA_ID, "apbaType": "",
            "reportFormRootNo": root_no,
            "search_word": "", "search_flag": "title",
            "bid_type": "", "enfc_istt": ""}
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        d = r.json()
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}

    result_list = d.get("data", {}).get("result", []) if d.get("data") else []
    if not result_list:
        return {"records": 0, "pdf": 0, "files": 0, "err": "한산공 자료 없음"}

    first = result_list[0]
    vmeta = {
        "report_form_no": first.get("reportFormNo", ""),
        "table_name": first.get("tableName", ""),
        "idx_name": first.get("idxName", ""),
        "idx": first.get("idx", ""),
        "submission_no": first.get("submissionNo", ""),
        "bid_type": first.get("bidType", "") or "",
        "disclosure_no": first.get("disclosureNo", ""),
    }

    file_count = 0
    pattern_kind = ""
    try:
        attachments = m.fetch_board_attachment_list(sess, TEST_APBA_ID, vmeta)
        if attachments:
            pattern_kind = attachments[0].get("kind", "")
            # 첫 첨부 1개만 다운로드 시도
            ok, _, _ = m.download_board_attachment(sess, attachments[0], tmpdir)
            if ok:
                file_count = 1
    except Exception as e:
        return {"records": len(result_list), "pdf": 0, "files": 0,
                "err": f"attach fetch err: {e}"}

    return {"records": len(result_list), "pdf": 0, "files": file_count,
            "err": "" if file_count else f"첨부 추출 0개" if not attachments else "추출했으나 다운로드 실패",
            "kind": pattern_kind}


def probe_special_item(m, sess, item_meta, root_no, item_type, tmpdir):
    """특화 type(rule/audit/safety/mgmt_eval/integrity/discipline/general) — 핵심 API만 호출."""
    if item_type == "rule":
        # 내부규정: findRuleList.json
        try:
            r = sess.get(f"{m.BASE_URL}/occasional/findRuleList.json",
                         params={"type": "apbaNa", "word": TEST_APBA_NAME,
                                 "pageNo": 1, "divis": "K1500"}, timeout=15)
            d = r.json()
            cnt = len(d.get("data", {}).get("result", []))
            return {"records": cnt, "pdf": 0, "files": 0, "err": ""}
        except Exception as e:
            return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}
    # audit, safety, mgmt_eval 등 — itemOrganListJung 호출만
    return probe_jung_item(m, sess, item_meta, root_no, tmpdir)


def main():
    m = load_module()
    sess = m.create_session(verify_ssl=False)
    sess.headers.update(m.HEADERS)
    items = m.get_alio_items()
    if not items:
        print("ERROR: 캐시 없음")
        sys.exit(1)
    print(f"{len(items)}개 항목 점검 시작 ({TEST_APBA_NAME} 기준)\n")

    to_legacy = m.ALIOCrawlerGUI._item_meta_to_legacy.__get__(type("X", (), {})())

    tmpdir = tempfile.mkdtemp(prefix="alio_full_dl_")

    def probe_one(it):
        name = m.build_item_display_name(it)
        rn = m.build_item_root_no(it)
        legacy = to_legacy(it)
        type_ = legacy.get("type", "?")
        yn = it.get("reportYn", "")
        if type_ in ("rule", "audit", "safety", "mgmt_eval", "integrity",
                    "discipline", "general"):
            r = probe_special_item(m, sess, it, rn, type_, tmpdir)
        elif type_ == "envlaw":
            r = probe_envlaw_item(m, sess, it, rn, tmpdir)
        else:  # jung
            r = probe_jung_item(m, sess, it, rn, tmpdir)
        return {
            "name": name, "rn": rn, "lcdnm": it.get("lcdnm", ""),
            "type": type_, "yn": yn, **r,
        }

    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(probe_one, it) for it in items]
        for i, fut in enumerate(as_completed(futs)):
            results.append(fut.result())
            if (i + 1) % 10 == 0:
                print(f"  진행: {i+1}/{len(items)}")
    shutil.rmtree(tmpdir, ignore_errors=True)

    # 출력 1: 전체 표
    print("\n" + "=" * 110)
    print(f'{"항목명":35s} {"type":10s} {"yn":3s} {"records":>7s} {"pdf":>4s} {"files":>5s} {"비고"}')
    print("=" * 110)
    for r in sorted(results, key=lambda x: (x["type"], x["name"])):
        nm = r["name"][:34]
        bigo = r.get("err", "")
        if r.get("kind"):
            bigo = f"kind={r['kind']} {bigo}".strip()
        print(f'{nm:35s} {r["type"]:10s} {r["yn"]:3s} '
              f'{r["records"]:>7d} {r["pdf"]:>4d} {r["files"]:>5d} {bigo[:40]}')

    # 출력 2: 문제 항목 식별
    print("\n" + "=" * 110)
    print("[문제 항목 - 한산공 자료 있는데 0건 다운로드]")
    print("=" * 110)
    issues = [r for r in results
              if r["records"] > 0 and r["pdf"] == 0 and r["files"] == 0]
    if not issues:
        print("  (없음)")
    else:
        for r in issues:
            print(f'  ⚠ {r["name"][:30]:30s} type={r["type"]:8s} '
                  f'rn={r["rn"]:18s} 자료{r["records"]}건  err={r.get("err", "")}')

    # 출력 3: 통계
    print("\n" + "=" * 110)
    print("[통계]")
    print("=" * 110)
    print(f'  전체: {len(results)}개')
    print(f'  자료 있음: {sum(1 for r in results if r["records"] > 0)}개')
    print(f'  자료 없음: {sum(1 for r in results if r["records"] == 0)}개')
    print(f'  PDF 다운로드 성공: {sum(1 for r in results if r["pdf"] > 0)}개')
    print(f'  첨부파일 다운로드 성공: {sum(1 for r in results if r["files"] > 0)}개')
    print(f'  문제 항목(자료 있는데 0 받음): {len(issues)}개')

    # type별 통계
    print("\n  [type별 결과]")
    by_type = {}
    for r in results:
        by_type.setdefault(r["type"], []).append(r)
    for t, lst in sorted(by_type.items()):
        with_data = [r for r in lst if r["records"] > 0]
        downloaded = [r for r in with_data if r["pdf"] > 0 or r["files"] > 0]
        print(f'    {t:10s}: 전체 {len(lst):3d}개, 자료있음 {len(with_data):3d}개, '
              f'다운로드성공 {len(downloaded):3d}개')


if __name__ == "__main__":
    main()
