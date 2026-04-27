# -*- coding: utf-8 -*-
"""
v5.4.1 정밀 probe — fetch_disclosures default 분기 흐름을 정확히 재현하여
모든 항목의 실제 다운로드 가능 여부 자동 판정.

이전 full_download_audit.py의 한계:
- disclosureNo 없는 경우 itemReportListSusi로 추출 시도하는 fetch_disclosures 로직 미포함
- audit, mgmt_eval, safety 등 특화 type별 흐름 단순화

이 스크립트는 type별로 정확한 흐름을 시뮬레이션:
- jung: disclosureNo 추출 fallback → PDF + files
- audit/mgmt_eval: itemReportFiles.json + dfile.json
- envlaw: 게시판형 첨부파일 (kind=upload/fileno)
- rule: findRuleList + rulefiledown
- safety: PDF + files + 안전경영보고서(dfile)
- general/integrity/discipline: itemOrganListJung 응답 + PDF

실행: python3 precise_audit.py [--items=20305,31921,...]
"""
from __future__ import annotations

import argparse
import importlib.util
import os
import sys
import tempfile
import shutil
import urllib3
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


def jung_precise(m, sess, root_no):
    """fetch_disclosures default 분기 흐름 그대로:
    1. itemOrganListJung → 한산공 행 → disclosureNo, files 추출
    2. disclosureNo 없으면 itemReportListSusi로 추출 시도
    3. PDF 다운로드 시도
    4. files 다운로드 시도
    """
    url = f"{m.BASE_URL}/item/itemOrganListJung.json"
    body = {"reportFormRootNo": root_no, "apbaType": [], "jidtDptm": [],
            "area": [], "apba_id": "", "pageNo": 1}
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        d = r.json()
        if d.get("status") != "success":
            return {"records": 0, "pdf": 0, "files": 0, "files_total": 0, "err": "API err"}
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "files_total": 0, "err": str(e)}

    organs = d.get("data", {}).get("organList", [])
    if isinstance(organs, dict):
        organs = organs.get("result", [])
    matches = [o for o in organs if o.get("apbaId") == TEST_APBA_ID]
    if not matches:
        return {"records": 0, "pdf": 0, "files": 0, "files_total": 0, "err": "한산공 매칭 없음"}

    case = matches[0]
    apba_type = case.get("apbaType", "")
    dno = case.get("disclosureNo", "") or ""
    files_str = case.get("files", "") or ""

    # disclosureNo 없으면 fetch_disclosures default와 동일하게 추출 시도
    # (v5.4.1 패치: 콤마 분리된 rootNos는 itemReportListSusi가 거부하므로 분기별 순차 시도)
    if not dno:
        rn_candidates = [x.strip() for x in (root_no or "").split(",") if x.strip()]
        for try_rn in rn_candidates:
            try:
                susi_url = f"{m.BASE_URL}/item/itemReportListSusi.json"
                susi_body = {"pageNo": 1, "apbaId": TEST_APBA_ID, "apbaType": apba_type,
                             "reportFormRootNo": try_rn,
                             "search_word": "", "search_flag": "title",
                             "bid_type": "", "enfc_istt": ""}
                sr = sess.post(susi_url, json=susi_body,
                               headers={"Content-Type": "application/json;charset=UTF-8"},
                               timeout=15)
                if sr.status_code != 200:
                    continue
                sd = sr.json()
                if sd.get("status") != "success":
                    continue
                rl = sd.get("data", {}).get("result", []) or []
                if not rl:
                    continue
                dno = str(rl[0].get("disclosureNo", "") or "")
                if dno:
                    break
            except Exception:
                continue

    # PDF 다운로드
    pdf_ok = 0
    if dno:
        try:
            pr = sess.get(f"{m.BASE_URL}/download/pdf.json?disclosureNo={dno}", timeout=30)
            if pr.status_code == 200:
                ct = pr.headers.get("Content-Type", "").lower()
                if "pdf" in ct or (len(pr.content) > 100 and pr.content[:4] == b"%PDF"):
                    pdf_ok = 1
        except Exception:
            pass

    # files 다운로드 (전체 시도)
    files_parsed = m.parse_files_field(files_str) if files_str else []
    files_total = len(files_parsed)
    files_ok = 0
    for fp in files_parsed[:3]:
        fid = fp.get("id", "")
        if not fid or not dno:
            continue
        try:
            fr = sess.get(f"{m.BASE_URL}/download/file.json?f={fid}&d={dno}", timeout=20)
            if fr.status_code == 200 and len(fr.content) > 100:
                ct = (fr.headers.get("Content-Type") or "").lower()
                if "html" not in ct:
                    files_ok += 1
        except Exception:
            pass

    return {
        "records": 1, "pdf": pdf_ok, "files": files_ok,
        "files_total": files_total,
        "err": "" if (pdf_ok or files_ok) else "PDF·files 모두 0",
        "dno": dno or "(없음)",
    }


def audit_precise(m, sess, root_no):
    """audit type: itemReportFiles.json + dfile.json"""
    # 1) itemOrganListJung로 disclosureNo, submissionNo 추출
    url = f"{m.BASE_URL}/item/itemOrganListJung.json"
    body = {"reportFormRootNo": root_no, "apbaType": [], "jidtDptm": [],
            "area": [], "apba_id": "", "pageNo": 1}
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        organs = r.json().get("data", {}).get("organList", [])
        if isinstance(organs, dict):
            organs = organs.get("result", [])
        matches = [o for o in organs if o.get("apbaId") == TEST_APBA_ID]
        if not matches:
            return {"records": 0, "pdf": 0, "files": 0, "err": "한산공 매칭 없음"}
        case = matches[0]
        dno = case.get("disclosureNo", "") or ""
        sno = case.get("submissionNo", "") or ""
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}

    if not dno or not sno:
        return {"records": 1, "pdf": 0, "files": 0, "err": "dno/sno 없음"}

    # 2) itemReportFiles.json 호출
    try:
        fr = sess.get(f"{m.BASE_URL}/item/itemReportFiles.json?disclosureNo={dno}", timeout=15)
        if fr.status_code != 200:
            return {"records": 1, "pdf": 0, "files": 0, "err": f"files API HTTP {fr.status_code}"}
        files_data = fr.json().get("data", []) or []
        if not files_data:
            return {"records": 1, "pdf": 0, "files": 0, "err": "itemReportFiles.json 빈 응답"}
    except Exception as e:
        return {"records": 1, "pdf": 0, "files": 0, "err": f"files API err: {e}"}

    # 3) dfile.json 다운로드 (1개만)
    from urllib.parse import quote
    files_ok = 0
    for fi in files_data[:1]:
        fname = fi.get("orcpFileNa", "")
        if not fname:
            continue
        try:
            dr = sess.get(f"{m.BASE_URL}/download/dfile.json?fileName={quote(fname, safe='')}&submissionNo={sno}",
                          timeout=30)
            if dr.status_code == 200 and len(dr.content) > 100:
                ct = (dr.headers.get("Content-Type") or "").lower()
                if "html" not in ct:
                    files_ok += 1
        except Exception:
            pass

    return {"records": 1, "pdf": 0, "files": files_ok,
            "err": "" if files_ok else "dfile.json 다운로드 실패",
            "files_total": len(files_data)}


def envlaw_precise(m, sess, root_no):
    """envlaw type: itemReportListSusi → fetch_board_attachment_list → 다운로드"""
    url = f"{m.BASE_URL}/item/itemReportListSusi.json"
    body = {"pageNo": 1, "apbaId": TEST_APBA_ID, "apbaType": "",
            "reportFormRootNo": root_no,
            "search_word": "", "search_flag": "title",
            "bid_type": "", "enfc_istt": ""}
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        result_list = r.json().get("data", {}).get("result", []) or []
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}

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
    tmpdir = tempfile.mkdtemp(prefix="prec_env_")
    try:
        atts = m.fetch_board_attachment_list(sess, TEST_APBA_ID, vmeta)
        if not atts:
            return {"records": len(result_list), "pdf": 0, "files": 0,
                    "err": "첨부 추출 0개"}
        ok, _, msg = m.download_board_attachment(sess, atts[0], tmpdir)
        return {"records": len(result_list), "pdf": 0, "files": 1 if ok else 0,
                "err": "" if ok else msg, "kind": atts[0].get("kind", ""),
                "atts_total": len(atts)}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def rule_precise(m, sess, _root_no):
    """rule type: findRuleList + rulefiledown"""
    try:
        # 정관(K1500) 1건 조회
        r = sess.get(f"{m.BASE_URL}/occasional/findRuleList.json",
                     params={"type": "apbaNa", "word": TEST_APBA_NAME,
                             "pageNo": 1, "divis": "K1500"}, timeout=15)
        rules = r.json().get("data", {}).get("result", []) or []
        if not rules:
            return {"records": 0, "pdf": 0, "files": 0, "err": "내부규정 없음"}

        first = rules[0]
        # files 필드에서 fileNo 추출 시도
        files_str = first.get("files", "") or ""
        files_parsed = m.parse_files_field(files_str)
        if not files_parsed:
            return {"records": len(rules), "pdf": 0, "files": 0, "err": "files 필드 비어있음"}

        # rulefiledown.json 다운로드
        first_file = files_parsed[0]
        fid = first_file.get("id", "")
        if not fid:
            return {"records": len(rules), "pdf": 0, "files": 0, "err": "fileNo 없음"}
        rr = sess.get(f"{m.BASE_URL}/download/rulefiledown.json?fileNo={fid}", timeout=20)
        if rr.status_code == 200 and len(rr.content) > 100:
            ct = (rr.headers.get("Content-Type") or "").lower()
            if "html" not in ct:
                return {"records": len(rules), "pdf": 0, "files": 1, "err": ""}
        return {"records": len(rules), "pdf": 0, "files": 0,
                "err": f"rulefiledown HTTP {rr.status_code}"}
    except Exception as e:
        return {"records": 0, "pdf": 0, "files": 0, "err": str(e)}


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--items", type=str, default="",
                   help="콤마 구분 mcd 또는 항목명 키워드 (선택)")
    args = p.parse_args()

    m = load_module()
    sess = m.create_session(verify_ssl=True)
    sess.headers.update(m.HEADERS)
    items = m.get_alio_items()

    # 필터
    if args.items:
        kws = [x.strip() for x in args.items.split(",") if x.strip()]
        filtered = []
        for it in items:
            if (it.get("mcd", "") in kws or
                m.build_item_root_no(it) in kws or
                any(kw in m.build_item_display_name(it) for kw in kws)):
                filtered.append(it)
        items = filtered
        print(f'필터 적용: {len(items)}개 항목')
    else:
        print(f'전체 {len(items)}개 항목 점검')

    to_legacy = m.ALIOCrawlerGUI._item_meta_to_legacy.__get__(type("X", (), {})())

    def probe_one(it):
        name = m.build_item_display_name(it)
        rn = m.build_item_root_no(it)
        legacy = to_legacy(it)
        type_ = legacy.get("type", "?")
        if type_ == "envlaw":
            r = envlaw_precise(m, sess, rn)
        elif type_ == "audit":
            r = audit_precise(m, sess, rn)
        elif type_ == "mgmt_eval":
            r = audit_precise(m, sess, rn)  # 동일 흐름
        elif type_ == "rule":
            r = rule_precise(m, sess, rn)
        else:
            r = jung_precise(m, sess, rn)
        return {"name": name, "rn": rn, "type": type_,
                "lcdnm": it.get("lcdnm", ""), **r}

    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(probe_one, it) for it in items]
        for i, fut in enumerate(as_completed(futs)):
            results.append(fut.result())
            if (i + 1) % 10 == 0:
                print(f"  진행: {i+1}/{len(items)}")

    # 결과 출력
    print("\n" + "=" * 110)
    print(f'{"항목명":35s} {"type":10s} {"records":>5s} {"pdf":>4s} {"files":>5s} {"비고"}')
    print("=" * 110)
    for r in sorted(results, key=lambda x: (x["type"], x["name"])):
        bigo = r.get("err", "")
        if r.get("kind"):
            bigo = f"kind={r['kind']} {bigo}".strip()
        if r.get("files_total") and not bigo:
            bigo = f"(files {r['files']}/{r['files_total']})"
        print(f'{r["name"][:34]:35s} {r["type"]:10s} '
              f'{r["records"]:>5d} {r["pdf"]:>4d} {r["files"]:>5d} {bigo[:40]}')

    # 문제 항목
    print("\n" + "=" * 110)
    print("[문제 항목 — 자료 있는데 PDF·files 모두 0]")
    print("=" * 110)
    issues = [r for r in results
              if r["records"] > 0 and r["pdf"] == 0 and r["files"] == 0]
    if issues:
        for r in issues:
            print(f'  ⚠ {r["name"][:30]:30s} type={r["type"]:10s} rn={r["rn"]:18s} '
                  f'err={r.get("err", "")}')
    else:
        print("  (없음)")

    # 통계
    print("\n[통계]")
    success = sum(1 for r in results if r["pdf"] > 0 or r["files"] > 0)
    no_data = sum(1 for r in results if r["records"] == 0)
    issue_n = len(issues)
    print(f"  전체: {len(results)}개 / 다운로드 성공: {success}개 / "
          f"자료 없음: {no_data}개 / 문제: {issue_n}개")


if __name__ == "__main__":
    main()
