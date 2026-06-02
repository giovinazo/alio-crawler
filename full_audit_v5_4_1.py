# -*- coding: utf-8 -*-
"""
v5.4.1 전 항목(83개) 동작 시뮬레이션 분석.

목적: 신규 72개 항목이 실제로 어떤 type으로 처리되며 search_organs / fetch_disclosures
      흐름과 정합성이 있는지 검증.

각 항목에 대해:
  1) _item_meta_to_legacy() 매핑 결과 type
  2) 한국산업단지공단(C0208) 기준 itemOrganListJung 응답 (보고서형)
  3) 한국산업단지공단(C0208) 기준 itemReportListSusi 응답 (게시판형)
  4) 항목별 잠재 이슈 라벨링

실행: python3 full_audit_v5_4_1.py
"""
from __future__ import annotations

import importlib.util
import os
import sys
import time
from collections import defaultdict
import urllib3

urllib3.disable_warnings()

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET_FILE = os.path.join(HERE, "alio_crawler_v5.4.py")
TEST_APBA_ID = "C0208"


def load_module():
    spec = importlib.util.spec_from_file_location("alio_v54", TARGET_FILE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def call_jung(sess, m, root_no):
    """itemOrganListJung 호출 → 한산공 매칭 자료 메타 반환"""
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
            return None, f"status={d.get('status')}"
        organs = d.get("data", {}).get("organList", [])
        if isinstance(organs, dict):
            organs = organs.get("result", [])
        # C0208 매칭
        c0208 = [o for o in organs if o.get("apbaId") == TEST_APBA_ID]
        return {
            "total": len(organs),
            "c0208_match": len(c0208),
            "first_keys": list(organs[0].keys()) if organs else [],
        }, None
    except Exception as e:
        return None, str(e)


def call_susi(sess, m, root_no):
    """itemReportListSusi (한산공 자료 목록)"""
    url = f"{m.BASE_URL}/item/itemReportListSusi.json"
    body = {
        "pageNo": 1, "apbaId": TEST_APBA_ID, "apbaType": "",
        "reportFormRootNo": root_no,
        "search_word": "", "search_flag": "title",
        "bid_type": "", "enfc_istt": "",
    }
    try:
        r = m.retry_request(sess, "POST", url, json=body,
                            headers={"Content-Type": "application/json;charset=UTF-8"},
                            timeout=15)
        d = r.json()
        if d.get("status") != "success":
            return None, f"status={d.get('status')}"
        result = d.get("data", {}).get("result", [])
        return {
            "count": len(result),
            "first_disclosure_no": result[0].get("disclosureNo") if result else "",
            "first_report_form_no": result[0].get("reportFormNo") if result else "",
        }, None
    except Exception as e:
        return None, str(e)


def main():
    print("=" * 90)
    print(f"v5.4.1 전 항목 동작 시뮬레이션 분석 (대상: {TEST_APBA_ID} 한국산업단지공단)")
    print("=" * 90)

    m = load_module()
    sess = m.create_session(verify_ssl=False)
    sess.headers.update(m.HEADERS)
    items = m.get_alio_items()
    if not items:
        print("ERROR: 항목 캐시 비어있음")
        sys.exit(1)
    print(f"총 항목: {len(items)}개\n")

    # _item_meta_to_legacy 호출용
    to_legacy = m.ALIOCrawlerGUI._item_meta_to_legacy.__get__(type("X", (), {})())

    # 결과 누적
    rows = []  # (대분류, 항목명, rootNo, reportYn, type, jung_result, susi_result, issue)
    for it in items:
        legacy = to_legacy(it)
        name = m.build_item_display_name(it)
        rn = m.build_item_root_no(it)
        yn = it.get("reportYn", "")
        type_ = legacy.get("type", "?")
        kind_ = legacy.get("_kind", "?")
        lcdnm = it.get("lcdnm", "?")

        # API 호출 (둘 다)
        jung, jung_err = call_jung(sess, m, rn)
        time.sleep(0.05)
        susi, susi_err = call_susi(sess, m, rn)
        time.sleep(0.05)

        # 이슈 판정
        issue = []
        # 1) v5.3 등록 항목 → 안전 (기존 type 그대로)
        is_v53_known = name in m.DISCLOSURE_ITEMS or any(
            li.get("rootNo") == rn for li in m.DISCLOSURE_ITEMS.values()
        )

        # 2) 보고서형(Y)인데 jung으로 매핑되었지만 itemOrganListJung 응답 0
        if not is_v53_known:
            if yn == "Y" and type_ == "jung":
                if jung is None or jung.get("total", 0) == 0:
                    issue.append("Y/jung-API응답0")
            elif yn == "N" and type_ == "envlaw":
                # 게시판형 → envlaw 흐름. itemReportListSusi 응답 확인
                if susi is None:
                    issue.append("N/envlaw-API오류")

        # 3) jung 매핑인데 disclosureNo가 안 잡힐 가능성
        if not is_v53_known and type_ == "jung" and yn == "Y":
            if jung and jung.get("c0208_match", 0) == 0:
                issue.append("한산공-매칭없음")

        rows.append({
            "lcdnm": lcdnm,
            "name": name,
            "rn": rn,
            "yn": yn,
            "type": type_,
            "kind": kind_,
            "is_v53": is_v53_known,
            "jung": jung,
            "jung_err": jung_err,
            "susi": susi,
            "susi_err": susi_err,
            "issue": issue,
        })
        print(".", end="", flush=True)
    print()

    # 출력 1: type 분포
    print("\n" + "=" * 90)
    print("[1] 매핑 type 분포")
    print("=" * 90)
    type_count = defaultdict(int)
    type_v53 = defaultdict(int)
    type_new = defaultdict(int)
    for r in rows:
        type_count[r["type"]] += 1
        if r["is_v53"]:
            type_v53[r["type"]] += 1
        else:
            type_new[r["type"]] += 1
    print(f"{'type':12s} {'전체':>6s} {'v5.3':>6s} {'신규':>6s}")
    for t in sorted(type_count):
        print(f"{t:12s} {type_count[t]:>6d} {type_v53[t]:>6d} {type_new[t]:>6d}")

    # 출력 2: 잠재 이슈 항목
    print("\n" + "=" * 90)
    print("[2] 잠재 이슈 항목")
    print("=" * 90)
    issued = [r for r in rows if r["issue"]]
    print(f"이슈 발견: {len(issued)}개\n")
    for r in issued:
        print(f"  [{','.join(r['issue']):20s}] {r['lcdnm']:15s} | {r['name'][:40]:40s} "
              f"rn={r['rn']:18s} yn={r['yn']} type={r['type']}")
        if r["jung"]:
            print(f"    jung: total={r['jung']['total']}, c0208={r['jung']['c0208_match']}")
        if r["jung_err"]:
            print(f"    jung_err: {r['jung_err']}")
        if r["susi"]:
            print(f"    susi: count={r['susi']['count']}, dno={r['susi']['first_disclosure_no']}")
        if r["susi_err"]:
            print(f"    susi_err: {r['susi_err']}")

    # 출력 3: type별 검증 — 보고서형 신규 jung 항목 동작 확인
    print("\n" + "=" * 90)
    print("[3] 신규 보고서형(jung) 항목 한산공 자료 매칭")
    print("=" * 90)
    cnt_match = 0
    cnt_nomatch = 0
    cnt_nodata = 0
    for r in rows:
        if not r["is_v53"] and r["type"] == "jung" and r["yn"] == "Y":
            if r["jung"] is None:
                cnt_nodata += 1
            elif r["jung"]["c0208_match"] > 0:
                cnt_match += 1
            else:
                cnt_nomatch += 1
    total_new_jung_y = cnt_match + cnt_nomatch + cnt_nodata
    print(f"  전체 신규 jung+Y 항목: {total_new_jung_y}개")
    print(f"  ✓ 한산공 매칭: {cnt_match}")
    print(f"  ⚠ 한산공 미매칭(데이터 자체는 있음): {cnt_nomatch}")
    print(f"  ✗ API 오류/응답없음: {cnt_nodata}")

    # 출력 4: 신규 게시판형(envlaw) 자료 유무
    print("\n" + "=" * 90)
    print("[4] 신규 게시판형(envlaw) 항목 한산공 자료 유무")
    print("=" * 90)
    has_data = 0
    no_data = 0
    err_count = 0
    for r in rows:
        if not r["is_v53"] and r["type"] == "envlaw":
            if r["susi"] is None:
                err_count += 1
            elif r["susi"]["count"] > 0:
                has_data += 1
            else:
                no_data += 1
    total_new_envlaw = has_data + no_data + err_count
    print(f"  전체 신규 envlaw 항목: {total_new_envlaw}개")
    print(f"  ✓ 한산공 자료 있음: {has_data}")
    print(f"  — 한산공 자료 없음: {no_data}")
    print(f"  ✗ API 오류: {err_count}")

    # 출력 5: 분기별 핵심 항목 5개씩 샘플
    print("\n" + "=" * 90)
    print("[5] type별 신규 항목 샘플 5개 (한산공 자료 있는 것 우선)")
    print("=" * 90)
    by_type = defaultdict(list)
    for r in rows:
        if not r["is_v53"]:
            by_type[r["type"]].append(r)
    for t, lst in sorted(by_type.items()):
        # 자료 있는 것 우선 정렬
        def has_d(r):
            if r["yn"] == "Y" and r["jung"]:
                return r["jung"]["c0208_match"]
            elif r["susi"]:
                return r["susi"]["count"]
            return 0
        lst.sort(key=has_d, reverse=True)
        print(f"\n  [{t}] 총 {len(lst)}개")
        for r in lst[:5]:
            d = has_d(r)
            mark = "✓" if d > 0 else "—"
            print(f"    {mark} {r['name'][:30]:30s} rn={r['rn']:18s} yn={r['yn']} (한산공 {d})")

    # 결론 요약
    print("\n" + "=" * 90)
    print("[결론]")
    print("=" * 90)
    print(f"  v5.3 등록 항목: {sum(1 for r in rows if r['is_v53'])}개")
    print(f"  v5.4 신규 항목: {sum(1 for r in rows if not r['is_v53'])}개")
    print(f"  잠재 이슈 항목: {len(issued)}개")
    print(f"  신규 보고서형(jung+Y) 한산공 매칭: {cnt_match}/{total_new_jung_y}")
    print(f"  신규 게시판형(envlaw) 한산공 자료 있음: {has_data}/{total_new_envlaw}")


if __name__ == "__main__":
    main()
