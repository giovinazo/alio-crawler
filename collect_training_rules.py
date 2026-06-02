"""공공기관 교육훈련규정 일괄 수집 스크립트.

알리오(alio.go.kr) "내부규정" API 3-step을 직접 호출하여 키워드(기본 "교육훈련")가
제목에 포함된 규정 파일을 기관별로 수집한다.

사용:
    python3 collect_training_rules.py --mode kicox          # KICOX 단독 검증
    python3 collect_training_rules.py --mode all            # 전체 약 350개 기관

산출물 (--out 미지정 시 NAS/04_법률_규정/공공기관_교육훈련규정_비교_YYYYMMDD/):
    {기관명}/{규정파일}
    수집현황.csv
    failed_log.csv
    processed_inst.json   (체크포인트, 재개용)
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Optional

import requests
import urllib3
from requests.adapters import HTTPAdapter

# ALIO는 일부 외부망의 SSL 검사(가로채기) 보안장비 뒤에 있어 검증 시 인증서 오류 발생 → 검증 끄고 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://www.alio.go.kr"

# 출력 기본 경로: 이식성을 위해 홈 디렉터리 기준. ALIO_RULES_OUT 환경변수로 변경 가능.
DEFAULT_OUT_BASE = Path(
    os.environ.get("ALIO_RULES_OUT", str(Path.home() / "alio_training_rules"))
)

HEADERS_GET = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
    "Referer": f"{BASE_URL}/item/itemList.do",
}

HEADERS_POST_JSON = {
    **HEADERS_GET,
    "Content-Type": "application/json;charset=UTF-8",
}

_INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WHITESPACE_RUN = re.compile(r"\s+")
_RETRIABLE_STATUS = {429, 500, 502, 503, 504}
_RETRIABLE_EXC = (requests.ConnectionError, requests.Timeout)
_TIMEOUT = (5, 60)
_PER_PAGE = 10  # 알리오 findRuleList 페이지당 건수

# 분류 코드 → 표준 한글명 (insdRuleDivis가 코드로 올 수도, 한글로 올 수도 있음)
DIVIS_NAME_MAP = {
    "K1500": "정관",
    "K1100": "인사·복무·징계",
    "K1200": "보수",
    "K1300": "직제",
    "K1400": "기타",
}


def sanitize_filename(name: str, max_len: int = 80) -> str:
    if not name:
        return "untitled"
    name = _INVALID_CHARS.sub("", name)
    name = _WHITESPACE_RUN.sub(" ", name).strip()
    name = name.strip(". ")
    if not name:
        return "untitled"
    base, ext = os.path.splitext(name)
    remaining = max_len - len(ext)
    if remaining < 1:
        remaining = 1
    return base[:remaining] + ext


def create_session() -> requests.Session:
    sess = requests.Session()
    sess.verify = False  # SSL 검사(가로채기) 보안장비 환경 대응
    sess.headers.update(HEADERS_GET)
    adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20)
    sess.mount("https://", adapter)
    sess.mount("http://", adapter)
    return sess


def retry_request(
    session: requests.Session,
    method: str,
    url: str,
    *,
    max_retries: int = 3,
    backoff: float = 1.0,
    **kwargs,
) -> requests.Response:
    kwargs.setdefault("timeout", _TIMEOUT)
    for attempt in range(max_retries + 1):
        try:
            resp = session.request(method, url, **kwargs)
            if resp.status_code not in _RETRIABLE_STATUS:
                return resp
            wait = backoff * (2 ** attempt)
            if resp.status_code == 429:
                ra = resp.headers.get("Retry-After")
                if ra:
                    try:
                        wait = float(ra)
                    except ValueError:
                        pass
            if attempt < max_retries:
                time.sleep(wait + random.uniform(0, wait * 0.5))
                continue
            return resp
        except _RETRIABLE_EXC:
            if attempt < max_retries:
                wait = backoff * (2 ** attempt)
                time.sleep(wait + random.uniform(0, wait * 0.5))
            else:
                raise


# ── 알리오 API 래퍼 ────────────────────────────────


def fetch_institutions(session: requests.Session) -> list[dict]:
    """전체 공공기관 목록을 알리오에서 가져온다."""
    url = f"{BASE_URL}/organ/findOrganApbaList.json"
    body = {"apbaType": [], "jidtDptm": [], "area": [], "apba_id": "", "pageNo": 1}
    resp = retry_request(session, "POST", url, json=body, headers=HEADERS_POST_JSON)
    resp.raise_for_status()
    data = resp.json()
    organ_list = data.get("data", {}).get("organList", {})
    total_page = organ_list.get("page", {}).get("totalPage", 1)

    insts: dict[str, dict] = {}

    def _collect(result: list[dict]) -> None:
        for it in result:
            name = it.get("apbaNa", "")
            if not name:
                continue
            insts[name] = {
                "apbaNa": name,
                "apbaId": it.get("apbaId", ""),
                "typeNa": it.get("typeNa", ""),
                "jidtNa": it.get("jidtNa", ""),
            }

    _collect(organ_list.get("result", []))

    if total_page > 1:
        def _page(pn: int) -> list[dict]:
            b = dict(body, pageNo=pn)
            r = retry_request(session, "POST", url, json=b, headers=HEADERS_POST_JSON)
            if r.status_code != 200:
                return []
            return (
                r.json().get("data", {}).get("organList", {}).get("result", []) or []
            )

        with ThreadPoolExecutor(max_workers=5) as ex:
            futures = [ex.submit(_page, p) for p in range(2, total_page + 1)]
            for fu in as_completed(futures):
                _collect(fu.result())

    return sorted(insts.values(), key=lambda x: x["apbaNa"])


def fetch_rule_list(session: requests.Session, inst_name: str) -> list[dict]:
    """기관의 모든 내부규정 목록 (페이지 순회)."""
    url = f"{BASE_URL}/occasional/findRuleList.json"
    rules: list[dict] = []
    page_no = 1
    total_cnt: Optional[int] = None
    while True:
        params = {
            "type": "apbaNa",
            "word": inst_name,
            "pageNo": page_no,
            "divis": "",
        }
        resp = retry_request(session, "GET", url, params=params)
        if resp.status_code != 200:
            break
        try:
            data = resp.json()
        except json.JSONDecodeError:
            break
        d = data.get("data", {})
        if total_cnt is None:
            total_cnt = int(d.get("totalCnt", 0) or 0)
            if total_cnt == 0:
                return []
        page_rules = d.get("result", []) or []
        if not page_rules:
            break
        rules.extend(page_rules)
        if page_no * _PER_PAGE >= total_cnt:
            break
        page_no += 1
        time.sleep(0.3)
    return rules


def fetch_rule_files(session: requests.Session, seq: str) -> list[tuple[int, str]]:
    """findRuleDtl로 bFiles 파싱. 반환: [(file_no:int, file_name:str), ...]."""
    url = f"{BASE_URL}/occasional/findRuleDtl.json"
    resp = retry_request(session, "GET", url, params={"seq": seq})
    if resp.status_code != 200:
        return []
    try:
        data = resp.json()
    except json.JSONDecodeError:
        return []
    b_files = (data.get("data") or {}).get("bFiles") or ""
    if not b_files:
        return []
    out: list[tuple[int, str]] = []
    # 알리오 rule API 응답: "fileNo|파일명,fileNo|파일명"  (라인 4415-4421 참조)
    # 일부 다른 메뉴는 "fileNo@파일명|fileNo@파일명" 포맷. 두 패턴 모두 처리.
    if "@" in b_files and "|" in b_files and "," not in b_files:
        # @ 패턴: | 로 entries 분리, @ 로 id/name
        for ent in b_files.split("|"):
            ent = ent.strip()
            if "@" in ent:
                fid, fname = ent.split("@", 1)
                try:
                    out.append((int(fid.strip()), fname.strip()))
                except ValueError:
                    pass
    else:
        # 표준 rule 패턴: , 로 entries 분리, | 로 id/name
        for ent in b_files.split(","):
            ent = ent.strip()
            if "|" in ent:
                fid, fname = ent.split("|", 1)
                try:
                    out.append((int(fid.strip()), fname.strip()))
                except ValueError:
                    pass
    return out


def download_rule_file(
    session: requests.Session, file_no: int
) -> tuple[Optional[bytes], int]:
    """rulefiledown.json으로 파일 바이너리 다운로드."""
    url = f"{BASE_URL}/download/rulefiledown.json"
    resp = retry_request(session, "GET", url, params={"fileNo": file_no})
    if resp.status_code != 200:
        return None, resp.status_code
    if len(resp.content) <= 100:
        return None, resp.status_code
    return resp.content, resp.status_code


# ── 출력 / 체크포인트 ─────────────────────────────


class OutputWriter:
    """스레드 안전한 CSV·체크포인트 writer."""

    def __init__(self, out_dir: Path) -> None:
        self.out_dir = out_dir
        self.lock = Lock()
        self.csv_path = out_dir / "수집현황.csv"
        self.fail_path = out_dir / "failed_log.csv"
        self.ckpt_path = out_dir / "processed_inst.json"
        self._ensure_csv(self.csv_path, [
            "기관명", "기관유형", "주무부처", "분류", "규정명", "seq",
            "파일명", "확장자", "파일크기_bytes", "다운로드결과",
            "HTTP상태", "다운로드일시",
        ])
        self._ensure_csv(self.fail_path, [
            "기관명", "단계", "사유", "HTTP상태", "timestamp",
        ])
        self.processed: set[str] = self._load_ckpt()

    def _ensure_csv(self, path: Path, header: list[str]) -> None:
        if path.exists() and path.stat().st_size > 0:
            return
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerow(header)

    def _load_ckpt(self) -> set[str]:
        if not self.ckpt_path.exists():
            return set()
        try:
            return set(json.loads(self.ckpt_path.read_text(encoding="utf-8")).get(
                "completed", []
            ))
        except (json.JSONDecodeError, OSError):
            return set()

    def is_done(self, inst_name: str) -> bool:
        with self.lock:
            return inst_name in self.processed

    def mark_done(self, inst_name: str) -> None:
        with self.lock:
            self.processed.add(inst_name)
            self.ckpt_path.write_text(
                json.dumps(
                    {"completed": sorted(self.processed)},
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )

    def write_row(self, row: list) -> None:
        with self.lock, self.csv_path.open(
            "a", encoding="utf-8-sig", newline=""
        ) as f:
            csv.writer(f).writerow(row)

    def write_fail(self, row: list) -> None:
        with self.lock, self.fail_path.open(
            "a", encoding="utf-8-sig", newline=""
        ) as f:
            csv.writer(f).writerow(row)


# ── 핵심 처리 ───────────────────────────────────────


def process_institution(
    session: requests.Session,
    inst: dict,
    keyword: str,
    out_dir: Path,
    writer: OutputWriter,
    verbose: bool = True,
) -> dict:
    """한 기관에 대해 키워드 매칭 규정 모두 다운로드. 통계 dict 반환."""
    name = inst["apbaNa"]
    stats = {"matched": 0, "downloaded": 0, "failed": 0, "zero": 0}
    safe_inst = sanitize_filename(name, max_len=60)
    inst_dir = out_dir / safe_inst

    try:
        rules = fetch_rule_list(session, name)
    except requests.RequestException as e:
        writer.write_fail([
            name, "findRuleList", str(e)[:200], "", _now_str(),
        ])
        if verbose:
            print(f"  [실패] findRuleList: {e}")
        return stats

    if not rules:
        writer.write_row([
            name, inst.get("typeNa", ""), inst.get("jidtNa", ""),
            "", "", "", "", "", "", "0건", "", _now_str(),
        ])
        stats["zero"] = 1
        if verbose:
            print(f"  [0건] 내부규정 자체가 없거나 응답 없음")
        return stats

    matched = [r for r in rules if keyword in (r.get("title") or "")]
    stats["matched"] = len(matched)

    if not matched:
        writer.write_row([
            name, inst.get("typeNa", ""), inst.get("jidtNa", ""),
            "", "", "", "", "", "", f"키워드불일치(전체{len(rules)}건)",
            "", _now_str(),
        ])
        if verbose:
            print(f"  [매칭 0건] 전체 {len(rules)}건 중 '{keyword}' 포함 없음")
        return stats

    if verbose:
        print(f"  [매칭 {len(matched)}건] 전체 {len(rules)}건 중")

    inst_dir.mkdir(parents=True, exist_ok=True)
    used_names: set[str] = set()

    for rule in matched:
        seq = str(rule.get("seq") or "").strip()
        title = (rule.get("title") or "").strip()
        divis_raw = (rule.get("insdRuleDivis") or "").strip()
        divis = DIVIS_NAME_MAP.get(divis_raw, divis_raw or "기타")

        if not seq:
            continue

        try:
            files = fetch_rule_files(session, seq)
        except requests.RequestException as e:
            writer.write_fail([name, "findRuleDtl", str(e)[:200], "", _now_str()])
            stats["failed"] += 1
            continue

        # 최신(file_no 최대) 1건, .zip 제외
        candidates = [(no, n) for (no, n) in files if not n.lower().endswith(".zip")]
        if not candidates:
            writer.write_row([
                name, inst.get("typeNa", ""), inst.get("jidtNa", ""),
                divis, title, seq, "", "", "", "첨부없음", "", _now_str(),
            ])
            continue

        latest_no, latest_name = max(candidates, key=lambda x: x[0])

        try:
            content, http = download_rule_file(session, latest_no)
        except requests.RequestException as e:
            writer.write_fail([
                name, f"rulefiledown:{latest_no}", str(e)[:200], "", _now_str(),
            ])
            stats["failed"] += 1
            continue

        if content is None:
            writer.write_row([
                name, inst.get("typeNa", ""), inst.get("jidtNa", ""),
                divis, title, seq, latest_name,
                Path(latest_name).suffix.lower().lstrip("."),
                0, "실패(빈응답)", http, _now_str(),
            ])
            stats["failed"] += 1
            continue

        # 저장: 한 기관 폴더 안에 직접, 파일명 충돌 시 접미
        safe_file = sanitize_filename(latest_name, max_len=80)
        target = _resolve_unique(inst_dir / safe_file, used_names)
        used_names.add(target.name)
        target.write_bytes(content)
        stats["downloaded"] += 1
        writer.write_row([
            name, inst.get("typeNa", ""), inst.get("jidtNa", ""),
            divis, title, seq, target.name,
            target.suffix.lower().lstrip("."), len(content),
            "성공", http, _now_str(),
        ])
        if verbose:
            kb = len(content) // 1024
            print(f"    ✓ [{divis}] {title[:30]} → {target.name} ({kb}KB)")
        time.sleep(0.5)

    return stats


def _resolve_unique(path: Path, used: set[str]) -> Path:
    """동일 파일명 충돌 시 _2, _3 … 접미."""
    if not path.exists() and path.name not in used:
        return path
    base = path.stem
    ext = path.suffix
    i = 2
    while True:
        cand = path.with_name(f"{base}_{i}{ext}")
        if not cand.exists() and cand.name not in used:
            return cand
        i += 1


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ── 진입점 ──────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--mode", choices=["kicox", "all"], default="kicox",
        help="kicox: 한국산업단지공단 1개 / all: 전체 약 350개",
    )
    p.add_argument(
        "--out", type=Path, default=None,
        help="출력 폴더. 미지정 시 ~/alio_training_rules/공공기관_교육훈련규정_비교_YYYYMMDD/ (ALIO_RULES_OUT로 변경 가능)",
    )
    p.add_argument("--keyword", default="교육훈련", help="제목 매칭 키워드")
    p.add_argument(
        "--workers", type=int, default=5,
        help="기관 단위 병렬 스레드 수 (기본 5, all 모드에서만 적용)",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    # 출력 폴더 결정
    if args.out is not None:
        out_dir = args.out
    else:
        date_tag = datetime.now().strftime("%Y%m%d")
        out_dir = DEFAULT_OUT_BASE / f"공공기관_교육훈련규정_비교_{date_tag}"
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"[설정] 출력 폴더: {out_dir}")
    print(f"[설정] 키워드: '{args.keyword}'")
    print(f"[설정] 모드: {args.mode}")

    writer = OutputWriter(out_dir)
    if writer.processed:
        print(f"[체크포인트] 기존 처리 완료 {len(writer.processed)}개 기관 (재개)")

    session = create_session()

    # 기관 목록 결정
    if args.mode == "kicox":
        institutions = [{
            "apbaNa": "한국산업단지공단", "apbaId": "", "typeNa": "", "jidtNa": "",
        }]
    else:
        print("[수집] 알리오에서 전체 공공기관 목록 조회 중...")
        try:
            institutions = fetch_institutions(session)
        except requests.RequestException as e:
            print(f"[치명] 기관 목록 조회 실패: {e}")
            return 2
        print(f"[수집] 전체 {len(institutions)}개 기관 확보")

    pending = [i for i in institutions if not writer.is_done(i["apbaNa"])]
    if len(pending) < len(institutions):
        print(
            f"[건너뜀] 체크포인트 기준 {len(institutions) - len(pending)}개 이미 완료, "
            f"{len(pending)}개 처리 예정"
        )

    started = time.time()
    total_stats = {"matched": 0, "downloaded": 0, "failed": 0, "zero": 0}

    if args.mode == "kicox" or args.workers <= 1 or len(pending) <= 1:
        # 직렬 처리 (검증용·디버깅 친화)
        for idx, inst in enumerate(pending, 1):
            print(f"\n[{idx}/{len(pending)}] {inst['apbaNa']}")
            stats = process_institution(
                session, inst, args.keyword, out_dir, writer, verbose=True,
            )
            for k, v in stats.items():
                total_stats[k] += v
            writer.mark_done(inst["apbaNa"])
            time.sleep(0.3)
    else:
        # 병렬 (기관 단위)
        progress_lock = Lock()
        done_count = [0]

        def _worker(inst: dict) -> dict:
            # 워커마다 별도 세션 (스레드 안전성 확보)
            local_sess = create_session()
            stats = process_institution(
                local_sess, inst, args.keyword, out_dir, writer, verbose=False,
            )
            writer.mark_done(inst["apbaNa"])
            with progress_lock:
                done_count[0] += 1
                pct = done_count[0] / len(pending) * 100
                print(
                    f"[{done_count[0]}/{len(pending)} {pct:5.1f}%] "
                    f"{inst['apbaNa']} - "
                    f"매칭{stats['matched']}/다운{stats['downloaded']}/"
                    f"실패{stats['failed']}/0건{stats['zero']}"
                )
            return stats

        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            futures = [ex.submit(_worker, i) for i in pending]
            for fu in as_completed(futures):
                try:
                    s = fu.result()
                    for k, v in s.items():
                        total_stats[k] += v
                except Exception as e:
                    print(f"[워커 예외] {e}")

    elapsed = time.time() - started
    print("\n" + "=" * 60)
    print(f"[완료] {elapsed/60:.1f}분 소요")
    print(f"  처리 기관: {len(pending)}개")
    print(f"  매칭 규정: {total_stats['matched']}건")
    print(f"  다운로드 성공: {total_stats['downloaded']}건")
    print(f"  실패: {total_stats['failed']}건")
    print(f"  내부규정 0건 기관: {total_stats['zero']}개")
    print(f"  결과 폴더: {out_dir}")
    print(f"  수집현황: {writer.csv_path}")
    print(f"  실패 로그: {writer.fail_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
