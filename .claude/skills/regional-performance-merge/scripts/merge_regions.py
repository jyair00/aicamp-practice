"""형식이 다른 지역별 실적 파일을 표준 표 하나로 통합하고 검증 시트를 만든다.

사용법: python merge_regions.py <설정.json>
설정 항목은 references/config_guide.md 참고.
결과: 설정의 output 경로에 xlsx 저장, 표준 출력에 지역별 요약 JSON.
"""
import csv
import io
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from unicodedata import normalize

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FIELDS = ["id", "rev", "date", "campaign", "product", "amount", "lead", "appl", "cancel", "done", "status"]
DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"]


def nfc(s):
    return normalize("NFC", str(s)).strip() if s is not None else ""


def key(s):
    """머리글 비교용: NFC + 공백·줄바꿈 제거"""
    return re.sub(r"\s+", "", nfc(s))


def find_file(base: Path, name: str) -> Path:
    for p in base.iterdir():
        if nfc(p.name) == nfc(name):
            return p
    raise FileNotFoundError(f"파일을 찾을 수 없음: {name} (폴더: {base})")


def parse_date(v, fmt=None):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = nfc(v)
    for f in ([fmt] if fmt else []) + DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except (ValueError, TypeError):
            pass
    raise ValueError(f"날짜를 읽을 수 없음: {v!r}")


def parse_num(v):
    """'120,600', '1,200원', '₩300', '30건' → 숫자. 비어 있으면 None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = re.sub(r"[,\s원₩건]", "", nfc(v))
    if s in ("", "-"):
        return None
    return float(s) if "." in s else int(s)


def read_csv_rows(p: Path):
    raw = p.read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    return list(csv.reader(io.StringIO(text)))


def load_source(base: Path, src: dict, cfg: dict):
    """한 원본 파일을 읽어 표준 레코드 목록을 돌려준다."""
    p = find_file(base, src["file"])
    cols = src["columns"]
    hr = int(src.get("header_row", 1))
    tables = []  # (sheet_name, header, [(row_no, values)], sheet_note)
    if p.suffix.lower() == ".csv":
        rows = read_csv_rows(p)
        tables.append(("csv", rows[hr - 1], [(i + 1, r) for i, r in enumerate(rows) if i >= hr], ""))
        skipped = []
    else:
        if not p.read_bytes()[:2] == b"PK":
            raise ValueError(f"{p.name}: 일반 xlsx가 아님 (DRM 암호화 가능성) — 보안 해제본이 필요합니다")
        wb = openpyxl.load_workbook(p, data_only=True)
        want = src.get("sheets")
        skip = {nfc(s) for s in src.get("skip_sheets", [])}
        skipped = [ws.title for ws in wb.worksheets if nfc(ws.title) in skip]
        for ws in wb.worksheets:
            if nfc(ws.title) in skip or (want and nfc(ws.title) not in {nfc(s) for s in want}):
                continue
            rows = list(ws.iter_rows(values_only=True))
            note = " ".join(nfc(c) for r in rows[: hr - 1] for c in r if c is not None)
            tables.append((ws.title, rows[hr - 1], [(i + 1, r) for i, r in enumerate(rows) if i >= hr], note))

    factor = float(src.get("amount_multiply", 1))
    divisor = float(src.get("amount_divide", 1))
    recs, warnings = [], []
    for sheet, header, body, note in tables:
        idx = {key(h): i for i, h in enumerate(header) if h is not None}
        pos = {}
        for f in FIELDS:
            name = cols.get(f)
            if name is None:
                continue
            if key(name) not in idx:
                raise KeyError(f"{p.name}/{sheet}: 머리글에 '{name}' 열이 없음. 실제 머리글: {list(header)}")
            pos[f] = idx[key(name)]
        sheet_camp = None
        if "campaign" not in pos:
            rule = src.get("campaign_from")
            if not rule:
                raise KeyError(f"{p.name}/{sheet}: 캠페인 열도 campaign_from 설정도 없음")
            text = {"sheet_note": note, "sheet_name": sheet}[rule.get("source", "sheet_note")]
            m = re.search(rule["pattern"], text)
            if not m:
                raise ValueError(f"{p.name}/{sheet}: '{text}'에서 캠페인 코드를 찾지 못함")
            sheet_camp = m.group(1)
        for row_no, r in body:
            if r is None or all(v is None or nfc(v) == "" for v in r):
                continue
            g = lambda f: r[pos[f]] if f in pos and pos[f] < len(r) else None
            raw_amt = g("amount")
            num = parse_num(raw_amt)
            memo = []
            if num is None:
                amt = None
                memo.append("금액 비어 있음 — 확인 필요")
            else:
                won = num * factor / divisor
                amt = int(round(won))
                if abs(won - amt) > 1e-6:
                    warnings.append(f"{src['region']} {nfc(g('id'))}: {raw_amt} → {won:.4f}원, 반올림 {amt}원")
                    memo.append(f"환산값 {won:.2f}원 반올림")
            camp = nfc(g("campaign")) if sheet_camp is None else sheet_camp
            prod_raw, st_raw = nfc(g("product")), nfc(g("status"))
            for label, val, mp in (("상품", prod_raw, cfg["product_map"]), ("상태", st_raw, cfg["status_map"]),
                                   ("캠페인", camp, cfg["campaign_map"])):
                if val not in mp:
                    raise ValueError(f"{p.name}/{sheet} {row_no}행: 매핑에 없는 {label} 값 '{val}' — 설정에 추가하세요")
            recs.append(dict(
                region=src["region"], file=p.name, sheet=sheet, row=row_no,
                id=nfc(g("id")), rev=int(parse_num(g("rev")) or 1),
                date=parse_date(g("date"), src.get("date_format")),
                camp=camp, act=cfg["campaign_map"][camp], prod=cfg["product_map"][prod_raw],
                amt_raw=raw_amt, amt=amt,
                lead=parse_num(g("lead")) or 0, appl=parse_num(g("appl")) or 0,
                canc=parse_num(g("cancel")) or 0, done=parse_num(g("done")) or 0,
                st_raw=st_raw, st=cfg["status_map"][st_raw], memo=memo,
            ))
    return recs, skipped, warnings


def classify(recs, period):
    y, m = map(int, period.split("-"))
    in_period = lambda x: x["date"].year == y and x["date"].month == m
    max_rev = defaultdict(int)
    for x in recs:
        if in_period(x):
            max_rev[(x["region"], x["id"])] = max(max_rev[(x["region"], x["id"])], x["rev"])
    seen = set()
    for x in recs:
        sig = (x["region"], x["id"], x["rev"], x["date"], x["camp"], x["prod"], x["amt"],
               x["lead"], x["appl"], x["canc"], x["done"], x["st_raw"])
        if not in_period(x):
            x["excl"] = f"보고기간({period}) 밖"
        elif sig in seen:
            x["excl"] = "완전히 같은 중복 행"
        elif x["rev"] < max_rev[(x["region"], x["id"])]:
            x["excl"] = f"이전 차수 (최신 {max_rev[(x['region'], x['id'])]}차 반영)"
        else:
            x["excl"] = ""
        seen.add(sig)
    kept = [x for x in recs if not x["excl"]]
    dup = [k for k, v in Counter((x["region"], x["id"]) for x in kept).items() if v > 1]
    if dup:
        raise ValueError(f"통합 후에도 같은 기록ID가 남음 (내용이 다른 같은 차수): {dup}")
    for x in kept:
        if x["st"] == "미마감":
            x["memo"].insert(0, f"원본 상태 '{x['st_raw']}' → 확정 합계 제외")
        if x["rev"] > 1:
            x["memo"].append(f"{x['rev']}차 수정값 반영")
    return kept, [x for x in recs if x["excl"]]


# ---------------- 엑셀 작성 ----------------
RED, TXT, TXT2, BG, BORDER = "EA002C", "202124", "5F6368", "F7F8FA", "E5E7EB"
FONT = "Malgun Gothic"
F_TITLE = Font(name=FONT, size=14, bold=True, color=TXT)
F_NOTE = Font(name=FONT, size=9, color=TXT2)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=FONT, size=10, color=TXT)
F_BOLD = Font(name=FONT, size=10, bold=True, color=TXT)
F_WARN = Font(name=FONT, size=10, bold=True, color="925000")
F_GOOD = Font(name=FONT, size=10, bold=True, color="287D4E")
F_BAD = Font(name=FONT, size=10, bold=True, color="B42318")
FILL_HEAD = PatternFill("solid", fgColor=RED)
FILL_SUB = PatternFill("solid", fgColor=BG)
FILL_WARN = PatternFill("solid", fgColor="FDF1E8")
BOX = Border(*(Side(style="thin", color=BORDER),) * 4)
NUM, DT = "#,##0", "yyyy-mm-dd"


def header(ws, title, note, heads, widths, hr=4):
    ws["A1"], ws["A2"] = title, note
    ws["A1"].font, ws["A2"].font = F_TITLE, F_NOTE
    for c, (h, w) in enumerate(zip(heads, widths), 1):
        cell = ws.cell(hr, c, h)
        cell.font, cell.fill, cell.border = F_HEAD, FILL_HEAD, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[hr].height = 30
    ws.sheet_view.showGridLines = False


def put(ws, r, c, v, fmt=None, font=None):
    cell = ws.cell(r, c, v)
    cell.font, cell.border = font or F_BODY, BOX
    if fmt:
        cell.number_format = fmt
    return cell


def write_book(cfg, recs, kept, excl, regions, skipped, warnings, out_path):
    wb = openpyxl.Workbook()
    note = cfg.get("note", "금액: 원, 부가세 제외")

    ws = wb.active
    ws.title = "통합실적"
    H = ["지역", "집행일", "활동구분", "캠페인ID", "기록ID", "수정차수", "대상상품", "사용금액(원/VAT별도)",
         "상담 건수", "접수건", "취소", "처리완료", "마감여부", "메모", "원본파일·시트"]
    header(ws, cfg.get("title", "실적 통합"), note, H, [10, 12, 13, 10, 17, 9, 13, 16, 10, 9, 8, 10, 10, 34, 36])
    kept.sort(key=lambda x: (regions.index(x["region"]), x["date"], x["camp"], x["id"]))
    for n, x in enumerate(kept):
        r = 5 + n
        vals = [x["region"], x["date"], x["act"], x["camp"], x["id"], x["rev"], x["prod"], x["amt"], x["lead"],
                x["appl"], x["canc"], x["done"], x["st"], " / ".join(x["memo"]), f"{x['file']} / {x['sheet']} {x['row']}행"]
        for c, v in enumerate(vals, 1):
            put(ws, r, c, v, DT if c == 2 else (NUM if c in (6, 8, 9, 10, 11, 12) else None))
        if x["st"] == "미마감":
            for c in range(1, len(H) + 1):
                ws.cell(r, c).fill = FILL_WARN
            ws.cell(r, 13).font = F_WARN
    last = 4 + max(len(kept), 1)
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(H))}{last}"
    rng = lambda col: f"통합실적!${col}$5:${col}${last}"

    ws2 = wb.create_sheet("지역별요약")
    H2 = ["지역", "원본 행", "제외 행", "통합 행", "미마감 행", "확정 행", "확정 금액(원)", "미마감 금액(원)",
          "확정 상담", "확정 접수", "확정 취소", "확정 완료"]
    header(ws2, '지역별 요약 (확정 = 마감여부 "마감")', "미마감·보고기간 밖 자료는 확정 실적에서 제외 · 수식으로 계산", H2,
           [10, 9, 9, 9, 10, 9, 16, 16, 11, 11, 10, 11])
    raw_cnt, exc_cnt = Counter(x["region"] for x in recs), Counter(x["region"] for x in excl)
    for n, reg in enumerate(regions):
        r = 5 + n
        A = rng("A")
        vals = [reg, raw_cnt[reg], exc_cnt[reg], f"=COUNTIFS({A},$A{r})",
                f'=COUNTIFS({A},$A{r},{rng("M")},"미마감")', f'=COUNTIFS({A},$A{r},{rng("M")},"마감")',
                f'=SUMIFS({rng("H")},{A},$A{r},{rng("M")},"마감")', f'=SUMIFS({rng("H")},{A},$A{r},{rng("M")},"미마감")']
        vals += [f'=SUMIFS({rng(col)},{A},$A{r},{rng("M")},"마감")' for col in "IJKL"]
        for c, v in enumerate(vals, 1):
            put(ws2, r, c, v, NUM if c > 1 else None)
    rt = 5 + len(regions)
    put(ws2, rt, 1, "합계", font=F_BOLD).fill = FILL_SUB
    for c in range(2, 13):
        L = get_column_letter(c)
        put(ws2, rt, c, f"=SUM({L}5:{L}{rt - 1})", NUM, F_BOLD).fill = FILL_SUB

    ws3 = wb.create_sheet("제외내역")
    H3 = ["지역", "제외 사유", "기록ID", "수정차수", "날짜", "캠페인ID", "원본 금액 값", "환산 금액(원)", "원본 상태", "원본파일·시트"]
    header(ws3, "통합에서 뺀 행", "보고기간 밖 → 완전 중복 → 이전 차수 순서로 판정", H3, [10, 28, 17, 9, 12, 10, 14, 15, 10, 36])
    for n, x in enumerate(sorted(excl, key=lambda x: (regions.index(x["region"]), x["excl"], x["id"]))):
        vals = [x["region"], x["excl"], x["id"], x["rev"], x["date"], x["camp"], x["amt_raw"], x["amt"], x["st_raw"],
                f"{x['file']} / {x['sheet']} {x['row']}행"]
        for c, v in enumerate(vals, 1):
            put(ws3, 5 + n, c, v, DT if c == 5 else (NUM if c == 8 else None))

    ws4 = wb.create_sheet("검증")
    H4 = ["지역", "항목", "원본 전체", "통합", "제외", "통합+제외", "차이", "결과"]
    header(ws4, "원본 대조 검증", "원본 전체(원 단위 환산) = 통합 + 제외 이어야 함", H4, [10, 16, 16, 16, 16, 16, 10, 10])
    r = 5
    checks = []
    for reg in regions:
        R = [x for x in recs if x["region"] == reg]
        K = [x for x in kept if x["region"] == reg]
        E = [x for x in excl if x["region"] == reg]
        for label, fn in (("행 수", len), ("금액(원)", lambda L: sum(x["amt"] or 0 for x in L)),
                          ("상담 건수", lambda L: sum(x["lead"] for x in L)), ("처리완료", lambda L: sum(x["done"] for x in L))):
            a, k, e = fn(R), fn(K), fn(E)
            ok = a == k + e
            checks.append(ok)
            for c, v in enumerate([reg, label, a, k, e, k + e, a - (k + e), "일치" if ok else "불일치"], 1):
                put(ws4, r, c, v, NUM if 3 <= c <= 7 else None, (F_GOOD if ok else F_BAD) if c == 8 else None)
            r += 1
    r += 1
    for msg in ([f"제외한 시트: {', '.join(skipped)}"] if skipped else []) + [f"반올림: {w}" for w in warnings]:
        ws4.cell(r, 1, msg).font = F_NOTE
        r += 1

    ws5 = wb.create_sheet("변환기준")
    header(ws5, "변환 기준", "설정 파일 기준", ["구분", "원본 값/지역", "변환"], [12, 30, 40])
    rows = [("금액 단위", s["region"], s.get("unit_note", f"×{s.get('amount_multiply', 1)} ÷{s.get('amount_divide', 1)}"))
            for s in cfg["sources"]]
    rows += [("상품", k, v) for k, v in cfg["product_map"].items()]
    rows += [("마감 상태", k, v) for k, v in cfg["status_map"].items()]
    rows += [("캠페인", k, v) for k, v in cfg["campaign_map"].items()]
    for n, row in enumerate(rows):
        for c, v in enumerate(row, 1):
            put(ws5, 5 + n, c, v)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return all(checks)


def main():
    cfg_path = Path(sys.argv[1]).resolve()
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    base = (cfg_path.parent / cfg.get("input_dir", ".")).resolve()
    out_path = (cfg_path.parent / cfg["output"]).resolve()
    recs, skipped, warnings = [], [], []
    for src in cfg["sources"]:
        r, s, w = load_source(base, src, cfg)
        recs += r
        skipped += [f"{src['region']}/{x}" for x in s]
        warnings += w
    regions = [s["region"] for s in cfg["sources"]]
    regions = list(dict.fromkeys(regions))
    kept, excl = classify(recs, cfg["period"])
    ok = write_book(cfg, recs, kept, excl, regions, skipped, warnings, out_path)

    summary = {"output": str(out_path), "verified": ok, "regions": {}}
    for reg in regions:
        K = [x for x in kept if x["region"] == reg]
        conf = [x for x in K if x["st"] == "마감"]
        summary["regions"][reg] = {
            "원본행": sum(1 for x in recs if x["region"] == reg),
            "제외행": sum(1 for x in excl if x["region"] == reg),
            "통합행": len(K), "확정행": len(conf),
            "확정금액": sum(x["amt"] or 0 for x in conf),
            "미마감금액": sum(x["amt"] or 0 for x in K if x["st"] == "미마감"),
            "확정상담": sum(x["lead"] for x in conf), "확정접수": sum(x["appl"] for x in conf),
            "확정취소": sum(x["canc"] for x in conf), "확정완료": sum(x["done"] for x in conf),
            "빈금액행": sum(1 for x in K if x["amt"] is None),
        }
    tot = defaultdict(int)
    for v in summary["regions"].values():
        for k, n in v.items():
            tot[k] += n
    summary["합계"] = dict(tot)
    summary["제외사유"] = dict(Counter(re.sub(r" \(.*\)", "", x["excl"]) for x in excl))
    summary["제외시트"] = skipped
    summary["반올림경고"] = warnings
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    if not ok:
        sys.exit("검증 불일치가 있습니다. 검증 시트를 확인하세요.")


if __name__ == "__main__":
    main()
