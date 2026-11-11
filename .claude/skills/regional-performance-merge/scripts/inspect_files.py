"""폴더 안의 엑셀(.xlsx)·CSV 파일 구조를 UTF-8 텍스트로 정리한다.

사용법: python inspect_files.py <입력폴더> <결과.txt> [미리볼 행 수=8]
"""
import csv
import io
import sys
from pathlib import Path

import openpyxl


def decode(raw: bytes):
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8(깨진 글자 대체)"


def short(row, width=400):
    s = repr(tuple(row))
    return s if len(s) <= width else s[:width] + " …"


def inspect_xlsx(p: Path, n: int, out: list):
    head = p.read_bytes()[:8]
    if not head.startswith(b"PK"):
        out.append(f"  ⚠ 일반 xlsx가 아님 (파일 시작: {head!r}). 사내 문서보안(DRM) 암호화 파일일 수 있음 → 보안 해제본 필요")
        return
    wb = openpyxl.load_workbook(p, data_only=True)
    for ws in wb.worksheets:
        merged = [str(m) for m in ws.merged_cells.ranges][:8]
        out.append(f"  --- 시트 '{ws.title}' 범위={ws.dimensions} 최대행={ws.max_row} 병합={merged}")
        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
        for i, r in enumerate(rows[:n], 1):
            out.append(f"  {i:>3}: {short(r)}")
        if len(rows) > n:
            out.append("  ...")
            for i in range(max(n, len(rows) - 3), len(rows)):
                out.append(f"  {i + 1:>3}: {short(rows[i])}")


def inspect_csv(p: Path, n: int, out: list):
    text, enc = decode(p.read_bytes())
    rows = list(csv.reader(io.StringIO(text)))
    out.append(f"  인코딩={enc} 행 수(머리글 포함)={len(rows)}")
    for i, r in enumerate(rows[:n], 1):
        out.append(f"  {i:>3}: {short(r)}")
    if len(rows) > n:
        out.append("  ...")
        for i in range(max(n, len(rows) - 3), len(rows)):
            out.append(f"  {i + 1:>3}: {short(rows[i])}")


def main():
    folder, dest = Path(sys.argv[1]), Path(sys.argv[2])
    n = int(sys.argv[3]) if len(sys.argv) > 3 else 8
    out = [f"폴더: {folder}"]
    files = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".csv"))
    for p in files:
        out.append(f"\n##### {p.name} ({p.stat().st_size:,} bytes)")
        try:
            (inspect_csv if p.suffix.lower() == ".csv" else inspect_xlsx)(p, n, out)
        except Exception as e:  # 한 파일 오류로 전체가 멈추지 않게
            out.append(f"  ⚠ 읽기 실패: {type(e).__name__}: {e}")
    dest.write_text("\n".join(out), encoding="utf-8")
    print(f"{len(files)}개 파일 확인 → {dest}")


if __name__ == "__main__":
    main()
