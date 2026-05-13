import os
import re
import sys
from collections import Counter
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory

APP_TITLE = "PDF → Excel 변환기"


def show_error(msg: str) -> None:
    messagebox.showerror(APP_TITLE, msg)


def show_info(msg: str) -> None:
    messagebox.showinfo(APP_TITLE, msg)


def show_warning(msg: str) -> None:
    messagebox.showwarning(APP_TITLE, msg)


# ─── 페이지 헤더/푸터 자동 감지·제거 ──────────────────────
def _norm_for_compare(s: str) -> str:
    # 숫자/공백 차이를 무시한 비교용 정규화
    return re.sub(r"\s+", " ", re.sub(r"\d+", "", s)).strip()


def detect_repeated_edges(page_texts: list) -> set:
    if len(page_texts) < 2:
        return set()
    edge_norms = []
    for txt in page_texts:
        lines = [l for l in txt.split("\n") if l.strip()]
        if not lines:
            continue
        for l in lines[:2] + lines[-2:]:
            n = _norm_for_compare(l)
            if n and len(n) >= 4:
                edge_norms.append(n)
    counter = Counter(edge_norms)
    threshold = max(2, len(page_texts))
    return {n for n, c in counter.items() if c >= threshold}


def strip_headers_footers(page_texts: list, patterns: set) -> str:
    cleaned = []
    for txt in page_texts:
        for line in txt.split("\n"):
            if _norm_for_compare(line) in patterns:
                continue
            cleaned.append(line)
    return "\n".join(cleaned)


# ─── 줄 합치기 + 문장 분할 ────────────────────────────────
_SENTENCE_END_RE = re.compile(
    r"(?:[다요까네소오죠임함됨음라용니]|[a-zA-Z\)\]”’\"])[\.!?][\"'”’]?$"
)
_SENT_PAT = re.compile(
    r"(?<=[다요까네소오죠임함됨음라용니])[\.!?][\"'”’]?(?=\s+|$)"
    r"|(?<=[a-zA-Z\)\]”’\"])[\.!?][\"'”’]?(?=\s+|$)"
)


def smart_split(full_text: str) -> list:
    raw_lines = [l for l in full_text.split("\n") if l.strip()]
    if not raw_lines:
        return []

    max_len = max(len(l) for l in raw_lines)
    # 본문 너비를 거의 채운 줄은 다음 줄로 이어진다고 본다
    threshold = max(int(max_len * 0.85), 30)

    merged = []
    buffer = ""
    for line in raw_lines:
        line = line.strip()
        if not buffer:
            buffer = line
            continue

        prev_long = len(buffer) >= threshold
        prev_ends = bool(_SENTENCE_END_RE.search(buffer.rstrip()))

        if prev_ends or not prev_long:
            merged.append(buffer)
            buffer = line
        else:
            last_char = buffer[-1]
            first_char = line[0]
            # 한글+(한글|여는괄호) → 공백 없이 (PDF 단어 잘림 복원)
            if re.match(r"[가-힣]", last_char) and re.match(r"[가-힣\(\[]", first_char):
                buffer += line
            else:
                buffer += " " + line
    if buffer:
        merged.append(buffer)

    out = []
    for line in merged:
        last = 0
        for m in _SENT_PAT.finditer(line):
            chunk = line[last:m.end()].strip()
            if chunk:
                out.append(chunk)
            last = m.end()
        tail = line[last:].strip()
        if tail:
            out.append(tail)
    return out


# ─── 인라인 번호 매김 분할 ("... 1. ... 2. ...") ───────────
_INLINE_NUM_RE = re.compile(r"(?<=\S)\s+(?=\d{1,2}\.\s+\S)")


def split_inline_numbering(items: list) -> list:
    out = []
    for s in items:
        for p in _INLINE_NUM_RE.split(s):
            p = p.strip()
            if p:
                out.append(p)
    return out


def main() -> int:
    Tk().withdraw()

    try:
        import pdfplumber
        import pandas as pd
    except ImportError as e:
        show_error(
            f"필수 라이브러리를 불러오지 못했습니다.\n\n{e}\n\n"
            "EXE 재빌드가 필요할 수 있습니다."
        )
        return 1

    folder_path = askdirectory(title="PDF 폴더를 선택하세요")
    if not folder_path:
        show_info("폴더 선택이 취소되었습니다.")
        return 0

    rows = []
    for file in os.listdir(folder_path):
        if not file.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, file)
        with pdfplumber.open(path) as pdf:
            page_texts = [p.extract_text() or "" for p in pdf.pages]

        patterns = detect_repeated_edges(page_texts)
        cleaned_text = strip_headers_footers(page_texts, patterns)

        sentences = smart_split(cleaned_text)
        sentences = split_inline_numbering(sentences)

        for s in sentences:
            rows.append([file, s])

    if not rows:
        show_warning(
            "텍스트를 추출하지 못했습니다.\n"
            "PDF가 이미지 기반일 수 있습니다 (OCR 필요)."
        )
        return 0

    df = pd.DataFrame(rows, columns=["파일", "텍스트"])
    output_path = os.path.join(folder_path, "통합_결과.xlsx")
    df.to_excel(output_path, index=False, engine="openpyxl")

    show_info(f"완료!\n\n결과 파일:\n{output_path}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        # --noconsole 빌드에서도 예기치 못한 예외를 사용자에게 보여주기 위함
        try:
            Tk().withdraw()
            show_error(f"예상치 못한 오류가 발생했습니다.\n\n{type(e).__name__}: {e}")
        finally:
            sys.exit(1)
