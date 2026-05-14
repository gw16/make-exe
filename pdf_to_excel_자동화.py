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
SHORT_LEN = 35  # 이 글자 수 미만이면 제목/항목 후보로 보고 다음 줄과 합치지 않음

# 한국어 종결어미: 종성 없는 "다/요/까/네/소/오/죠/라/용/니" + ㅁ 받침을 가진 모든 한글(임/함/됨/음/짐/님/감 등)
_TERMINAL_NON_M = set("다요까네소오죠라용니")
_CLOSE_BRACKETS = set(")]”’\"'")
_END_PUNCT = set(".!?")

# 불릿/정의 박스 시작 패턴 → 항상 별도 행으로 처리
_BULLET_RE = re.compile(r"^\s*([·•▪▫◦■□▶▷◆◇★☆※]|\[)")
# 닫는 기호만 단독으로 있는 줄 → 이전 줄에 공백 없이 합치기 (예: PDF가 ">"를 다음 줄로 떼어놓은 경우)
_CLOSE_ONLY_RE = re.compile(r"^\s*[>)\]\"”’]+\s*$")


def _has_m_jongseong(ch: str) -> bool:
    if not ch or not ("가" <= ch <= "힣"):
        return False
    return (ord(ch) - 0xAC00) % 28 == 16


def _is_korean_terminal(ch: str) -> bool:
    return ch in _TERMINAL_NON_M or _has_m_jongseong(ch)


def _is_sentence_end(text: str) -> bool:
    """줄 끝이 한국어 종결로 끝나는지 (닫는 괄호/따옴표는 건너뛴 뒤 직전 글자 검사)."""
    text = text.rstrip()
    if len(text) < 2 or text[-1] not in _END_PUNCT:
        return False
    idx = len(text) - 2
    while idx >= 0 and text[idx] in _CLOSE_BRACKETS:
        idx -= 1
    if idx < 0:
        return False
    return _is_korean_terminal(text[idx])


def _is_open_bracket(text: str) -> bool:
    """'['로 시작했지만 같은 줄에 ']'가 없는 줄 (멀티라인 정의 박스의 시작)."""
    s = text.strip()
    return s.startswith("[") and "]" not in s


def _looks_word_break(next_line: str) -> bool:
    """다음 줄 첫 어절이 1~3자 한글이고 종결 부호로 끝나면 단어 잘림 케이스
    (예: "되었\n다." → "되었다.")."""
    stripped = next_line.lstrip()
    if not stripped:
        return False
    first_token = stripped.split()[0]
    core = re.sub(r"[\.\!\?\"'”’]+$", "", first_token)
    if not (1 <= len(core) <= 3):
        return False
    # 숫자만이면 번호 매김 → 단어 잘림 아님
    if core.isdigit():
        return False
    return bool(re.search(r"[\.\!\?]", first_token))


def smart_split(full_text: str) -> list:
    raw_lines = [l for l in full_text.split("\n") if l.strip()]
    if not raw_lines:
        return []

    merged = []
    buffer = ""
    for line in raw_lines:
        line = line.strip()
        if not buffer:
            buffer = line
            continue

        # 닫는 기호 단독 줄 (예: ">", ")") → 이전 buffer에 공백 없이 합치기
        if _CLOSE_ONLY_RE.fullmatch(line):
            buffer += line
            continue

        # buffer가 '[' 시작이지만 ']'로 닫히지 않았으면 닫힐 때까지 계속 합치기
        buffer_unclosed = _is_open_bracket(buffer)

        # 다음 줄이 불릿/정의 박스로 시작 → 새 항목, 합치지 않고 끊음
        # 단, buffer가 미완성 '['이면 새 항목 트리거 무시
        if _BULLET_RE.match(line) and not buffer_unclosed:
            merged.append(buffer)
            buffer = line
            continue
        # 현재 buffer가 불릿/정의 박스로 시작 → 그 자체로 마무리 (단, 미완성 '['은 예외)
        if _BULLET_RE.match(buffer) and not buffer_unclosed:
            merged.append(buffer)
            buffer = line
            continue

        prev_ends = _is_sentence_end(buffer)
        prev_short = len(buffer) < SHORT_LEN

        if (prev_ends or prev_short) and not buffer_unclosed:
            merged.append(buffer)
            buffer = line
        else:
            if _looks_word_break(line):
                buffer += line  # 단어 잘림 복원
            else:
                buffer += " " + line
    if buffer:
        merged.append(buffer)

    # 합쳐진 행 내부에서 한국어 종결 마침표 위치마다 문장 분할
    out = []
    end_punct_re = re.compile(r"[\.!?][\"'”’]?")
    closing_chars = ">)]”’\"'"
    for line in merged:
        last = 0
        for m in end_punct_re.finditer(line):
            end = m.end()
            # 종결 부호 뒤가 공백/줄 끝/닫는 기호 이외이면 skip (닫는 기호는 아래에서 흡수)
            if end < len(line) and not line[end].isspace() and line[end] not in closing_chars:
                continue
            # 종결 부호 다음의 "공백 + 닫는 기호" 묶음을 모두 분할 위치 뒤로 흡수
            # (예: "< 예시. > 그 다음" → 첫 묶음은 "< 예시. >" 가 되고 분할은 그 뒤에서 일어남)
            extended = end
            while extended < len(line):
                j = extended
                while j < len(line) and line[j] == " ":
                    j += 1
                if j < len(line) and line[j] in closing_chars:
                    k = j
                    while k < len(line) and line[k] in closing_chars:
                        k += 1
                    extended = k
                else:
                    break
            # 흡수 후 위치가 줄 끝이거나 공백이어야 분할 가능
            if extended < len(line) and not line[extended].isspace():
                continue
            if not _is_sentence_end(line[:end]):
                continue
            chunk = line[last:extended].strip()
            if chunk:
                out.append(chunk)
            last = extended
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
    image_only_files = []  # 텍스트 추출이 안 된 PDF (이미지 기반)
    for file in os.listdir(folder_path):
        if not file.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, file)
        with pdfplumber.open(path) as pdf:
            page_texts = [p.extract_text() or "" for p in pdf.pages]

        if not any(t.strip() for t in page_texts):
            image_only_files.append(file)
            continue

        patterns = detect_repeated_edges(page_texts)
        cleaned_text = strip_headers_footers(page_texts, patterns)

        sentences = smart_split(cleaned_text)
        sentences = split_inline_numbering(sentences)

        for s in sentences:
            rows.append([file, s])

    if not rows:
        msg = "텍스트를 추출하지 못했습니다."
        if image_only_files:
            msg += "\n\n이미지 기반 PDF로 보이는 파일:\n" + "\n".join(
                f" - {f}" for f in image_only_files
            )
            msg += "\n\nOCR 기능이 없는 현재 버전으로는 처리할 수 없습니다."
        show_warning(msg)
        return 0

    df = pd.DataFrame(rows, columns=["파일", "텍스트"])
    output_path = os.path.join(folder_path, "통합_결과.xlsx")
    df.to_excel(output_path, index=False, engine="openpyxl")

    completion = f"완료!\n\n결과 파일:\n{output_path}"
    if image_only_files:
        completion += "\n\n다음 파일은 이미지 기반으로 보여 건너뛰었습니다:\n" + "\n".join(
            f" - {f}" for f in image_only_files
        )
    show_info(completion)
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
