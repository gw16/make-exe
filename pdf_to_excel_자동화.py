import os
import re
import sys
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory

APP_TITLE = "PDF → Excel 변환기"


def show_error(msg: str) -> None:
    messagebox.showerror(APP_TITLE, msg)


def show_info(msg: str) -> None:
    messagebox.showinfo(APP_TITLE, msg)


def show_warning(msg: str) -> None:
    messagebox.showwarning(APP_TITLE, msg)


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
            full_text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    full_text += t + "\n"

        lines = [l.strip() for l in full_text.split("\n") if l.strip()]
        for line in lines:
            sentences = re.split(r"(?<=[\.\?\!])\s+", line)
            for s in sentences:
                if s.strip():
                    rows.append([file, s.strip()])

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
