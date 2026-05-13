@echo off
REM 로컬 Windows 빌드용. 회사 PC에서 Python 설치 불가 시 GitHub Actions 사용.

REM 이전 빌드 산출물 정리
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist pdf_to_excel.spec del /q pdf_to_excel.spec

REM 의존성 설치
pip install -r requirements.txt pyinstaller

REM EXE 빌드 (출력명은 영문으로 고정해 인코딩 이슈 회피)
pyinstaller --onefile --noconsole --name pdf_to_excel pdf_to_excel_자동화.py

echo ================
echo EXE 생성 완료!
echo dist\pdf_to_excel.exe 를 확인하세요
echo ================
pause
