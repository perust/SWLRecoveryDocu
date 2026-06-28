# SWLRecoveryDocu

데이터 복구 작업 후 고객에게 전달할 작업내역서 문서를 빠르게 생성하기 위한 데스크톱 문서 생성 도구입니다. PyQt6 기반 GUI에서 입고 정보, 매체 정보, 작업 내용, 최종 결론을 입력하고 Word/PDF 문서를 생성하는 것을 목표로 합니다.

## 주요 기능

- 작업내역서 입력용 PyQt6 데스크톱 UI
- 관리번호 기반 문서 번호 생성
- 엔지니어, 매체 종류, 파일 시스템, 모델명, S/N 등 입고 정보 입력
- 증상, 복구/분석 요청 사항, 작업 로그, 최종 결론 입력
- `docxtpl` 기반 Word 템플릿 치환
- 선택적으로 PDF 변환 및 암호화 ZIP 압축 지원
- 환경설정(`settings.json`)을 통한 프리셋 관리

## 기술 스택

- Python
- PyQt6
- docxtpl
- docx2pdf
- pyzipper
- PyInstaller

## 프로젝트 구조

```text
.
├── Data_Recovery_Report_Generator_PRD.md
├── template_type1.docx
├── AS2601010000111 업체명.docx
└── SWLRecoveryDocu/
    ├── main.py             # PyQt6 GUI 진입점
    ├── doc_generator.py    # 문서 번호/Word/PDF/ZIP 생성 로직
    └── template_type1.docx # 문서 템플릿
```

## 설치 및 실행

```bash
git clone https://github.com/perust/SWLRecoveryDocu.git
cd SWLRecoveryDocu/SWLRecoveryDocu

python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install PyQt6 docxtpl docx2pdf pyzipper

python main.py
```

## 사용 흐름

1. 담당 엔지니어와 고객/매체 정보를 입력합니다.
2. 증상, 요청 사항, 작업 로그, 최종 결론을 선택하거나 직접 작성합니다.
3. 양식과 PDF/압축 옵션을 선택합니다.
4. **작업내역서 생성** 버튼으로 문서를 내보냅니다.

## 참고

- PDF 변환에는 일반적으로 Microsoft Word가 설치된 Windows 환경이 필요합니다.
- 자세한 요구사항은 `Data_Recovery_Report_Generator_PRD.md`를 참고하세요.
