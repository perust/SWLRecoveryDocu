# 제품 요구사항 정의서 (PRD)

**프로젝트명**: Data Recovery Report Generator (DRRG)
**버전**: 1.6 (Stable / UI Fixed)
**작성일**: 2026년 1월 27일
**작성자**: Gemini (for User)

## 1. 프로젝트 개요 (Overview)
**목적**: 데이터 복구 엔지니어가 작업 완료 후 고객에게 전달할 ‘작업내역서’(Word/PDF)를 빠르고 표준화된 양식으로 자동 생성하기 위함.

**핵심 가치**: 반복적인 문서 작성 시간 단축, 휴먼 에러(오타, 번호 실수) 방지, 보안 전송(암호화 압축) 지원.

**타겟 사용자**: 데이터 복구 전문 엔지니어.

## 2. 기술 스택 (Tech Stack)
*   **Language**: Python 3.14+
*   **GUI Framework**: PyQt6 (Windows Desktop Application)
*   **Document Engine**:
    *   `docxtpl`: MS Word (.docx) 템플릿 치환 및 생성.
    *   `docx2pdf`: Word를 PDF로 변환 (MS Word 설치 필수).
    *   `pyzipper`: AES 암호화 방식의 ZIP 압축 지원.
*   **Build Tool**: PyInstaller (Single File EXE).

## 3. 주요 기능 및 UI 구성 (Functional Requirements)
### 3.1. 메인 화면 레이아웃 (Split Layout)
화면은 좌우 크기 조절이 가능한 2단 분할(Splitter) 구조로 구성하며, 모든 입력 그룹은 흰색 배경(QGroupBox)과 파란색 포커스 테두리(#007bff)를 적용한다.

**A. 좌측 패널: 입고 및 수주 정보 (Intake Info)**
*   **관리번호**: `GR0101_123` 또는 `AS260101...` 형식 입력.
    *   **Logic**: 입력된 번호를 파싱하여 문서 연도(YY) 및 날짜 추출. 미래 날짜일 경우 작년으로 자동 보정 후 팝업 확인.
*   **담당 엔지니어**: 설정(`settings.json`)에 저장된 이름 자동 입력 (수정 가능).
*   **고객명**: 수신인 입력.
*   **매체 정보**:
    *   **매체 종류**: HDD(3.5/2.5), SSD, USB/SD, Server 등 (설정된 '담당 파트'에 따라 기본값 자동 선택).
    *   **파일 시스템**: NTFS, FAT32, exFAT, APFS, EXT4 등 선택.
*   **모델명 / S/N**: 텍스트 입력.
*   **매체 사진**: 이미지 파일(JPG, PNG) 업로드 및 즉시 미리보기(Preview) 제공.

**B. 우측 패널: 작업 내용 및 결과 (Work Details)**
*   **증상 (Symptom)**: 콤보박스 선택 + 텍스트 수정 가능 (`setEditable(True)`).
    *   **Data**: 설정 메뉴에서 프리셋 목록 편집 가능.
*   **복구 및 분석 요청 사항**: `QTextEdit`(다중 라인) 입력. (높이 약 60px)
*   **작업 로그 (Diagnosis Log)**:
    *   **기능**: 미리 정의된 리스트에서 항목 클릭 시, 하단 '상세 작업 내용' 창에 번호(1. 2. 3...)와 함께 자동 입력.
    *   **입력 방식**: 다중 선택(Multi-selection) 지원.
*   **상세 작업 내용**: 자동 입력된 텍스트를 직접 수정/보완 가능.
*   **최종 결론 (Conclusion)**:
    *   **기능**: 결과 프리셋 리스트 제공 (다중 선택 가능). 클릭 시 텍스트 자동 입력.
    *   **입력 방식**: `QTextEdit` 사용 (높이 약 80px).

**C. 하단 액션바 (Footer)**
*   **양식 선택**: Type 1 (기본) 등 템플릿 선택.
*   **PDF 옵션**:
    *   `[v] PDF 생성`: 체크 시 docx와 함께 pdf 생성.
    *   `[v] PDF 압축(암호)`: 체크 시 비밀번호 입력창 활성화 → PDF를 AES 암호화된 ZIP으로 압축.
*   **생성 버튼**: "작업내역서 생성 (Export)"

### 3.2. 환경설정 (Settings Dialog)
상단 메뉴 **[설정] > [환경설정]**을 통해 진입하며, 탭(Tab)으로 구분된다. 데이터는 `settings.json`에 영구 저장된다.

*   **내 정보**: 기본 엔지니어명, 담당 파트(HDD, SSD 등) 설정.
*   **증상 목록**: 메인 화면 콤보박스에 들어갈 항목 편집.
*   **진단/로그**: 작업 로그 프리셋 편집.
*   **최종 결론**: 결론 프리셋 편집.

### 3.3. 문서 생성 로직 (Output Logic)
*   **문서 번호 생성 규칙**:
    *   **포맷**: `MIT-AS-01-{YYMMDD}-{Suffix}`
    *   **YYMMDD**: 관리번호에서 날짜 추출. (`GR0101_123` → 2026년일 경우 `260101`)
    *   **Suffix**: 관리번호 뒤 3자리 또는 `_` 뒤 숫자.
*   **파일 저장 규칙**:
    *   **파일명**: `{관리번호}_작업내역서.docx` (PDF/ZIP도 동일 명칭)
    *   **저장 경로**: 사용자 지정 (`QFileDialog`).
*   **데이터 치환 (Template Mapping)**:
    *   `{{ doc_number }}`, `{{ today_date }}`, `{{ customer_name }}`, `{{ model_name }}`, `{{ sn }}`, `{{ symptom_type }}`, `{{ request_detail }}`, `{{ work_detail }}`, `{{ conclusion }}`, `{{ media_photo }}` 등.
    *   텍스트 입력란의 앞뒤 공백 및 불필요한 엔터(`\n`)는 자동으로 제거(`strip`) 후 삽입.

## 4. UI/UX 요구사항 (Design Guidelines)
*   **스타일링 (CSS)**:
    *   모든 `QTextEdit`, `QLineEdit`, `QComboBox`는 명확한 테두리(`border: 1px solid #a0a0a0`)를 가질 것.
    *   입력창 포커스 시 테두리는 파란색(`#007bff`)으로 강조.
    *   **배경색**: 메인 윈도우(`#f4f6f9`), 그룹박스(White).
*   **사용성**:
    *   엔터키로 줄바꿈되는 리스트 데이터 처리.
    *   설정 변경 시 메인 UI 즉시 반영 (Refresh).

## 5. 예외 처리 및 안정성 (Error Handling)
*   **Console-less Execution**: `sys.stdout`을 리다이렉션하여 `docx2pdf` 등 외부 라이브러리가 콘솔 출력을 시도할 때 발생하는 `NoneType` 에러 방지.
*   **Crash Logging**: 프로그램 강제 종료 시 `crash_log.txt`에 에러 스택 트레이스 기록.
*   **Path Resolution**: `EXE` 실행 위치 기준(`sys.executable`)으로 템플릿 및 설정 파일 경로 탐색.
*   **Hidden Imports**: PyInstaller 빌드 시 PyQt6 모듈 강제 포함.

## 6. 산출물 파일 구조 (File Structure)
배포 시 아래 파일들이 같은 폴더 내에 존재해야 함.

```plaintext
📁 ProgramFolder
 ├── 📄 JobReportGenerator.exe  (메인 실행 파일)
 ├── 📄 template_type1.docx     (필수: Word 템플릿)
 ├── 📄 settings.json           (자동생성: 설정 파일)
 └── 📄 crash_log.txt           (자동생성: 에러 발생 시)
```
