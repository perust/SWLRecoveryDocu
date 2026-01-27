# [중요] 최상단에서 콘솔 출력 차단 (NoneType 에러 방지용)
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTextEdit,
    QFileDialog,
    QCheckBox,
    QListWidget,
    QAbstractItemView,
    QSplitter,
    QMessageBox,
    QDialog,
    QDialogButtonBox,
    QTabWidget,
    QInputDialog)
from doc_generator import DocGenerator
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QAction
from datetime import datetime
import traceback
import re
import json
import os
import sys


class NullWriter:
    def write(self, text): pass
    def flush(self): pass


# 이미 Frozen 상태(exe)이거나 stdout이 없으면 강제 할당
if getattr(sys, 'frozen', False) or sys.stdout is None:
    sys.stdout = NullWriter()
    sys.stderr = NullWriter()


# exe 실행 위치 찾기 함수
def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))


BASE_PATH = get_base_path()
SETTINGS_FILE = os.path.join(BASE_PATH, "settings.json")

# 프로그램 충돌 로그 기록


def exception_hook(exctype, value, tb):
    error_msg = "".join(traceback.format_exception(exctype, value, tb))
    with open(os.path.join(BASE_PATH, "crash_log.txt"), "w",
              encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] CRASH REPORT\n")
        f.write(error_msg)
    sys.exit(1)


sys.excepthook = exception_hook


class SettingsDialog(QDialog):
    def __init__(self, parent=None, current_settings=None):
        super().__init__(parent)
        self.setWindowTitle("환경설정")
        self.resize(550, 650)
        self.settings = current_settings or {}

        # [UI 개선] 설정창 전용 스타일시트 (테두리 추가)
        self.setStyleSheet("""
            QTextEdit {
                border: 1px solid #a0a0a0;
                border-radius: 4px;
                background: #fff;
            }
            QLineEdit {
                border: 1px solid #a0a0a0;
                border-radius: 4px;
                padding: 4px;
                background: #fff;
            }
            QLabel {
                font-weight: bold;
                margin-top: 10px;
            }
        """)

        layout = QVBoxLayout()
        tabs = QTabWidget()

        # 탭 1: 내 정보
        tab_info = QWidget()
        layout_info = QVBoxLayout()
        layout_info.addWidget(QLabel("이름 (기본 엔지니어명)"))
        self.input_name = QLineEdit(self.settings.get("engineer_name", ""))
        layout_info.addWidget(self.input_name)
        layout_info.addWidget(QLabel("담당 파트"))
        self.combo_part = QComboBox()
        self.combo_part.addItems(
            ["HDD", "SSD", "SW", "Memory", "Mobile", "Server/NAS", "Other"])
        self.combo_part.setCurrentText(
            self.settings.get("engineer_part", "HDD"))
        layout_info.addWidget(self.combo_part)
        layout_info.addStretch()
        tab_info.setLayout(layout_info)

        # 탭 2: 증상 목록
        tab_symptom = QWidget()
        layout_symptom = QVBoxLayout()
        layout_symptom.addWidget(QLabel("증상 콤보박스 목록 (엔터로 구분)"))
        self.text_symptoms = QTextEdit()
        symptom_list = self.settings.get("symptom_list", [])
        self.text_symptoms.setPlainText("\n".join(symptom_list))
        layout_symptom.addWidget(self.text_symptoms)
        tab_symptom.setLayout(layout_symptom)

        # 탭 3: 진단 목록
        tab_diag = QWidget()
        layout_diag = QVBoxLayout()
        layout_diag.addWidget(QLabel("진단 자동입력 목록 (엔터로 구분)"))
        self.text_diagnosis = QTextEdit()
        diag_list = self.settings.get("diagnosis_list", [])
        self.text_diagnosis.setPlainText("\n".join(diag_list))
        layout_diag.addWidget(self.text_diagnosis)
        tab_diag.setLayout(layout_diag)

        # 탭 4: 결론 목록
        tab_conclusion = QWidget()
        layout_conclusion = QVBoxLayout()
        layout_conclusion.addWidget(QLabel("결론 자동입력 목록 (엔터로 구분)"))
        self.text_conclusion_list = QTextEdit()
        conclusion_list = self.settings.get("conclusion_list", [])
        self.text_conclusion_list.setPlainText("\n".join(conclusion_list))
        layout_conclusion.addWidget(self.text_conclusion_list)
        tab_conclusion.setLayout(layout_conclusion)

        tabs.addTab(tab_info, "내 정보")
        tabs.addTab(tab_symptom, "증상")
        tabs.addTab(tab_diag, "진단/로그")
        tabs.addTab(tab_conclusion, "최종 결론")

        layout.addWidget(tabs)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save |
            QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_data(self):
        s_list = [
            line.strip()
            for line in self.text_symptoms.toPlainText().split('\n')
            if line.strip()
        ]
        d_list = [
            line.strip()
            for line in self.text_diagnosis.toPlainText().split('\n')
            if line.strip()
        ]
        c_list = [
            line.strip()
            for line in self.text_conclusion_list.toPlainText().split('\n')
            if line.strip()
        ]
        return {
            "engineer_name": self.input_name.text(),
            "engineer_part": self.combo_part.currentText(),
            "symptom_list": s_list,
            "diagnosis_list": d_list,
            "conclusion_list": c_list
        }


class JobReportApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("데이터 복구 작업내역서 생성기 v1.6 (Fix)")
        self.resize(1150, 850)

        self.generator = DocGenerator()
        self.generator.base_path = BASE_PATH
        self.generator.template_path = os.path.join(
            BASE_PATH, "template_type1.docx")

        self.photo_path = None
        self.load_settings()
        self.create_menu_bar()

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.left_panel = self.create_intake_section()
        self.right_panel = self.create_work_section()

        splitter.addWidget(self.left_panel)
        splitter.addWidget(self.right_panel)
        splitter.setSizes([400, 750])

        self.main_layout.addWidget(splitter)
        self.create_footer_section()
        self.apply_styles()
        self.apply_media_type_by_part()

    def load_settings(self):
        default_settings = {
            "engineer_name": "",
            "engineer_part": "HDD",
            "symptom_list": [
                "직접 입력",
                "SW-Overwrite (덮어쓰기)",
                "SW-Format (포맷)",
                "SW-Delete (삭제)",
                "HW-Head Stuck",
                "HW-Bad Sector",
                "Ransomware"],
            "diagnosis_list": [
                "전원 인가 시 소음 발생",
                "PCB 쇼트 확인",
                "Firmware 손상",
                "Bad Sector 다수",
                "MFT 손상",
                "Head Stuck 해제"],
            "conclusion_list": [
                "복구 성공 / 데이터 무결성 확인",
                "복구 실패 (Media 손상)",
                "부분 복구 (일부 손상)",
                "요청 데이터 정상 복구"]}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    default_settings.update(saved)
            except Exception:
                pass
        self.settings = default_settings

    def save_settings(self):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def create_menu_bar(self):
        menubar = self.menuBar()
        settings_menu = menubar.addMenu("설정(Settings)")
        pref_action = QAction("환경설정...", self)
        pref_action.triggered.connect(self.open_settings)
        settings_menu.addAction(pref_action)

    def open_settings(self):
        dialog = SettingsDialog(self, self.settings)
        if dialog.exec():
            self.settings.update(dialog.get_data())
            self.save_settings()
            self.refresh_ui_from_settings()

    def refresh_ui_from_settings(self):
        self.input_engineer.setText(self.settings["engineer_name"])
        self.combo_symptom_type.clear()
        self.combo_symptom_type.addItems(self.settings["symptom_list"])
        self.symptom_list.clear()
        self.symptom_list.addItems(self.settings["diagnosis_list"])
        self.conclusion_list_widget.clear()
        self.conclusion_list_widget.addItems(
            self.settings["conclusion_list"])
        self.apply_media_type_by_part()

    def apply_media_type_by_part(self):
        part = self.settings.get("engineer_part", "HDD")
        target = "HDD (3.5)"
        if part == "SSD":
            target = "SSD"
        elif part in ["Memory", "Mobile"]:
            target = "USB/SD"
        elif part == "Server/NAS":
            target = "Server"

        idx = self.combo_media_type.findText(
            target, Qt.MatchFlag.MatchContains)
        if idx >= 0:
            self.combo_media_type.setCurrentIndex(idx)

    def create_intake_section(self):
        group = QGroupBox("1. 입고 및 수주 정보")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("관리번호 (예: GR0101_123)"))
        self.input_order_id = QLineEdit()
        layout.addWidget(self.input_order_id)
        layout.addWidget(QLabel("담당 엔지니어"))
        self.input_engineer = QLineEdit()
        self.input_engineer.setText(self.settings.get("engineer_name", ""))
        layout.addWidget(self.input_engineer)
        layout.addWidget(QLabel("고객명"))
        self.input_customer = QLineEdit()
        layout.addWidget(self.input_customer)
        layout.addWidget(QLabel("매체 종류 / 파일 시스템"))
        row_media = QHBoxLayout()
        self.combo_media_type = QComboBox()
        self.combo_media_type.addItems(
            ["HDD (3.5)", "HDD (2.5)", "SSD", "USB/SD", "Server"])
        self.combo_fs = QComboBox()
        self.combo_fs.addItems(
            ["NTFS", "FAT32", "exFAT", "APFS", "HFS+", "EXT4"])
        row_media.addWidget(self.combo_media_type)
        row_media.addWidget(self.combo_fs)
        layout.addLayout(row_media)
        layout.addWidget(QLabel("모델명 / S/N"))
        self.input_model = QLineEdit()
        self.input_sn = QLineEdit()
        self.input_model.setPlaceholderText("Model Name")
        self.input_sn.setPlaceholderText("Serial Number")
        layout.addWidget(self.input_model)
        layout.addWidget(self.input_sn)

        layout.addWidget(QLabel("매체 사진"))
        btn = QPushButton("사진 찾기")
        btn.clicked.connect(self.load_photo)
        layout.addWidget(btn)
        self.lbl_photo_preview = QLabel("미리보기")
        self.lbl_photo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_photo_preview.setStyleSheet(
            "border: 1px dashed #aaa; background: #f9f9f9;")
        self.lbl_photo_preview.setFixedHeight(150)
        layout.addWidget(self.lbl_photo_preview)
        layout.addStretch()
        group.setLayout(layout)
        return group

    def create_work_section(self):
        group = QGroupBox("2. 작업 내용 및 결과")
        layout = QVBoxLayout()

        layout.addWidget(QLabel("증상"))
        self.combo_symptom_type = QComboBox()
        self.combo_symptom_type.setEditable(True)
        self.combo_symptom_type.addItems(
            self.settings.get("symptom_list", []))
        layout.addWidget(self.combo_symptom_type)

        layout.addWidget(QLabel("요청 사항"))
        self.input_request_detail = QTextEdit()
        self.input_request_detail.setFixedHeight(60)
        layout.addWidget(self.input_request_detail)

        layout.addWidget(QLabel("▼ 작업 로그 (다중 선택)"))
        self.symptom_list = QListWidget()
        self.symptom_list.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection)
        self.symptom_list.addItems(self.settings.get("diagnosis_list", []))
        self.symptom_list.setFixedHeight(100)
        self.symptom_list.itemClicked.connect(self.build_work_detail)
        layout.addWidget(self.symptom_list)

        self.text_work_detail = QTextEdit()
        layout.addWidget(self.text_work_detail)

        # [수정됨] 최종 결론 다중 선택 가능하도록 변경
        layout.addWidget(QLabel("▼ 최종 결론 (다중 선택)"))
        self.conclusion_list_widget = QListWidget()
        self.conclusion_list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection)  # 다중선택
        self.conclusion_list_widget.addItems(
            self.settings.get("conclusion_list", []))
        self.conclusion_list_widget.setFixedHeight(80)
        self.conclusion_list_widget.itemClicked.connect(
            self.build_conclusion)  # 클릭 시 조합
        layout.addWidget(self.conclusion_list_widget)

        self.text_conclusion = QTextEdit()
        self.text_conclusion.setMinimumHeight(80)
        self.text_conclusion.setText("데이터 복구 성공 / 데이터 무결성 확인 완료")
        layout.addWidget(self.text_conclusion)

        group.setLayout(layout)
        return group

    def create_footer_section(self):
        layout = QHBoxLayout()
        self.chk_zip = QCheckBox("PDF 압축(암호)")
        self.chk_zip.toggled.connect(self.toggle_zip_password)
        layout.addWidget(self.chk_zip)
        self.input_zip_pw = QLineEdit()
        self.input_zip_pw.setPlaceholderText("비밀번호")
        self.input_zip_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.input_zip_pw.setFixedWidth(100)
        self.input_zip_pw.setEnabled(False)
        layout.addWidget(self.input_zip_pw)
        layout.addStretch()
        self.chk_pdf = QCheckBox("PDF 생성")
        self.chk_pdf.setChecked(True)
        layout.addWidget(self.chk_pdf)
        btn = QPushButton("문서 생성 (Export)")
        btn.setStyleSheet("""
            background-color: #007bff;
            color: white;
            font-weight: bold;
            padding: 10px;
        """)
        btn.clicked.connect(self.run_export)
        layout.addWidget(btn)
        self.main_layout.addLayout(layout)

    def toggle_zip_password(self, checked):
        self.input_zip_pw.setEnabled(checked)
        if checked:
            self.input_zip_pw.setFocus()
            self.chk_pdf.setChecked(True)

    def load_photo(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, '사진 선택', '', 'Images (*.png *.jpg *.jpeg)')
        if fname:
            self.photo_path = fname
            pix = QPixmap(fname)
            self.lbl_photo_preview.setPixmap(pix.scaled(
                self.lbl_photo_preview.size(),
                Qt.AspectRatioMode.KeepAspectRatio))

    def build_work_detail(self):
        items = self.symptom_list.selectedItems()
        txt = ""
        for i, item in enumerate(items, 1):
            txt += f"{i}. {item.text()}\n"
        self.text_work_detail.setText(txt)

    def build_conclusion(self):
        # [수정됨] 결론도 다중 선택 시 번호 매겨서 입력
        items = self.conclusion_list_widget.selectedItems()
        txt = ""
        # 하나만 선택했을 땐 번호 없이, 여러개면 번호 붙이기
        if len(items) == 1:
            txt = items[0].text()
        else:
            for i, item in enumerate(items, 1):
                txt += f"{i}. {item.text()}\n"
        self.text_conclusion.setText(txt)

    def run_export(self):
        order_id = self.input_order_id.text().strip()
        if not order_id:
            QMessageBox.warning(self, "경고", "관리번호를 입력하세요.")
            return

        calc_date = None
        match = re.match(r"GR(\d{4})_(\d+)", order_id, re.IGNORECASE)
        if match:
            date_digits = match.group(1)
            try:
                dt = datetime(
                    datetime.now().year,
                    int(date_digits[:2]),
                    int(date_digits[2:])
                )
                year = dt.year - 1 if dt > datetime.now() else dt.year
            except BaseException:
                year = datetime.now().year

            res, ok = QInputDialog.getInt(
                self, "연도 확인", f"문서 연도: {year}", year, 2000, 2100, 1)
            if not ok:
                return
            calc_date = f"{str(res)[2:]}{date_digits}"

        data = {
            'order_id': order_id,
            'engineer': self.input_engineer.text(),
            'customer': self.input_customer.text(),
            'model': self.input_model.text(),
            'sn': self.input_sn.text(),
            'filesystem': self.combo_fs.currentText(),
            'symptom_type': self.combo_symptom_type.currentText(),
            'request_detail': self.input_request_detail.toPlainText(),
            'symptom_chk': ", ".join(
                [i.text() for i in self.symptom_list.selectedItems()]),
            'work_content': self.text_work_detail.toPlainText(),
            'conclusion': self.text_conclusion.toPlainText(),
            'photo_path': self.photo_path,
            'doc_date_part': calc_date
        }

        zip_pw = self.input_zip_pw.text() if self.chk_zip.isChecked() else None

        save_path, _ = QFileDialog.getSaveFileName(
            self, "저장", f"{order_id}_작업내역서.docx", "Word Files (*.docx)")
        if save_path:
            success, msg = self.generator.generate_report(
                data, save_path, self.chk_pdf.isChecked(),
                zip_password=zip_pw)
            if success:
                QMessageBox.information(self, "성공", msg)
            else:
                QMessageBox.critical(self, "실패", msg)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f4f6f9; }
            QGroupBox {
                background: white;
                border: 1px solid #ccc;
                border-radius: 8px;
                margin-top: 10px;
                padding: 15px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QTextEdit, QLineEdit, QComboBox {
                border: 1px solid #a0a0a0;
                border-radius: 4px;
                padding: 6px;
                background: #fff;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 2px solid #007bff;
            }
        """)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = JobReportApp()
    window.show()
    sys.exit(app.exec())
