import os
import re
from datetime import datetime
from docxtpl import DocxTemplate

# 라이브러리 체크
try:
    from docx2pdf import convert
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import pyzipper
    ZIP_AVAILABLE = True
except ImportError:
    ZIP_AVAILABLE = False


class DocGenerator:
    def __init__(self):
        # 기본 경로는 main.py에서 주입받거나 현재 경로 사용
        self.base_path = os.getcwd()
        self.template_path = os.path.join(
            self.base_path, "template_type1.docx")

    def _make_doc_number(self, order_id, custom_date_part=None):
        try:
            if "_" in order_id:
                suffix = order_id.split("_")[-1]
            else:
                suffix = order_id[-3:] if len(order_id) >= 3 else "000"

            if not suffix.isdigit():
                suffix = "000"

            if custom_date_part:
                date_part = custom_date_part
            else:
                date_match = re.search(r'\d{6}', order_id)
                if date_match:
                    date_part = date_match.group(0)
                else:
                    date_part = datetime.now().strftime("%y%m%d")

            return f"MIT-AS-01-{date_part}-{suffix}"
        except Exception:
            return f"MIT-AS-01-{datetime.now().strftime('%y%m%d')}-000"

    def create_encrypted_zip(self, pdf_path, password):
        if not ZIP_AVAILABLE:
            return False, "pyzipper 라이브러리 미설치"

        try:
            zip_path = pdf_path.replace(".pdf", ".zip")
            with pyzipper.AESZipFile(
                zip_path,
                'w',
                compression=pyzipper.ZIP_LZMA,
                encryption=pyzipper.WZ_AES
            ) as zf:
                zf.setpassword(password.encode('utf-8'))
                zf.write(pdf_path, os.path.basename(pdf_path))
            return True, zip_path
        except Exception as e:
            return False, str(e)

    # [수정] zip_password 인자 추가 (에러 해결 핵심)
    def generate_report(
            self,
            ui_data,
            save_path,
            create_pdf=False,
            zip_password=None):
        if not os.path.exists(self.template_path):
            return False, (
                f"템플릿 파일을 찾을 수 없습니다:\n"
                f"{self.template_path}"
            )

        try:
            doc = DocxTemplate(self.template_path)
            doc_date = ui_data.get('doc_date_part')
            final_doc_num = self._make_doc_number(
                ui_data['order_id'], custom_date_part=doc_date)

            clean_work_detail = ui_data['work_content'].strip()
            clean_request_detail = ui_data['request_detail'].strip()

            context = {
                'doc_number': final_doc_num,
                'today_date': datetime.now().strftime("%Y년 %m월 %d일"),
                'management_id': ui_data['order_id'],
                'engineer_name': ui_data['engineer'],
                'customer_name': ui_data['customer'],
                'client_name': ui_data['customer'],
                'model_name': ui_data['model'],
                'serial_number': ui_data['sn'],
                'file_system': ui_data['filesystem'],
                'symptom_type': ui_data['symptom_type'],
                'request_detail': clean_request_detail,
                'symptom_initial': ui_data['symptom_chk'],
                'work_detail': clean_work_detail,
                'conclusion': ui_data['conclusion']
            }

            if ui_data.get('photo_path') and os.path.exists(
                    ui_data['photo_path']):
                from docxtpl import InlineImage
                from docx.shared import Mm
                context['media_photo'] = InlineImage(
                    doc, ui_data['photo_path'], width=Mm(50))
            else:
                context['media_photo'] = ""

            doc.render(context)
            doc.save(save_path)

            result_msg = f"Word 생성 완료: {os.path.basename(save_path)}"

            if create_pdf and PDF_AVAILABLE:
                pdf_path = save_path.replace(".docx", ".pdf")
                # PDF 변환 (에러 방지: docx2pdf 내부 print 억제)
                convert(save_path, pdf_path)
                result_msg += f"\nPDF 생성 완료: {os.path.basename(pdf_path)}"

                if zip_password:
                    zip_success, zip_res = self.create_encrypted_zip(
                        pdf_path, zip_password)
                    if zip_success:
                        result_msg += (
                            f"\nZIP 압축 완료(암호): "
                            f"{os.path.basename(zip_res)}"
                        )
                    else:
                        result_msg += f"\nZIP 압축 실패: {zip_res}"

            return True, result_msg

        except Exception as e:
            return False, f"에러 발생: {str(e)}"
