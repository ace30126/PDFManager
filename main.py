import sys
import os
import fitz # PyMuPDF 라이브러리 임포트
from PyPDF2 import PdfReader, PdfWriter, PdfMerger

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox,
    QFileDialog, QInputDialog, QLineEdit, QStackedWidget, QScrollArea, QLabel,
    QSizePolicy, QFrame, QGridLayout, QProgressBar, QStatusBar
)
from PyQt6.QtCore import Qt, QSize, QDir

class PDFEditorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.current_pdf_path = None # 현재 작업 중인 PDF 파일 경로
        self.current_pdf_doc = None # PyMuPDF Document 객체
        self.initUI()

    def initUI(self):
        # 창 제목 설정
        self.setWindowTitle("PDF 편집기")
        self.setGeometry(100, 100, 800, 600) # 창 크기 확장

        # 메인 레이아웃 (QStackedWidget과 QStatusBar를 포함)
        main_v_layout = QVBoxLayout(self)
        self.setLayout(main_v_layout)

        # QStackedWidget 생성: 여러 화면을 전환할 때 사용
        self.stacked_widget = QStackedWidget(self)
        main_v_layout.addWidget(self.stacked_widget)

        # 상태 바 생성
        self.status_bar = QStatusBar()
        main_v_layout.addWidget(self.status_bar)
        self.status_bar.showMessage("준비 완료")

        # 메인 메뉴 화면 설정
        self._setup_main_menu_page()
        # 페이지 미리보기 및 작업 화면 설정
        self._setup_operation_page()

        # 초기 화면을 메인 메뉴로 설정
        self.stacked_widget.setCurrentIndex(0)

    def _setup_main_menu_page(self):
        """메인 메뉴 화면 (버튼들)을 설정합니다."""
        main_menu_page = QWidget()
        # QGridLayout을 사용하여 버튼을 더 크고 정돈되게 배치
        grid_layout = QGridLayout()
        grid_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        grid_layout.setSpacing(20) # 버튼 사이 간격

        # 각 기능 버튼 생성
        self.btn_merge_files = QPushButton("PDF 파일 합치기")
        self.btn_merge_folder = QPushButton("폴더 내 PDF 합치기") # 새 기능 버튼
        self.btn_extract = QPushButton("페이지 추출")
        self.btn_delete_reorder = QPushButton("페이지 삭제/순서 변경")
        self.btn_unlock = QPushButton("암호 해제")
        self.btn_add_cover = QPushButton("표지 추가")
        self.btn_extract_text = QPushButton("텍스트 추출")

        # 버튼 스타일 및 크기 정책 설정
        buttons = [
            self.btn_merge_files, self.btn_merge_folder, self.btn_extract,
            self.btn_delete_reorder, self.btn_unlock, self.btn_add_cover,
            self.btn_extract_text
        ]
        for btn in buttons:
            btn.setMinimumHeight(60) # 버튼 높이 증가
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            btn.setStyleSheet("""
                QPushButton {
                    font-size: 18pt;
                    padding: 15px;
                    border-radius: 15px;
                    background-color: #4CAF50;
                    color: white;
                    border: 2px solid #388E3C;
                }
                QPushButton:hover {
                    background-color: #5cb85c;
                }
                QPushButton:pressed {
                    background-color: #388E3C;
                    border-style: inset;
                }
            """)

        # 버튼을 그리드 레이아웃에 추가
        grid_layout.addWidget(self.btn_merge_files, 0, 0)
        grid_layout.addWidget(self.btn_merge_folder, 0, 1)
        grid_layout.addWidget(self.btn_extract, 1, 0)
        grid_layout.addWidget(self.btn_delete_reorder, 1, 1)
        grid_layout.addWidget(self.btn_unlock, 2, 0)
        grid_layout.addWidget(self.btn_add_cover, 2, 1)
        grid_layout.addWidget(self.btn_extract_text, 3, 0, 1, 2) # 텍스트 추출 버튼은 두 컬럼에 걸쳐 배치

        main_menu_page.setLayout(grid_layout)
        self.stacked_widget.addWidget(main_menu_page) # 스택 위젯에 메인 메뉴 페이지 추가

        # 버튼 클릭 시그널 연결
        self.btn_merge_files.clicked.connect(self.merge_pdfs_files)
        self.btn_merge_folder.clicked.connect(self.merge_pdfs_from_folder)
        self.btn_extract.clicked.connect(self._start_page_operation_extract)
        self.btn_delete_reorder.clicked.connect(self._start_page_operation_delete_reorder)
        self.btn_unlock.clicked.connect(self.unlock_pdf)
        self.btn_add_cover.clicked.connect(self.add_cover)
        self.btn_extract_text.clicked.connect(self.extract_text)

    def _setup_operation_page(self):
        """페이지 미리보기 및 작업 화면을 설정합니다."""
        self.operation_page = QWidget()
        main_h_layout = QHBoxLayout()

        # 왼쪽: PDF 페이지 미리보기 영역 (스크롤 가능)
        left_layout = QVBoxLayout()
        self.preview_scroll_area = QScrollArea()
        self.preview_scroll_area.setWidgetResizable(True)
        self.preview_scroll_area.setFrameShape(QFrame.Shape.StyledPanel)
        self.preview_contents_widget = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_contents_widget)
        self.preview_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.preview_scroll_area.setWidget(self.preview_contents_widget)
        left_layout.addWidget(self.preview_scroll_area)

        # 미리보기 로딩 진행률 바
        self.preview_progress_bar = QProgressBar(self)
        self.preview_progress_bar.setTextVisible(True)
        self.preview_progress_bar.setFormat("페이지 로딩 중... %p%")
        self.preview_progress_bar.hide() # 초기에는 숨김
        left_layout.addWidget(self.preview_progress_bar)

        main_h_layout.addLayout(left_layout, 3) # 미리보기 영역에 더 많은 공간 할당

        # 오른쪽: 작업 입력 및 버튼 영역
        right_layout = QVBoxLayout()
        right_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.current_file_label = QLabel("선택된 파일: 없음")
        self.current_file_label.setWordWrap(True)
        self.current_file_label.setStyleSheet("font-size: 14pt; font-weight: bold;")
        right_layout.addWidget(self.current_file_label)

        self.operation_label = QLabel("작업 지시:")
        self.operation_label.setStyleSheet("font-size: 12pt; margin-top: 10px;")
        right_layout.addWidget(self.operation_label)

        self.input_line_edit = QLineEdit()
        self.input_line_edit.setPlaceholderText("페이지 범위 또는 순서 입력")
        self.input_line_edit.setStyleSheet("font-size: 12pt; padding: 5px; border-radius: 5px;")
        right_layout.addWidget(self.input_line_edit)

        self.apply_button = QPushButton("적용")
        self.apply_button.setStyleSheet("""
            QPushButton {
                font-size: 14pt;
                padding: 10px;
                border-radius: 10px;
                background-color: #2196F3;
                color: white;
                border: 2px solid #1976D2;
            }
            QPushButton:hover {
                background-color: #42a5f5;
            }
            QPushButton:pressed {
                background-color: #1976D2;
                border-style: inset;
            }
        """)
        self.apply_button.clicked.connect(self._apply_page_operation)
        right_layout.addWidget(self.apply_button)

        self.back_button = QPushButton("메인 메뉴로 돌아가기")
        self.back_button.setStyleSheet("""
            QPushButton {
                font-size: 14pt;
                padding: 10px;
                border-radius: 10px;
                background-color: #f44336;
                color: white;
                border: 2px solid #d32f2f;
            }
            QPushButton:hover {
                background-color: #ef5350;
            }
            QPushButton:pressed {
                background-color: #d32f2f;
                border-style: inset;
            }
        """)
        self.back_button.clicked.connect(self._go_to_main_menu)
        right_layout.addWidget(self.back_button)

        right_layout.addStretch(1) # 하단에 공간 추가

        main_h_layout.addLayout(right_layout, 1) # 작업 영역에 공간 할당

        self.operation_page.setLayout(main_h_layout)
        self.stacked_widget.addWidget(self.operation_page) # 스택 위젯에 작업 페이지 추가

    def _go_to_main_menu(self):
        """메인 메뉴 화면으로 돌아갑니다."""
        self.stacked_widget.setCurrentIndex(0)
        # 작업 후 상태 초기화
        self.current_pdf_path = None
        if self.current_pdf_doc:
            self.current_pdf_doc.close()
            self.current_pdf_doc = None
        self._clear_preview()
        self.input_line_edit.clear()
        self.current_file_label.setText("선택된 파일: 없음")
        self.status_bar.showMessage("준비 완료")
        self.preview_progress_bar.hide()


    def _clear_preview(self):
        """미리보기 영역의 모든 페이지 이미지를 제거합니다."""
        while self.preview_layout.count():
            item = self.preview_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _load_and_display_pdf_preview(self, pdf_path):
        """
        주어진 PDF 파일의 페이지들을 미리보기 영역에 표시합니다.
        """
        self._clear_preview() # 기존 미리보기 지우기
        self.current_file_label.setText(f"선택된 파일: {os.path.basename(pdf_path)}")
        self.preview_progress_bar.show()
        self.status_bar.showMessage("페이지 미리보기 로딩 중...")

        try:
            self.current_pdf_doc = fitz.open(pdf_path) # PyMuPDF Document 열기
            total_pages = len(self.current_pdf_doc)
            self.preview_progress_bar.setMaximum(total_pages)

            for i, page in enumerate(self.current_pdf_doc):
                # 페이지를 픽스맵으로 렌더링 (해상도 조절 가능)
                # 높은 해상도는 메모리 사용량과 로딩 시간을 증가시킬 수 있습니다.
                pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8)) # 미리보기 크기 조절

                # QImage로 변환
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)

                # QLabel에 페이지 번호와 이미지 표시
                page_label = QLabel(f"Page {i + 1}")
                page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                page_label.setStyleSheet("font-weight: bold; margin-top: 5px;")
                image_label = QLabel()
                image_label.setPixmap(pixmap)
                image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                image_label.setFixedSize(pixmap.size()) # 이미지 크기에 맞춰 라벨 크기 고정
                image_label.setFrameShape(QFrame.Shape.StyledPanel) # 프레임 추가

                # 미리보기 레이아웃에 추가
                self.preview_layout.addWidget(page_label)
                self.preview_layout.addWidget(image_label)
                self.preview_layout.addSpacing(10)

                self.preview_progress_bar.setValue(i + 1) # 진행률 업데이트
                QApplication.processEvents() # UI 업데이트를 위한 이벤트 처리

            self.status_bar.showMessage(f"{total_pages} 페이지 미리보기 로딩 완료.")

        except Exception as e:
            QMessageBox.critical(self, "미리보기 오류", f"PDF 미리보기를 로드하는 중 오류가 발생했습니다: {e}")
            self._go_to_main_menu() # 오류 발생 시 메인 메뉴로 돌아가기
        finally:
            self.preview_progress_bar.hide()


    # --- 각 기능별 메소드 ---

    def show_coming_soon_message(self, feature_name):
        """아직 구현되지 않은 기능에 대한 메시지 박스를 표시합니다."""
        QMessageBox.information(self, "기능 준비 중", f"{feature_name} 기능은 현재 준비 중입니다.")
        self.status_bar.showMessage(f"{feature_name} 기능 준비 중...")

    def _merge_pdfs_logic(self, file_paths, save_path):
        """실제 PDF 합치기 로직을 수행합니다."""
        if not file_paths:
            return False, "합칠 파일이 없습니다."
        if len(file_paths) < 2:
            return False, "두 개 이상의 PDF 파일을 선택해야 합니다."
        if not save_path:
            return False, "저장 경로가 지정되지 않았습니다."

        self.status_bar.showMessage("PDF 합치기 작업 시작...")
        QApplication.processEvents()

        try:
            pdf_merger = PdfMerger()
            for i, file_path in enumerate(file_paths):
                self.status_bar.showMessage(f"파일 추가 중: {os.path.basename(file_path)} ({i+1}/{len(file_paths)})")
                QApplication.processEvents()
                pdf_merger.append(file_path)

            with open(save_path, 'wb') as output_pdf:
                pdf_merger.write(output_pdf)
            pdf_merger.close()
            self.status_bar.showMessage("PDF 합치기 완료.")
            return True, "PDF 합치기가 완료되었습니다."
        except Exception as e:
            self.status_bar.showMessage("PDF 합치기 오류 발생.")
            return False, f"PDF 합치기 중 오류가 발생했습니다: {e}"

    def merge_pdfs_files(self):
        """
        사용자가 여러 PDF 파일을 선택하고, 이를 하나의 PDF 파일로 합쳐 저장합니다.
        """
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "합칠 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )

        if not file_paths:
            self.status_bar.showMessage("PDF 파일 합치기 취소됨.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "합쳐진 PDF 저장", "merged_files.pdf", "PDF 파일 (*.pdf)"
        )

        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        success, message = self._merge_pdfs_logic(file_paths, save_path)
        if success:
            QMessageBox.information(self, "작업 완료", message)
        else:
            QMessageBox.critical(self, "오류 발생", message)
        self.status_bar.showMessage("준비 완료")


    def merge_pdfs_from_folder(self):
        """
        사용자가 폴더를 선택하고, 해당 폴더 내의 모든 PDF 파일을 재귀적으로 합쳐 저장합니다.
        """
        folder_path = QFileDialog.getExistingDirectory(
            self, "PDF를 합칠 폴더 선택", ""
        )

        if not folder_path:
            self.status_bar.showMessage("폴더 선택 취소됨.")
            return

        self.status_bar.showMessage(f"폴더 '{os.path.basename(folder_path)}'에서 PDF 파일 검색 중...")
        QApplication.processEvents()

        pdf_files = self._find_pdf_files_recursive(folder_path)

        if not pdf_files:
            QMessageBox.warning(self, "파일 없음", "선택한 폴더 및 하위 폴더에서 PDF 파일을 찾을 수 없습니다.")
            self.status_bar.showMessage("PDF 파일 없음.")
            return

        self.status_bar.showMessage(f"{len(pdf_files)}개의 PDF 파일 발견. 합칠 파일 선택 중...")
        QApplication.processEvents()

        # 합쳐진 파일을 저장할 위치 및 파일명 선택 다이얼로그
        save_path, _ = QFileDialog.getSaveFileName(
            self, "합쳐진 PDF 저장", "merged_folder.pdf", "PDF 파일 (*.pdf)"
        )

        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        success, message = self._merge_pdfs_logic(pdf_files, save_path)
        if success:
            QMessageBox.information(self, "작업 완료", message)
        else:
            QMessageBox.critical(self, "오류 발생", message)
        self.status_bar.showMessage("준비 완료")

    def _find_pdf_files_recursive(self, folder_path):
        """
        주어진 폴더 및 모든 하위 폴더에서 PDF 파일을 찾아 리스트로 반환합니다.
        """
        pdf_files = []
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        return sorted(pdf_files) # 파일 순서를 위해 정렬

    def _start_page_operation_extract(self):
        """페이지 추출 작업을 시작합니다 (파일 선택 및 미리보기 로드)."""
        source_path, _ = QFileDialog.getOpenFileName(
            self, "원본 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )
        if not source_path:
            self.status_bar.showMessage("페이지 추출 취소됨.")
            return

        self.current_pdf_path = source_path
        self.operation_mode = "extract" # 작업 모드 설정
        self.operation_label.setText("추출할 페이지 범위를 입력하세요 (예: 1,3-5,7):")
        self.input_line_edit.setPlaceholderText("예: 1,3-5,7")
        self.stacked_widget.setCurrentIndex(1) # 작업 페이지로 전환
        self._load_and_display_pdf_preview(source_path)

    def _start_page_operation_delete_reorder(self):
        """페이지 삭제/순서 변경 작업을 시작합니다 (파일 선택 및 미리보기 로드)."""
        source_path, _ = QFileDialog.getOpenFileName(
            self, "원본 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )
        if not source_path:
            self.status_bar.showMessage("페이지 삭제/순서 변경 취소됨.")
            return

        self.current_pdf_path = source_path
        self.operation_mode = "delete_reorder" # 작업 모드 설정
        self.operation_label.setText("삭제하거나 순서를 변경할 페이지 번호를 입력하세요 (예: 삭제: 2,4 / 순서 변경: 5,1,3,2):")
        self.input_line_edit.setPlaceholderText("예: 삭제: 2,4 / 순서 변경: 5,1,3,2")
        self.stacked_widget.setCurrentIndex(1) # 작업 페이지로 전환
        self._load_and_display_pdf_preview(source_path)

    def _apply_page_operation(self):
        """현재 설정된 작업 모드에 따라 페이지 작업을 적용합니다."""
        if not self.current_pdf_path:
            QMessageBox.warning(self, "오류", "먼저 PDF 파일을 선택하세요.")
            self.status_bar.showMessage("작업할 PDF 파일이 선택되지 않았습니다.")
            return

        if self.operation_mode == "extract":
            self._execute_extract_pages()
        elif self.operation_mode == "delete_reorder":
            self._execute_delete_reorder_pages()
        else:
            QMessageBox.warning(self, "오류", "알 수 없는 작업 모드입니다.")
            self.status_bar.showMessage("알 수 없는 작업 모드.")

    def _execute_extract_pages(self):
        """페이지 추출 기능을 실행합니다."""
        page_range_text = self.input_line_edit.text()
        if not page_range_text:
            QMessageBox.warning(self, "입력 오류", "추출할 페이지 범위를 입력하세요.")
            self.status_bar.showMessage("페이지 범위 입력 필요.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "추출된 PDF 저장", "extracted.pdf", "PDF 파일 (*.pdf)"
        )
        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        self.status_bar.showMessage("페이지 추출 작업 시작...")
        QApplication.processEvents()

        try:
            reader = PdfReader(self.current_pdf_path)
            writer = PdfWriter()

            pages_to_extract = self._parse_page_range(page_range_text, len(reader.pages))

            if not pages_to_extract:
                QMessageBox.warning(self, "입력 오류", "잘못된 페이지 범위 형식입니다. 유효한 페이지를 찾을 수 없습니다.")
                self.status_bar.showMessage("잘못된 페이지 범위 형식.")
                return

            for page_num in pages_to_extract:
                if 0 <= page_num < len(reader.pages):
                    writer.add_page(reader.pages[page_num])
                else:
                    QMessageBox.warning(self, "경고", f"입력된 페이지 번호 {page_num + 1}은(는) PDF의 범위를 벗어납니다. 해당 페이지는 건너뜁니다.")
                    self.status_bar.showMessage(f"페이지 {page_num + 1} 건너뜀 (범위 벗어남).")


            if not writer.pages:
                QMessageBox.warning(self, "작업 실패", "추출할 페이지가 없습니다. 올바른 페이지 범위를 입력했는지 확인하세요.")
                self.status_bar.showMessage("추출할 페이지 없음.")
                return

            with open(save_path, 'wb') as output_pdf:
                writer.write(output_pdf)

            QMessageBox.information(self, "작업 완료", "페이지 추출이 완료되었습니다.")
            self.status_bar.showMessage("페이지 추출 완료.")
            self._go_to_main_menu() # 작업 완료 후 메인 메뉴로 돌아가기

        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"페이지 추출 중 오류가 발생했습니다: {e}")
            self.status_bar.showMessage("페이지 추출 오류 발생.")

    def _execute_delete_reorder_pages(self):
        """페이지 삭제/순서 변경 기능을 실행합니다."""
        input_text = self.input_line_edit.text()
        if not input_text:
            QMessageBox.warning(self, "입력 오류", "삭제하거나 순서를 변경할 페이지 번호를 입력하세요.")
            self.status_bar.showMessage("페이지 삭제/순서 변경 입력 필요.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "수정된 PDF 저장", "modified.pdf", "PDF 파일 (*.pdf)"
        )
        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        self.status_bar.showMessage("페이지 삭제/순서 변경 작업 시작...")
        QApplication.processEvents()

        try:
            reader = PdfReader(self.current_pdf_path)
            total_pages = len(reader.pages)
            writer = PdfWriter()

            pages_to_keep = list(range(total_pages)) # 기본적으로 모든 페이지를 유지

            delete_pages_str = ""
            reorder_pages_str = ""

            # 입력 파싱: "삭제: ..." 와 "순서 변경: ..." 구분
            if "삭제:" in input_text:
                delete_part = input_text.split("삭제:")[1].split("/")[0].strip()
                delete_pages_str = delete_part.split("순서 변경:")[0].strip()

            if "순서 변경:" in input_text:
                reorder_pages_str = input_text.split("순서 변경:")[1].strip()

            # 1. 페이지 삭제 처리
            if delete_pages_str:
                delete_indices = set()
                for p_str in delete_pages_str.split(','):
                    try:
                        page_num = int(p_str.strip())
                        if 1 <= page_num <= total_pages:
                            delete_indices.add(page_num - 1) # 0-인덱스로 변환
                        else:
                            QMessageBox.warning(self, "경고", f"삭제할 페이지 번호 {page_num}은(는) PDF 범위를 벗어납니다. 건너뜁니다.")
                            self.status_bar.showMessage(f"삭제 페이지 {page_num} 건너뜀 (범위 벗어남).")
                    except ValueError:
                        QMessageBox.warning(self, "입력 오류", f"잘못된 삭제 페이지 형식: '{p_str}'. 숫자를 입력하세요.")
                        self.status_bar.showMessage("잘못된 삭제 페이지 형식.")
                        return

                # 삭제할 페이지를 제외한 페이지들만 남김
                pages_to_keep = [i for i in pages_to_keep if i not in delete_indices]
                self.status_bar.showMessage(f"{len(delete_indices)} 페이지 삭제됨.")
                QApplication.processEvents()


            # 2. 페이지 순서 변경 처리
            if reorder_pages_str:
                new_order_indices = []
                # 현재 남아있는 페이지들 (삭제 후)의 1-인덱스 번호를 맵핑
                current_page_map = {i + 1: page_idx for page_idx, i in enumerate(pages_to_keep)}

                for p_str in reorder_pages_str.split(','):
                    try:
                        page_num = int(p_str.strip())
                        if page_num in current_page_map:
                            new_order_indices.append(pages_to_keep[current_page_map[page_num]])
                        else:
                            QMessageBox.warning(self, "경고", f"순서 변경할 페이지 번호 {page_num}은(는) 유효하지 않거나 이미 삭제되었습니다. 건너뜁니다.")
                            self.status_bar.showMessage(f"순서 변경 페이지 {page_num} 건너뜀 (유효하지 않음).")
                            return # 잘못된 순서 변경 입력은 전체 취소
                    except ValueError:
                        QMessageBox.warning(self, "입력 오류", f"잘못된 순서 변경 페이지 형식: '{p_str}'. 숫자를 입력하세요.")
                        self.status_bar.showMessage("잘못된 순서 변경 페이지 형식.")
                        return

                if len(set(new_order_indices)) != len(pages_to_keep):
                    QMessageBox.warning(self, "입력 오류", "순서 변경 페이지 번호가 중복되거나 누락되었습니다. 모든 페이지를 정확히 한 번씩 지정해야 합니다.")
                    self.status_bar.showMessage("순서 변경 페이지 번호 오류.")
                    return

                pages_to_keep = new_order_indices # 새 순서로 업데이트
                self.status_bar.showMessage("페이지 순서 변경됨.")
                QApplication.processEvents()


            # 최종 페이지들을 writer에 추가
            for page_idx in pages_to_keep:
                writer.add_page(reader.pages[page_idx])

            if not writer.pages:
                QMessageBox.warning(self, "작업 실패", "수정 후 남은 페이지가 없습니다. 올바른 페이지를 지정했는지 확인하세요.")
                self.status_bar.showMessage("수정 후 남은 페이지 없음.")
                return

            with open(save_path, 'wb') as output_pdf:
                writer.write(output_pdf)

            QMessageBox.information(self, "작업 완료", "페이지 삭제/순서 변경이 완료되었습니다.")
            self.status_bar.showMessage("페이지 삭제/순서 변경 완료.")
            self._go_to_main_menu() # 작업 완료 후 메인 메뉴로 돌아가기

        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"페이지 삭제/순서 변경 중 오류가 발생했습니다: {e}")
            self.status_bar.showMessage("페이지 삭제/순서 변경 오류 발생.")

    def _parse_page_range(self, page_range_str, total_pages):
        """
        페이지 범위 문자열(예: '1,3-5,7')을 파싱하여 0-인덱스 페이지 번호 리스트를 반환합니다.
        """
        pages = set()
        parts = page_range_str.replace(" ", "").split(',')
        for part in parts:
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    # 사용자 입력은 1부터 시작하므로 0-인덱스로 변환
                    pages.update(range(start - 1, end))
                except ValueError:
                    return [] # 잘못된 형식
            else:
                try:
                    page_num = int(part)
                    pages.add(page_num - 1) # 0-인덱스로 변환
                except ValueError:
                    return [] # 잘못된 형식
        # 유효한 페이지 번호만 필터링하고 정렬
        valid_pages = sorted([p for p in list(pages) if 0 <= p < total_pages])
        return valid_pages

    def unlock_pdf(self):
        """
        암호가 걸린 PDF 파일을 선택하고, 암호를 입력받아 잠금을 해제하여 새 파일로 저장합니다.
        """
        source_path, _ = QFileDialog.getOpenFileName(
            self, "암호 해제할 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )

        if not source_path:
            self.status_bar.showMessage("암호 해제 취소됨.")
            return

        password, ok = QInputDialog.getText(
            self, "암호 입력", "PDF 암호를 입력하세요:", QLineEdit.EchoMode.Password, ""
        )

        if not ok:
            self.status_bar.showMessage("암호 입력 취소됨.")
            return # 암호 입력 취소

        save_path, _ = QFileDialog.getSaveFileName(
            self, "암호 해제된 PDF 저장", "unlocked.pdf", "PDF 파일 (*.pdf)"
        )

        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        self.status_bar.showMessage("PDF 암호 해제 작업 시작...")
        QApplication.processEvents()

        try:
            reader = PdfReader(source_path)
            if reader.is_encrypted:
                if reader.decrypt(password):
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)

                    with open(save_path, 'wb') as output_pdf:
                        writer.write(output_pdf)
                    QMessageBox.information(self, "작업 완료", "PDF 암호 해제가 완료되었습니다.")
                    self.status_bar.showMessage("PDF 암호 해제 완료.")
                else:
                    QMessageBox.warning(self, "암호 오류", "암호가 잘못되었습니다.")
                    self.status_bar.showMessage("암호 오류: 암호가 잘못되었습니다.")
            else:
                QMessageBox.information(self, "정보", "선택한 PDF 파일은 암호화되어 있지 않습니다.")
                self.status_bar.showMessage("선택한 PDF 파일은 암호화되어 있지 않습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"PDF 암호 해제 중 오류가 발생했습니다: {e}")
            self.status_bar.showMessage("PDF 암호 해제 오류 발생.")
        finally:
            self.status_bar.showMessage("준비 완료")


    def add_cover(self):
        """
        표지로 사용할 PDF 파일(1페이지)을 선택하고, 본문 PDF의 맨 앞에 합쳐 새 파일로 저장합니다.
        """
        cover_path, _ = QFileDialog.getOpenFileName(
            self, "표지 PDF 파일 선택 (1페이지 권장)", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )
        if not cover_path:
            self.status_bar.showMessage("표지 PDF 선택 취소됨.")
            return

        main_pdf_path, _ = QFileDialog.getOpenFileName(
            self, "본문 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )
        if not main_pdf_path:
            self.status_bar.showMessage("본문 PDF 선택 취소됨.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "표지가 추가된 PDF 저장", "with_cover.pdf", "PDF 파일 (*.pdf)"
        )
        if not save_path:
            self.status_bar.showMessage("PDF 저장 취소됨.")
            return

        self.status_bar.showMessage("표지 추가 작업 시작...")
        QApplication.processEvents()

        try:
            merger = PdfMerger()
            merger.append(cover_path) # 표지 먼저 추가
            merger.append(main_pdf_path) # 그 다음 본문 추가

            with open(save_path, 'wb') as output_pdf:
                merger.write(output_pdf)
            merger.close()

            QMessageBox.information(self, "작업 완료", "표지 추가가 완료되었습니다.")
            self.status_bar.showMessage("표지 추가 완료.")
        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"표지 추가 중 오류가 발생했습니다: {e}")
            self.status_bar.showMessage("표지 추가 오류 발생.")
        finally:
            self.status_bar.showMessage("준비 완료")

    def extract_text(self):
        """
        선택된 PDF 파일의 모든 텍스트를 추출하여 .txt 파일로 저장합니다.
        """
        source_path, _ = QFileDialog.getOpenFileName(
            self, "텍스트를 추출할 PDF 파일 선택", "", "PDF 파일 (*.pdf);;모든 파일 (*)"
        )

        if not source_path:
            self.status_bar.showMessage("텍스트 추출 취소됨.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "텍스트 파일 저장", "extracted_text.txt", "텍스트 파일 (*.txt);;모든 파일 (*)"
        )

        if not save_path:
            self.status_bar.showMessage("텍스트 파일 저장 취소됨.")
            return

        self.status_bar.showMessage("텍스트 추출 작업 시작...")
        QApplication.processEvents()

        try:
            reader = PdfReader(source_path)
            full_text = ""
            total_pages = len(reader.pages)
            for i, page in enumerate(reader.pages):
                full_text += page.extract_text() + "\n" # 각 페이지 텍스트 추출
                self.status_bar.showMessage(f"텍스트 추출 중: 페이지 {i+1}/{total_pages}")
                QApplication.processEvents()


            with open(save_path, 'w', encoding='utf-8') as output_txt:
                output_txt.write(full_text)

            QMessageBox.information(self, "작업 완료", "텍스트 추출이 완료되었습니다.")
            self.status_bar.showMessage("텍스트 추출 완료.")
        except Exception as e:
            QMessageBox.critical(self, "오류 발생", f"텍스트 추출 중 오류가 발생했습니다: {e}")
            self.status_bar.showMessage("텍스트 추출 오류 발생.")
        finally:
            self.status_bar.showMessage("준비 완료")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFEditorApp()
    ex.show()
    sys.exit(app.exec())
