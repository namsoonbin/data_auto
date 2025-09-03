import sys
import os
import logging
import json
import pandas as pd
from datetime import datetime, date
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLineEdit, QTextEdit, QFileDialog, QLabel, QGroupBox, QGridLayout,
    QDialog, QTableWidget, QTableWidgetItem, QDateEdit, QHeaderView,
    QMessageBox, QSpinBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QDate

# --- Reward Manager Dialog ---
class RewardManagerDialog(QDialog):
    """리워드 관리 팝업창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('리워드 관리')
        self.setFixedSize(900, 700)
        self.setModal(True)
        
        # 데이터 저장 경로 (exe 파일과 같은 디렉토리)
        from modules import config
        self.reward_file = os.path.join(config.BASE_DIR, '리워드설정.json')
        self.margin_file = config.MARGIN_FILE
        
        self.initUI()
        self.load_products()
        self.load_existing_rewards()

    def initUI(self):
        layout = QVBoxLayout(self)
        
        # 제목
        title_label = QLabel("상품별 리워드 설정")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        # 날짜 설정 그룹
        date_group = QGroupBox("적용 날짜 범위")
        date_layout = QHBoxLayout()
        
        date_layout.addWidget(QLabel("시작일:"))
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        date_layout.addWidget(self.start_date)
        
        date_layout.addWidget(QLabel("종료일:"))
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate().addDays(6))  # 7일간
        self.end_date.setCalendarPopup(True)
        date_layout.addWidget(self.end_date)
        
        date_layout.addStretch()
        date_group.setLayout(date_layout)
        layout.addWidget(date_group)
        
        # 일괄 설정 그룹
        bulk_group = QGroupBox("일괄 설정")
        bulk_layout = QHBoxLayout()
        
        bulk_layout.addWidget(QLabel("리워드 금액:"))
        self.bulk_reward = QSpinBox()
        self.bulk_reward.setRange(0, 999999)
        self.bulk_reward.setSuffix(" 원")
        self.bulk_reward.setSingleStep(1000)
        self.bulk_reward.setValue(1000)
        bulk_layout.addWidget(self.bulk_reward)
        
        self.apply_all_button = QPushButton("전체 적용")
        self.apply_all_button.clicked.connect(self.apply_bulk_reward)
        bulk_layout.addWidget(self.apply_all_button)
        
        # 자주 사용하는 값 버튼들
        quick_buttons = [
            ("0원", 0), ("3000원", 3000), ("6000원", 6000), ("9000원", 9000)
        ]
        for text, value in quick_buttons:
            btn = QPushButton(text)
            btn.clicked.connect(lambda checked, v=value: self.bulk_reward.setValue(v))
            bulk_layout.addWidget(btn)
        
        bulk_layout.addStretch()
        bulk_group.setLayout(bulk_layout)
        layout.addWidget(bulk_group)
        
        # 검색 박스
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("검색:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("상품명으로 검색...")
        self.search_box.textChanged.connect(self.filter_products)
        search_layout.addWidget(self.search_box)
        layout.addLayout(search_layout)
        
        # 상품 테이블
        self.product_table = QTableWidget()
        self.product_table.setColumnCount(4)
        self.product_table.setHorizontalHeaderLabels(['상품ID', '상품명', '현재 리워드', '새 리워드'])
        
        # 테이블 컬럼 너비 설정
        header = self.product_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 상품ID
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 상품명
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # 현재 리워드
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # 새 리워드
        
        layout.addWidget(self.product_table)
        
        # 버튼들
        button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("저장")
        self.save_button.clicked.connect(self.save_rewards)
        self.save_button.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px 16px;")
        button_layout.addWidget(self.save_button)
        
        self.cancel_button = QPushButton("취소")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(button_layout)

    def load_products(self):
        """마진정보.xlsx에서 상품 목록 로드"""
        try:
            if not os.path.exists(self.margin_file):
                QMessageBox.warning(self, "경고", "마진정보.xlsx 파일을 찾을 수 없습니다.")
                return
            
            df = pd.read_excel(self.margin_file, engine='openpyxl')
            if '상품번호' in df.columns:
                df = df.rename(columns={'상품번호': '상품ID'})
            
            # 대표옵션만 표시 (리워드는 대표옵션에만 적용)
            if '대표옵션' in df.columns:
                df['대표옵션'] = df['대표옵션'].astype(str).str.upper().isin(['O', 'Y', 'TRUE'])
                df = df[df['대표옵션'] == True]
            
            self.products_df = df[['상품ID', '상품명']].drop_duplicates()
            self.populate_table()
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"상품 목록을 로드하는 중 오류가 발생했습니다:\n{e}")

    def populate_table(self):
        """테이블에 상품 목록 채우기"""
        self.product_table.setRowCount(len(self.products_df))
        
        for row, (_, product) in enumerate(self.products_df.iterrows()):
            # 상품ID
            self.product_table.setItem(row, 0, QTableWidgetItem(str(product['상품ID'])))
            
            # 상품명
            self.product_table.setItem(row, 1, QTableWidgetItem(str(product['상품명'])))
            
            # 현재 리워드 (처음엔 0)
            self.product_table.setItem(row, 2, QTableWidgetItem("0"))
            
            # 새 리워드 (편집 가능한 SpinBox)
            spinbox = QSpinBox()
            spinbox.setRange(0, 999999)
            spinbox.setSuffix(" 원")
            spinbox.setValue(0)
            self.product_table.setCellWidget(row, 3, spinbox)

    def load_existing_rewards(self):
        """기존 리워드 설정 로드"""
        if not os.path.exists(self.reward_file):
            return
        
        try:
            with open(self.reward_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            current_date = date.today()
            
            # 현재 날짜에 해당하는 리워드 설정 찾기
            for reward_entry in data.get('rewards', []):
                start_date = datetime.strptime(reward_entry['start_date'], '%Y-%m-%d').date()
                end_date = datetime.strptime(reward_entry['end_date'], '%Y-%m-%d').date()
                
                if start_date <= current_date <= end_date:
                    product_id = reward_entry['product_id']
                    reward_value = reward_entry['reward']
                    
                    # 테이블에서 해당 상품 찾아서 현재 리워드 업데이트
                    for row in range(self.product_table.rowCount()):
                        if self.product_table.item(row, 0).text() == str(product_id):
                            self.product_table.item(row, 2).setText(str(reward_value))
                            spinbox = self.product_table.cellWidget(row, 3)
                            if spinbox:
                                spinbox.setValue(reward_value)
                            break
                            
        except Exception as e:
            print(f"기존 리워드 로드 중 오류: {e}")

    def apply_bulk_reward(self):
        """일괄 리워드 적용"""
        bulk_value = self.bulk_reward.value()
        
        for row in range(self.product_table.rowCount()):
            if not self.product_table.isRowHidden(row):  # 필터링되지 않은 행만
                spinbox = self.product_table.cellWidget(row, 3)
                if spinbox:
                    spinbox.setValue(bulk_value)

    def filter_products(self):
        """상품명으로 필터링"""
        search_text = self.search_box.text().lower()
        
        for row in range(self.product_table.rowCount()):
            product_name = self.product_table.item(row, 1).text().lower()
            should_show = search_text in product_name
            self.product_table.setRowHidden(row, not should_show)

    def save_rewards(self):
        """리워드 설정 저장"""
        try:
            start_date_str = self.start_date.date().toString("yyyy-MM-dd")
            end_date_str = self.end_date.date().toString("yyyy-MM-dd")
            
            # 기존 데이터 로드
            reward_data = {'rewards': []}
            if os.path.exists(self.reward_file):
                with open(self.reward_file, 'r', encoding='utf-8') as f:
                    reward_data = json.load(f)
            
            # 새로운 설정들 추가
            for row in range(self.product_table.rowCount()):
                product_id = self.product_table.item(row, 0).text()
                spinbox = self.product_table.cellWidget(row, 3)
                if spinbox:
                    reward_value = spinbox.value()
                    
                    reward_entry = {
                        'start_date': start_date_str,
                        'end_date': end_date_str,
                        'product_id': product_id,
                        'reward': reward_value
                    }
                    reward_data['rewards'].append(reward_entry)
            
            # 파일 저장
            with open(self.reward_file, 'w', encoding='utf-8') as f:
                json.dump(reward_data, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(self, "완료", f"리워드 설정이 저장되었습니다.\n({start_date_str} ~ {end_date_str})")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"리워드 설정 저장 중 오류가 발생했습니다:\n{e}")

# --- Purchase Manager Dialog ---
class PurchaseManagerDialog(QDialog):
    """가구매 개수 관리 팝업창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('가구매 개수 관리')
        self.setFixedSize(900, 700)
        self.setModal(True)
        
        # 데이터 저장 경로 (exe 파일과 같은 디렉토리)
        from modules import config
        self.purchase_file = os.path.join(config.BASE_DIR, '가구매설정.json')
        self.margin_file = config.MARGIN_FILE
        
        self.initUI()
        self.load_products()
        self.load_existing_purchases()

    def initUI(self):
        layout = QVBoxLayout(self)
        
        # 제목
        title_label = QLabel("상품별 가구매 개수 설정")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        # 날짜 설정 그룹
        date_group = QGroupBox("적용 날짜 범위")
        date_layout = QHBoxLayout()
        
        date_layout.addWidget(QLabel("시작일:"))
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        date_layout.addWidget(self.start_date)
        
        date_layout.addWidget(QLabel("종료일:"))
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate().addDays(6))  # 7일간
        self.end_date.setCalendarPopup(True)
        date_layout.addWidget(self.end_date)
        
        date_layout.addStretch()
        date_group.setLayout(date_layout)
        layout.addWidget(date_group)
        
        # 일괄 설정 그룹
        bulk_group = QGroupBox("일괄 설정")
        bulk_layout = QHBoxLayout()
        
        bulk_layout.addWidget(QLabel("가구매 개수:"))
        self.bulk_purchase = QSpinBox()
        self.bulk_purchase.setRange(0, 9999)
        self.bulk_purchase.setSuffix(" 개")
        self.bulk_purchase.setSingleStep(1)
        self.bulk_purchase.setValue(0)
        bulk_layout.addWidget(self.bulk_purchase)
        
        self.apply_all_button = QPushButton("전체 적용")
        self.apply_all_button.clicked.connect(self.apply_bulk_purchase)
        bulk_layout.addWidget(self.apply_all_button)
        
        # 자주 사용하는 값 버튼들
        quick_buttons = [
            ("0개", 0), ("1개", 1), ("3개", 3), ("5개", 5), ("10개", 10)
        ]
        for text, value in quick_buttons:
            btn = QPushButton(text)
            btn.clicked.connect(lambda checked, v=value: self.bulk_purchase.setValue(v))
            bulk_layout.addWidget(btn)
        
        bulk_layout.addStretch()
        bulk_group.setLayout(bulk_layout)
        layout.addWidget(bulk_group)
        
        # 검색 박스
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("검색:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("상품명으로 검색...")
        self.search_box.textChanged.connect(self.filter_products)
        search_layout.addWidget(self.search_box)
        layout.addLayout(search_layout)
        
        # 상품 테이블
        self.product_table = QTableWidget()
        self.product_table.setColumnCount(4)
        self.product_table.setHorizontalHeaderLabels(['상품ID', '상품명', '현재 가구매', '새 가구매'])
        
        # 테이블 컬럼 너비 설정
        header = self.product_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 상품ID
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 상품명
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # 현재 가구매
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # 새 가구매
        
        layout.addWidget(self.product_table)
        
        # 버튼들
        button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("저장")
        self.save_button.clicked.connect(self.save_purchases)
        self.save_button.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px 16px;")
        button_layout.addWidget(self.save_button)
        
        self.cancel_button = QPushButton("취소")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(button_layout)

    def load_products(self):
        """마진정보.xlsx에서 상품 목록 로드"""
        try:
            if not os.path.exists(self.margin_file):
                QMessageBox.warning(self, "경고", "마진정보.xlsx 파일을 찾을 수 없습니다.")
                return
            
            df = pd.read_excel(self.margin_file, engine='openpyxl')
            if '상품번호' in df.columns:
                df = df.rename(columns={'상품번호': '상품ID'})
            
            # 대표옵션만 표시 (가구매는 대표옵션에만 적용)
            if '대표옵션' in df.columns:
                df['대표옵션'] = df['대표옵션'].astype(str).str.upper().isin(['O', 'Y', 'TRUE'])
                df = df[df['대표옵션'] == True]
            
            self.products_df = df[['상품ID', '상품명']].drop_duplicates()
            self.populate_table()
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"상품 목록을 로드하는 중 오류가 발생했습니다:\n{e}")

    def populate_table(self):
        """테이블에 상품 목록 채우기"""
        self.product_table.setRowCount(len(self.products_df))
        
        for row, (_, product) in enumerate(self.products_df.iterrows()):
            # 상품ID
            self.product_table.setItem(row, 0, QTableWidgetItem(str(product['상품ID'])))
            
            # 상품명
            self.product_table.setItem(row, 1, QTableWidgetItem(str(product['상품명'])))
            
            # 현재 가구매 개수 (처음엔 0)
            self.product_table.setItem(row, 2, QTableWidgetItem("0"))
            
            # 새 가구매 개수 (편집 가능한 SpinBox)
            spinbox = QSpinBox()
            spinbox.setRange(0, 9999)
            spinbox.setSuffix(" 개")
            spinbox.setValue(0)
            self.product_table.setCellWidget(row, 3, spinbox)

    def load_existing_purchases(self):
        """기존 가구매 설정 로드"""
        if not os.path.exists(self.purchase_file):
            return
        
        try:
            with open(self.purchase_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            current_date = date.today()
            
            # 현재 날짜에 해당하는 가구매 설정 찾기
            for purchase_entry in data.get('purchases', []):
                start_date = datetime.strptime(purchase_entry['start_date'], '%Y-%m-%d').date()
                end_date = datetime.strptime(purchase_entry['end_date'], '%Y-%m-%d').date()
                
                if start_date <= current_date <= end_date:
                    product_id = purchase_entry['product_id']
                    purchase_count = purchase_entry['purchase_count']
                    
                    # 테이블에서 해당 상품 찾아서 현재 가구매 개수 업데이트
                    for row in range(self.product_table.rowCount()):
                        if self.product_table.item(row, 0).text() == str(product_id):
                            self.product_table.item(row, 2).setText(str(purchase_count))
                            spinbox = self.product_table.cellWidget(row, 3)
                            if spinbox:
                                spinbox.setValue(purchase_count)
                            break
                            
        except Exception as e:
            print(f"기존 가구매 설정 로드 중 오류: {e}")

    def apply_bulk_purchase(self):
        """일괄 가구매 개수 적용"""
        bulk_value = self.bulk_purchase.value()
        
        for row in range(self.product_table.rowCount()):
            if not self.product_table.isRowHidden(row):  # 필터링되지 않은 행만
                spinbox = self.product_table.cellWidget(row, 3)
                if spinbox:
                    spinbox.setValue(bulk_value)

    def filter_products(self):
        """상품명으로 필터링"""
        search_text = self.search_box.text().lower()
        
        for row in range(self.product_table.rowCount()):
            product_name = self.product_table.item(row, 1).text().lower()
            should_show = search_text in product_name
            self.product_table.setRowHidden(row, not should_show)

    def save_purchases(self):
        """가구매 설정 저장"""
        try:
            start_date_str = self.start_date.date().toString("yyyy-MM-dd")
            end_date_str = self.end_date.date().toString("yyyy-MM-dd")
            
            # 기존 데이터 로드
            purchase_data = {'purchases': []}
            if os.path.exists(self.purchase_file):
                with open(self.purchase_file, 'r', encoding='utf-8') as f:
                    purchase_data = json.load(f)
            
            # 새로운 설정들 추가
            for row in range(self.product_table.rowCount()):
                product_id = self.product_table.item(row, 0).text()
                spinbox = self.product_table.cellWidget(row, 3)
                if spinbox:
                    purchase_count = spinbox.value()
                    
                    purchase_entry = {
                        'start_date': start_date_str,
                        'end_date': end_date_str,
                        'product_id': product_id,
                        'purchase_count': purchase_count
                    }
                    purchase_data['purchases'].append(purchase_entry)
            
            # 파일 저장
            with open(self.purchase_file, 'w', encoding='utf-8') as f:
                json.dump(purchase_data, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(self, "완료", f"가구매 개수 설정이 저장되었습니다.\n({start_date_str} ~ {end_date_str})")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"가구매 설정 저장 중 오류가 발생했습니다:\n{e}")


# --- Custom Logging Handler ---
class PyQtSignalHandler(logging.Handler):
    """A logging handler that emits a PyQt signal."""
    def __init__(self, signal):
        super().__init__()
        self.signal = signal

    def emit(self, record):
        msg = self.format(record)
        self.signal.emit(msg)

# --- Worker Thread ---
class Worker(QThread):
    """
    Runs the file monitoring and processing logic in a separate thread.
    """
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, download_folder_path, password=None):
        super().__init__()
        self.download_folder_path = download_folder_path
        self.password = password
        self.handler = None

    def run(self):
        """
        Configures logging for this thread, sets the download directory,
        and starts the file monitoring process.
        """
        # Configure logging to emit signals
        self.handler = PyQtSignalHandler(self.output_signal)
        self.handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        logging.getLogger().addHandler(self.handler)
        logging.getLogger().setLevel(logging.INFO)

        try:
            # Dynamically import and set config
            from modules import config
            config.DOWNLOAD_DIR = self.download_folder_path
            
            # Set password if provided
            if hasattr(self, 'password') and self.password:
                config.ORDER_FILE_PASSWORD = self.password
                logging.info(f"주문조회 파일 암호가 설정되었습니다.")
            
            # Dynamically import file_handler and start monitoring
            from modules import file_handler
            file_handler.start_monitoring()

        except Exception as e:
            logging.error(f"자동화 프로세스 실행 중 오류 발생: {e}")
        finally:
            # 안전한 정리 작업
            try:
                if self.handler:
                    logging.getLogger().removeHandler(self.handler)
            except:
                pass  # 로깅 핸들러 제거 실패해도 계속 진행
            
            self.finished_signal.emit()

# --- Manual Process Worker Thread ---
class ManualProcessWorker(QThread):
    """작업폴더의 미완료 파일들을 수동으로 처리하는 워커 스레드"""
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, download_folder_path, password):
        super().__init__()
        self.download_folder_path = download_folder_path
        self.password = password
        self.handler = None

    def run(self):
        # Configure logging to emit signals
        self.handler = PyQtSignalHandler(self.output_signal)
        self.handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        logging.getLogger().addHandler(self.handler)
        logging.getLogger().setLevel(logging.INFO)

        try:
            # Dynamically import and set config
            from modules import config, file_handler
            config.DOWNLOAD_DIR = self.download_folder_path
            
            if self.password:
                config.ORDER_FILE_PASSWORD = self.password
            
            # 작업폴더 초기화
            file_handler.initialize_folders()
            
            # 미완료 파일들 처리
            file_handler.process_incomplete_files()
            
            # 최종 정리 수행 (전체 통합 리포트 생성 및 파일 이동)
            file_handler.finalize_all_processing()
            
        except Exception as e:
            logging.error(f"수동 처리 중 오류 발생: {e}")
        finally:
            try:
                if self.handler:
                    logging.getLogger().removeHandler(self.handler)
            except:
                pass
            
            self.finished_signal.emit()

# --- Main Application UI ---
class DesktopApp(QWidget):
    def __init__(self):
        super().__init__()
        self.is_monitoring = False
        self.is_manual_processing = False  # 수동 처리 상태 추가
        self.worker = None
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.stop_flag_path = os.path.join(self.base_dir, 'stop.flag')
        self.download_folder_path = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle('판매 데이터 자동화')
        self.setGeometry(100, 100, 800, 600)
        
        # Stylesheet
        self.setStyleSheet("""
            QWidget { background-color: #f0f2f5; font-family: '맑은 고딕'; }
            QLabel { font-size: 14px; color: #333; }
            QLineEdit { background-color: #fff; border: 1px solid #ccc; padding: 8px; border-radius: 4px; font-size: 14px; }
            QTextEdit { background-color: #fff; border: 1px solid #ccc; border-radius: 4px; color: #333; font-size: 13px; }
            QPushButton { background-color: #007bff; color: white; font-size: 15px; font-weight: bold; padding: 10px 15px; border-radius: 5px; border: none; }
            QPushButton:hover { background-color: #0056b3; }
            QPushButton:disabled { background-color: #999; }
            #stopButton { background-color: #dc3545; }
            #stopButton:hover { background-color: #c82333; }
            QGroupBox { font-size: 16px; font-weight: bold; margin-top: 10px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px 0 5px; }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Folder Selection
        folder_layout = QHBoxLayout()
        self.folder_path_input = QLineEdit()
        self.folder_path_input.setReadOnly(True)
        self.browse_button = QPushButton("폴더 선택")
        self.browse_button.clicked.connect(self.browse_folder)
        folder_layout.addWidget(QLabel("다운로드 폴더:"))
        folder_layout.addWidget(self.folder_path_input)
        folder_layout.addWidget(self.browse_button)
        main_layout.addLayout(folder_layout)

        # 설정 그룹
        settings_group = QGroupBox("설정")
        settings_layout = QGridLayout()
        
        # 암호 설정
        password_label = QLabel("주문조회 파일 암호:")
        self.password_input = QLineEdit()
        self.password_input.setText("1234")  # 기본값
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("암호를 입력하세요 (기본: 1234)")
        
        # 암호 표시/숨기기 버튼
        self.show_password_button = QPushButton("표시")
        self.show_password_button.setMaximumWidth(60)
        self.show_password_button.clicked.connect(self.toggle_password_visibility)
        
        password_layout = QHBoxLayout()
        password_layout.addWidget(self.password_input)
        password_layout.addWidget(self.show_password_button)
        
        settings_layout.addWidget(password_label, 0, 0)
        settings_layout.addLayout(password_layout, 0, 1)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)

        # Control Buttons
        button_layout = QHBoxLayout()
        
        self.toggle_button = QPushButton("자동화 시작")
        self.toggle_button.clicked.connect(self.toggle_monitoring)
        button_layout.addWidget(self.toggle_button)
        
        self.manual_process_button = QPushButton("작업폴더 처리")
        self.manual_process_button.clicked.connect(self.manual_process)
        self.manual_process_button.setStyleSheet("background-color: #28a745;")
        self.manual_process_button.setToolTip("작업폴더의 미완료 파일들을 수동으로 처리합니다")
        button_layout.addWidget(self.manual_process_button)
        
        self.reward_button = QPushButton("리워드 설정")
        self.reward_button.clicked.connect(self.open_reward_manager)
        self.reward_button.setStyleSheet("background-color: #ffc107; color: black;")
        self.reward_button.setToolTip("상품별 리워드를 설정합니다")
        button_layout.addWidget(self.reward_button)
        
        self.purchase_button = QPushButton("가구매 설정")
        self.purchase_button.clicked.connect(self.open_purchase_manager)
        self.purchase_button.setStyleSheet("background-color: #17a2b8; color: white;")
        self.purchase_button.setToolTip("상품별 가구매 개수를 설정합니다")
        button_layout.addWidget(self.purchase_button)
        
        main_layout.addLayout(button_layout)

        # 상태 정보 표시
        status_group = QGroupBox("상태")
        status_layout = QGridLayout()
        
        self.status_label = QLabel("대기 중")
        self.status_label.setStyleSheet("color: #666; font-size: 16px; font-weight: bold;")
        
        status_layout.addWidget(QLabel("현재 상태:"), 0, 0)
        status_layout.addWidget(self.status_label, 0, 1)
        
        status_group.setLayout(status_layout)
        main_layout.addWidget(status_group)

        # Log Display
        main_layout.addWidget(QLabel("실행 로그"))
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        main_layout.addWidget(self.log_output)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "다운로드 폴더 선택")
        if folder:
            self.download_folder_path = folder
            self.folder_path_input.setText(folder)
            self.update_log(f"[INFO] 다운로드 폴더 설정: {folder}")

    def toggle_monitoring(self):
        if self.is_monitoring:
            self.stop_monitoring()
        else:
            self.start_monitoring()

    def start_monitoring(self):
        if not self.download_folder_path:
            self.update_log("[ERROR] 다운로드 폴더를 먼저 선택해주세요.")
            return

        self.log_output.clear()
        
        if os.path.exists(self.stop_flag_path):
            os.remove(self.stop_flag_path)

        self.is_monitoring = True
        self.toggle_button.setText("자동화 중지")
        self.toggle_button.setObjectName("stopButton")
        self.setStyleSheet(self.styleSheet()) # Refresh stylesheet for ID
        self.browse_button.setEnabled(False)
        self.password_input.setEnabled(False)
        
        # 상태 업데이트
        self.status_label.setText("실행 중")
        self.status_label.setStyleSheet("color: #28a745; font-size: 16px; font-weight: bold;")

        # 암호 값 가져오기
        password = self.password_input.text().strip() if self.password_input.text().strip() else "1234"
        
        self.worker = Worker(self.download_folder_path, password)
        self.worker.output_signal.connect(self.update_log)
        self.worker.finished_signal.connect(self.on_monitoring_finished)
        self.worker.start()

    def stop_monitoring(self):
        if not self.is_monitoring:
            return
        self.update_log("[INFO] 자동화 중지를 요청합니다...")
        
        # Worker 스레드에 중지 신호 전송
        try:
            with open(self.stop_flag_path, 'w') as f:
                f.write('stop')
        except Exception as e:
            self.update_log(f"[ERROR] 중지 신호 전송 실패: {e}")
        
        self.toggle_button.setEnabled(False)
        
        # 상태 업데이트
        self.status_label.setText("중지 중")
        self.status_label.setStyleSheet("color: #ffc107; font-size: 16px; font-weight: bold;")
        
        # 타임아웃과 함께 Worker 종료 대기 (비동기)
        if self.worker and self.worker.isRunning():
            self.worker.terminate()  # 강제 종료 시도
            
            # 1초 후 강제 정리 (타이머 사용으로 UI 블로킹 방지)
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, self.force_cleanup)

    def force_cleanup(self):
        """타임아웃 후 강제로 정리 작업 수행"""
        self.update_log("[INFO] 강제 정리 작업을 수행합니다...")
        if self.worker and self.worker.isRunning():
            self.worker.kill()  # 완전 강제 종료
        self.on_monitoring_finished()

    def update_log(self, text):
        self.log_output.append(text)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())


    def on_monitoring_finished(self):
        self.update_log("[INFO] 자동화 프로세스가 종료되었습니다.")
        self.is_monitoring = False
        
        # Worker 정리
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
            
        self.toggle_button.setText("자동화 시작")
        self.toggle_button.setObjectName("")
        self.setStyleSheet(self.styleSheet()) # Refresh stylesheet
        self.toggle_button.setEnabled(True)
        self.browse_button.setEnabled(True)
        self.password_input.setEnabled(True)
        
        # 상태 업데이트
        self.status_label.setText("대기 중")
        self.status_label.setStyleSheet("color: #666; font-size: 16px; font-weight: bold;")
        
        if os.path.exists(self.stop_flag_path):
            os.remove(self.stop_flag_path)
    
    def toggle_password_visibility(self):
        """암호 표시/숨기기 토글"""
        if self.password_input.echoMode() == QLineEdit.Password:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.show_password_button.setText("숨기기")
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.show_password_button.setText("표시")

    def manual_process(self):
        """작업폴더의 미완료 파일들을 수동으로 처리 또는 중지"""
        if self.is_manual_processing:
            # 수동 처리 중지
            self.stop_manual_process()
            return
        
        if not self.download_folder_path:
            self.update_log("[ERROR] 다운로드 폴더를 먼저 선택해주세요.")
            return
        
        if self.is_monitoring:
            self.update_log("[WARNING] 자동화 실행 중에는 수동 처리를 할 수 없습니다.")
            return
        
        # 수동 처리 시작
        self.start_manual_process()
    
    def start_manual_process(self):
        """수동 처리 시작"""
        self.log_output.clear()
        
        if os.path.exists(self.stop_flag_path):
            os.remove(self.stop_flag_path)
        
        self.is_manual_processing = True
        self.manual_process_button.setText("처리 중지")
        self.manual_process_button.setStyleSheet("background-color: #dc3545; color: white;")  # 빨간색
        self.toggle_button.setEnabled(False)  # 자동화 버튼 비활성화
        self.reward_button.setEnabled(False)  # 리워드 버튼 비활성화
        
        # 상태 업데이트
        self.status_label.setText("수동 처리 중")
        self.status_label.setStyleSheet("color: #ffc107; font-size: 16px; font-weight: bold;")
        
        self.update_log("[INFO] 작업폴더의 미완료 파일들을 수동 처리합니다...")
        
        # Worker 스레드로 수동 처리 실행
        self.manual_worker = ManualProcessWorker(self.download_folder_path, self.password_input.text().strip() or "1234")
        self.manual_worker.output_signal.connect(self.update_log)
        self.manual_worker.finished_signal.connect(self.on_manual_process_finished)
        self.manual_worker.start()
    
    def stop_manual_process(self):
        """수동 처리 중지"""
        if not self.is_manual_processing:
            return
        
        self.update_log("[INFO] 수동 처리 중지를 요청합니다...")
        
        # Worker 스레드에 중지 신호 전송
        try:
            with open(self.stop_flag_path, 'w') as f:
                f.write('stop')
        except Exception as e:
            self.update_log(f"[ERROR] 중지 신호 전송 실패: {e}")
        
        self.manual_process_button.setEnabled(False)
        self.manual_process_button.setText("중지 중...")
        
        # 상태 업데이트
        self.status_label.setText("중지 중")
        self.status_label.setStyleSheet("color: #ffc107; font-size: 16px; font-weight: bold;")
        
        # 타임아웃과 함께 Worker 종료 대기
        if hasattr(self, 'manual_worker') and self.manual_worker and self.manual_worker.isRunning():
            self.manual_worker.terminate()
            
            # 1초 후 강제 정리
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, self.force_manual_cleanup)

    def force_manual_cleanup(self):
        """수동 처리 강제 정리"""
        self.update_log("[INFO] 수동 처리를 강제로 정리합니다...")
        if hasattr(self, 'manual_worker') and self.manual_worker and self.manual_worker.isRunning():
            self.manual_worker.kill()
        self.on_manual_process_finished()
    
    def on_manual_process_finished(self):
        """수동 처리 완료 시 호출"""
        # ManualProcessWorker 정리
        if hasattr(self, 'manual_worker') and self.manual_worker:
            self.manual_worker.deleteLater()
            self.manual_worker = None
        
        # 상태 초기화
        self.is_manual_processing = False
        
        # 버튼 및 UI 복원
        self.manual_process_button.setEnabled(True)
        self.manual_process_button.setText("작업폴더 처리")
        self.manual_process_button.setStyleSheet("background-color: #28a745;")  # 원래 초록색
        self.toggle_button.setEnabled(True)  # 자동화 버튼 활성화
        self.reward_button.setEnabled(True)  # 리워드 버튼 활성화
        
        # 상태 업데이트
        self.status_label.setText("대기 중")
        self.status_label.setStyleSheet("color: #666; font-size: 16px; font-weight: bold;")
        
        # stop.flag 파일 정리
        if os.path.exists(self.stop_flag_path):
            os.remove(self.stop_flag_path)
        
        self.update_log("[INFO] 수동 처리가 완료되었습니다.")

    def open_reward_manager(self):
        """리워드 관리 팝업창 열기"""
        if self.is_monitoring:
            self.update_log("[WARNING] 자동화 실행 중에는 리워드 설정을 할 수 없습니다.")
            return
        
        if self.is_manual_processing:
            self.update_log("[WARNING] 수동 처리 중에는 리워드 설정을 할 수 없습니다.")
            return
        
        try:
            dialog = RewardManagerDialog(self)
            result = dialog.exec_()
            
            if result == QDialog.Accepted:
                self.update_log("[INFO] 리워드 설정이 저장되었습니다.")
            
        except Exception as e:
            self.update_log(f"[ERROR] 리워드 관리 창을 여는 중 오류 발생: {e}")
            QMessageBox.critical(self, "오류", f"리워드 관리 창을 여는 중 오류가 발생했습니다:\n{e}")

    def open_purchase_manager(self):
        """가구매 관리 팝업창 열기"""
        if self.is_monitoring:
            self.update_log("[WARNING] 자동화 실행 중에는 가구매 설정을 할 수 없습니다.")
            return
        
        if self.is_manual_processing:
            self.update_log("[WARNING] 수동 처리 중에는 가구매 설정을 할 수 없습니다.")
            return
        
        try:
            dialog = PurchaseManagerDialog(self)
            result = dialog.exec_()
            
            if result == QDialog.Accepted:
                self.update_log("[INFO] 가구매 설정이 저장되었습니다.")
            
        except Exception as e:
            self.update_log(f"[ERROR] 가구매 관리 창을 여는 중 오류 발생: {e}")
            QMessageBox.critical(self, "오류", f"가구매 관리 창을 여는 중 오류가 발생했습니다:\n{e}")

    def closeEvent(self, event):
        if self.is_monitoring or self.is_manual_processing:
            self.update_log("[INFO] 프로그램 종료 중...")
            
            # 자동화 중지
            if self.is_monitoring:
                self.stop_monitoring()
                if self.worker and self.worker.isRunning():
                    if not self.worker.wait(2000):  # 2초 타임아웃
                        self.worker.terminate()
                        if not self.worker.wait(1000):  # 1초 추가 대기
                            self.worker.kill()  # 완전 강제 종료
            
            # 수동 처리 중지
            if self.is_manual_processing:
                self.stop_manual_process()
                if hasattr(self, 'manual_worker') and self.manual_worker and self.manual_worker.isRunning():
                    if not self.manual_worker.wait(2000):  # 2초 타임아웃
                        self.manual_worker.terminate()
                        if not self.manual_worker.wait(1000):  # 1초 추가 대기
                            self.manual_worker.kill()  # 완전 강제 종료
        
        # stop.flag 파일 정리
        if os.path.exists(self.stop_flag_path):
            try:
                os.remove(self.stop_flag_path)
            except:
                pass  # 파일 삭제 실패해도 프로그램 종료 진행
                
        event.accept()

if __name__ == '__main__':
    # This is important for multiprocessing support in frozen apps
    import multiprocessing
    multiprocessing.freeze_support()
    
    app = QApplication(sys.argv)
    ex = DesktopApp()
    ex.show()
    sys.exit(app.exec_())