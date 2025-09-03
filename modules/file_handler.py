# -*- coding: utf-8 -*-
import os
import re
import shutil
import time
import logging
import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from . import config
from . import report_generator

STOP_FLAG_FILE = os.path.join(config.BASE_DIR, 'stop.flag')

def validate_excel_file(file_path):
    """Excel 파일 검증 (암호 보호된 파일 포함)"""
    if not file_path.lower().endswith('.xlsx'):
        raise ValueError(f"지원하지 않는 파일 형식입니다: {file_path}")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    
    # 파일 크기 체크 (100MB 제한)
    file_size = os.path.getsize(file_path)
    if file_size > 100 * 1024 * 1024:
        raise ValueError(f"파일 크기가 너무 큽니다 (100MB 초과): {file_path}")
    
    # 암호 보호된 파일인지 확인 (파일 헤더 체크)
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            # Microsoft Office 암호화된 파일의 시그니처
            if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
                logging.info(f"암호 보호된 파일 감지: {os.path.basename(file_path)}")
    except Exception:
        pass  # 헤더 읽기 실패해도 계속 진행
    
    return True

def get_file_info(src_path):
    """파일 경로를 분석하여 스토어, 날짜, 파일 타입, 새 파일명을 반환합니다."""
    try:
        validate_excel_file(src_path)
        path_parts = src_path.split(os.sep)
        if len(path_parts) > 1 and path_parts[-2] != os.path.basename(config.DOWNLOAD_DIR):
            store_name = path_parts[-2]
        else:
            return None, None, None, None

        original_filename = os.path.basename(src_path)
        date_str, file_type, new_filename = None, None, None

        order_match = re.match(r"스마트스토어_주문조회_(\d{4}-\d{2}-\d{2})\.xlsx", original_filename)
        if order_match:
            date_str = order_match.group(1)
            file_type = '주문'
            new_filename = f"{store_name} 스마트스토어_주문조회_{date_str}.xlsx"
        
        perf_match = re.match(r"상품성과_(\d{4}-\d{2}-\d{2}).*?\.xlsx", original_filename)
        if perf_match:
            date_str = perf_match.group(1)
            file_type = '성과'
            new_filename = f"{store_name} 상품성과_{date_str}.xlsx"

        if date_str and file_type and new_filename:
            return store_name, date_str, file_type, new_filename
        return None, None, None, None
    except ValueError as e:
        logging.warning(f"[get_file_info] 파일 검증 실패: {e}")
        return None, None, None, None
    except FileNotFoundError as e:
        logging.error(f"[get_file_info] 파일 없음: {e}")
        return None, None, None, None
    except Exception as e:
        logging.error(f"[get_file_info] 정보 추출 중 예상치 못한 오류: {e}")
        return None, None, None, None

def _check_and_process_data(store, date):
    """파일 쌍이 준비되었는지 확인하고 리포트 생성을 트리거합니다."""
    logging.info(f"[{store}, {date}] 파일 쌍 확인 및 데이터 처리 시작...")
    perf_file = f"{store} 상품성과_{date}.xlsx"
    order_file = f"{store} 스마트스토어_주문조회_{date}.xlsx"
    perf_path = os.path.join(config.get_processing_dir(), perf_file)
    order_path = os.path.join(config.get_processing_dir(), order_file)

    if os.path.exists(perf_path) and os.path.exists(order_path):
        logging.info(f"[{store}, {date}] 파일 쌍 발견! 데이터 처리를 시작합니다.")
        
        # 이미 리포트가 생성되어 있는지 확인
        individual_report = f'{store}_통합_리포트_{date}.xlsx'
        individual_report_path = os.path.join(config.get_processing_dir(), individual_report)
        
        if os.path.exists(individual_report_path):
            logging.info(f"[{store}, {date}] 이미 리포트가 생성되어 있습니다.")
        else:
            # 개별 리포트만 생성 (파일 이동은 하지 않음)
            processed_groups = report_generator.generate_individual_reports()
            if not processed_groups:
                logging.error(f"[{store}, {date}] 리포트 생성에 실패했습니다.")
                return
        
        logging.info(f"[{store}, {date}] 개별 리포트 처리 완료.")
    else:
        logging.info(f"[{store}, {date}] 아직 파일 쌍이 준비되지 않았습니다.")

def process_file(src_path):
    """감지된 파일을 처리 폴더로 옮기고, 데이터 처리를 시작합니다."""
    logging.info(f"[process_file] 파일 처리 시작: {src_path}")
    store, date, file_type, new_filename = get_file_info(src_path)
    if not all([store, date, file_type, new_filename]):
        logging.warning(f"[process_file] 파일 정보가 올바르지 않아 무시합니다: {src_path}")
        return

    dest_path = os.path.join(config.get_processing_dir(), new_filename)
    try:
        logging.info(f"[process_file] 파일 이동: '{src_path}' -> '{dest_path}'")
        shutil.move(src_path, dest_path)
        logging.info("[process_file] 파일 이동 완료.")
        _check_and_process_data(store, date)
    except Exception as e:
        logging.error(f"[process_file] 파일 이동/처리 중 오류: {e}")

class FileProcessorHandler(FileSystemEventHandler):
    """파일 시스템 이벤트를 감지하여 파일 처리를 시작하는 핸들러"""
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith('.xlsx') and not os.path.basename(event.src_path).startswith('~$'):
            logging.info(f"[on_created] 새 파일 감지: {event.src_path}")
            time.sleep(1)
            process_file(event.src_path)

def process_existing_files():
    """프로그램 시작 시 다운로드 폴더에 이미 있는 파일들을 처리합니다."""
    logging.info("===== 기존 파일 스캔 시작 ====")
    
    if not os.path.exists(config.DOWNLOAD_DIR):
        logging.warning(f"감시할 다운로드 폴더가 존재하지 않습니다: {config.DOWNLOAD_DIR}")
        return
    
    for store_folder in os.listdir(config.DOWNLOAD_DIR):
        store_path = os.path.join(config.DOWNLOAD_DIR, store_folder)
        if os.path.isdir(store_path):
            for filename in os.listdir(store_path):
                if filename.endswith('.xlsx') and not filename.startswith('~'):
                    file_path = os.path.join(store_path, filename)
                    logging.info(f"[기존 파일] 처리 시도: '{file_path}'")
                    # 개별 파일 처리 (최종 정리는 나중에 일괄 수행)
                    process_file(file_path)
    
    # 2단계: 작업폴더의 미완료 처리 파일들 검사 및 처리
    process_incomplete_files()
    
    # 3단계: 모든 파일 처리 완료 후 최종 정리 수행 (한 번만)
    finalize_all_processing()
    
    logging.info("===== 기존 파일 스캔 완료 ====")

def process_incomplete_files():
    """작업폴더에 있는 미완료 처리 파일들을 검사하고 리포트 생성을 시도합니다."""
    logging.info("--- 작업폴더 미완료 파일 검사 시작 ---")
    
    # 중지 신호 확인
    if os.path.exists(STOP_FLAG_FILE):
        logging.info("중지 신호 감지. 작업폴더 처리를 중단합니다.")
        return
    
    if not os.path.exists(config.get_processing_dir()):
        return
    
    # 작업폴더의 모든 엑셀 파일 스캔
    all_files = [f for f in os.listdir(config.get_processing_dir()) if f.endswith('.xlsx') and not f.startswith('~')]
    source_files = [f for f in all_files if '통합_리포트' not in f and '마진정보' not in f]
    
    if not source_files:
        logging.info("작업폴더에 미처리 파일이 없습니다.")
        return
    
    # 스토어별, 날짜별 파일 그룹 생성
    file_groups = {}
    for f in source_files:
        store, date, file_type = None, None, None
        if '상품성과' in f:
            parts = f.split(' 상품성과_')
            if len(parts) == 2: 
                store, date, file_type = parts[0], parts[1].replace('.xlsx',''), '성과'
        elif '스마트스토어_주문조회' in f:
            parts = f.split(' 스마트스토어_주문조회_')
            if len(parts) == 2: 
                store, date, file_type = parts[0], parts[1].replace('.xlsx',''), '주문'
        
        if store and date and file_type:
            key = (store, date)
            if key not in file_groups: 
                file_groups[key] = {}
            file_groups[key][file_type] = f
    
    # 완전한 파일 쌍이 있는데 리포트가 없는 경우 처리
    processed_any = False
    for (store, date), files in file_groups.items():
        # 중지 신호 확인
        if os.path.exists(STOP_FLAG_FILE):
            logging.info("중지 신호 감지. 미완료 파일 처리를 중단합니다.")
            return
            
        if '성과' in files and '주문' in files:
            individual_report = f'{store}_통합_리포트_{date}.xlsx'
            individual_report_path = os.path.join(config.get_processing_dir(), individual_report)
            
            if not os.path.exists(individual_report_path):
                logging.info(f"[미완료 처리 발견] {store} ({date}) - 리포트 생성을 재시도합니다.")
                _check_and_process_data(store, date)
    
    logging.info("--- 작업폴더 미완료 파일 검사 완료 ---")

def finalize_all_processing():
    """모든 개별 처리 완료 후 전체 통합 리포트 생성 및 파일 정리를 일괄 수행합니다."""
    # 중지 신호 확인
    if os.path.exists(STOP_FLAG_FILE):
        logging.info("중지 신호 감지. 최종 정리 작업을 중단합니다.")
        return
    
    processing_dir = config.get_processing_dir()
    
    # 처리할 것이 있는지 확인
    if not os.path.exists(processing_dir):
        return
    
    # 원본 파일이나 개별 리포트가 있는지 확인
    source_files = [f for f in os.listdir(processing_dir) 
                   if f.endswith('.xlsx') and '통합_리포트' not in f and not f.startswith('~')]
    report_files = [f for f in os.listdir(processing_dir) 
                   if f.endswith('.xlsx') and '통합_리포트' in f and not f.startswith('~')]
    
    if not source_files and not report_files:
        logging.info("정리할 파일이 없습니다.")
        return
    
    logging.info("=== 최종 정리 작업 시작 ===")
    
    # 1단계: 전체 통합 리포트 생성 (개별 리포트가 있는 경우에만)
    if report_files:
        logging.info("1단계: 전체 통합 리포트 생성 중...")
        report_generator.consolidate_daily_reports()
    
    # 2단계: 모든 원본 파일들을 원본_보관함으로 이동
    if source_files:
        logging.info("2단계: 원본 파일들을 원본_보관함으로 이동 중...")
        move_source_files_to_archive()
    
    # 3단계: 모든 리포트 파일들을 리포트보관함으로 이동
    logging.info("3단계: 리포트 파일들을 리포트보관함으로 이동 중...")
    move_reports_to_archive()
    
    logging.info("=== 최종 정리 작업 완료 ===")

def move_source_files_to_archive():
    """작업폴더의 모든 원본 파일들(상품성과, 주문조회)을 원본_보관함으로 이동합니다."""
    processing_dir = config.get_processing_dir()
    archive_dir = config.get_archive_dir()
    
    if not os.path.exists(processing_dir):
        return
    
    # 원본 파일들 찾기 (통합_리포트가 아닌 파일들)
    source_files = [f for f in os.listdir(processing_dir) 
                   if f.endswith('.xlsx') and '통합_리포트' not in f and not f.startswith('~')]
    
    if not source_files:
        logging.info("이동할 원본 파일이 없습니다.")
        return
    
    logging.info(f"--- 원본 파일들을 원본_보관함으로 이동 시작 ({len(source_files)}개 파일) ---")
    
    for source_file in source_files:
        try:
            src_path = os.path.join(processing_dir, source_file)
            dst_path = os.path.join(archive_dir, source_file)
            shutil.move(src_path, dst_path)
            logging.info(f"원본 파일 이동 완료: {source_file}")
        except Exception as e:
            logging.error(f"원본 파일 이동 실패 ({source_file}): {e}")
    
    logging.info("--- 원본 파일 이동 완료 ---")

def move_reports_to_archive():
    """작업폴더의 리포트 파일들을 리포트보관함으로 이동합니다."""
    processing_dir = config.get_processing_dir()
    report_archive_dir = config.get_report_archive_dir()
    
    if not os.path.exists(processing_dir):
        return
    
    # 리포트 파일들 찾기 (통합_리포트로 시작하는 파일들)
    report_files = [f for f in os.listdir(processing_dir) 
                   if f.endswith('.xlsx') and '통합_리포트' in f and not f.startswith('~')]
    
    if not report_files:
        return
    
    logging.info(f"--- 리포트 파일들을 리포트보관함으로 이동 시작 ({len(report_files)}개 파일) ---")
    
    for report_file in report_files:
        try:
            src_path = os.path.join(processing_dir, report_file)
            dst_path = os.path.join(report_archive_dir, report_file)
            
            # 이미 같은 이름의 파일이 존재하는 경우에만 백업 (과도한 백업 방지)
            if os.path.exists(dst_path):
                # 파일 크기나 수정 시간이 다른 경우에만 백업
                src_stat = os.path.getsize(src_path)
                dst_stat = os.path.getsize(dst_path)
                
                if src_stat != dst_stat:  # 크기가 다르면 새로운 데이터
                    timestamp = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S")
                    name, ext = os.path.splitext(report_file)
                    backup_name = f"{name}_backup{timestamp}{ext}"
                    backup_path = os.path.join(report_archive_dir, backup_name)
                    shutil.move(dst_path, backup_path)
                    logging.info(f"기존 리포트 백업: {backup_name}")
                else:
                    # 같은 크기면 덮어쓰기 (백업하지 않음)
                    os.remove(dst_path)
                    logging.info(f"동일한 리포트 덮어쓰기: {report_file}")
            
            shutil.move(src_path, dst_path)
            logging.info(f"리포트 이동 완료: {report_file}")
        except Exception as e:
            logging.error(f"리포트 이동 실패 ({report_file}): {e}")
    
    logging.info("--- 리포트 파일 이동 완료 ---")

def initialize_folders():
    """필요한 모든 폴더가 존재하는지 확인하고 없으면 생성합니다."""
    if not os.path.exists(config.get_processing_dir()): os.makedirs(config.get_processing_dir())
    if not os.path.exists(config.get_archive_dir()): os.makedirs(config.get_archive_dir())
    if not os.path.exists(config.get_report_archive_dir()): os.makedirs(config.get_report_archive_dir())

def start_monitoring():
    """파일 시스템 모니터링을 시작하고, stop.flag 파일이 생기면 중지합니다."""
    initialize_folders()
    # 시작 시 혹시 남아있을지 모르는 플래그 파일 삭제
    if os.path.exists(STOP_FLAG_FILE):
        os.remove(STOP_FLAG_FILE)

    process_existing_files()
    
    logging.info("\n===== 스마트 폴더 실시간 모니터링 시작 =====")
    logging.info(f"- 감시 대상: {config.DOWNLOAD_DIR} (하위 폴더 포함)")
    logging.info("- 파일을 각 스토어 폴더에 넣으면 처리가 시작됩니다.")
    
    event_handler = FileProcessorHandler()
    observer = Observer()
    observer.schedule(event_handler, config.DOWNLOAD_DIR, recursive=True)
    observer.start()

    try:
        while True:
            if os.path.exists(STOP_FLAG_FILE):
                logging.info("'stop.flag' 파일 감지. 모니터링을 중지합니다.")
                break
            time.sleep(1)
    finally:
        observer.stop()
        observer.join() # 스레드가 완전히 종료될 때까지 대기
        if os.path.exists(STOP_FLAG_FILE):
            os.remove(STOP_FLAG_FILE)
        logging.info("\n===== 모니터링이 정상적으로 종료되었습니다. =====")
