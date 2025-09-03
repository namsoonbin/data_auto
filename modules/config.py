# -*- coding: utf-8 -*-
import os
import sys
import logging

def get_exe_directory():
    """exe 파일이 위치한 디렉토리 반환 (개발 환경에서는 프로젝트 루트)"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller로 빌드된 경우: exe 파일이 있는 실제 디렉토리
        return os.path.dirname(sys.executable)
    else:
        # 개발 환경: 프로젝트 루트 디렉토리
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- 기본 경로 설정 ---
# 개발 환경에서는 프로젝트 루트, exe에서는 exe 파일 위치
BASE_DIR = get_exe_directory()

# --- 폴더 경로 --- 
DOWNLOAD_DIR = None # This will be set by desktop_app.py

def validate_directory(path):
    """디렉토리 경로 검증"""
    if not path or not isinstance(path, str):
        raise ValueError("경로는 비어있지 않은 문자열이어야 합니다.")
    
    # 경로 순회 공격 방지 (Windows와 Unix 모두 고려)
    normalized_path = os.path.normpath(path)
    if '..' in normalized_path:
        raise ValueError("안전하지 않은 경로입니다: 상위 디렉토리 참조 금지")
    
    # Windows의 절대 경로는 허용 (C:\, D:\ 등)
    if os.name == 'nt':  # Windows
        if not (normalized_path[1:3] == ':\\' and normalized_path[0].isalpha()):
            raise ValueError("Windows에서는 절대 경로(드라이브 문자 포함)만 허용됩니다.")
    else:  # Unix/Linux
        if normalized_path.startswith('/'):
            raise ValueError("Unix 시스템에서 루트 경로는 허용되지 않습니다.")
    
    return normalized_path

def get_processing_dir():
    if DOWNLOAD_DIR is None:
        raise ValueError("DOWNLOAD_DIR has not been set in config.")
    
    validated_path = validate_directory(DOWNLOAD_DIR)
    return os.path.join(validated_path, '작업폴더')

def get_archive_dir():
    if DOWNLOAD_DIR is None:
        raise ValueError("DOWNLOAD_DIR has not been set in config.")
    
    validated_path = validate_directory(DOWNLOAD_DIR)
    return os.path.join(validated_path, '원본_보관함')

def get_report_archive_dir():
    if DOWNLOAD_DIR is None:
        raise ValueError("DOWNLOAD_DIR has not been set in config.")
    
    validated_path = validate_directory(DOWNLOAD_DIR)
    return os.path.join(validated_path, '리포트보관함')

MARGIN_FILE = os.path.join(BASE_DIR, '마진정보.xlsx')

# --- 암호 설정 ---
# 주문조회 파일의 기본 암호
ORDER_FILE_PASSWORD = "1234"  # 기본 암호, 필요시 외부에서 변경 가능

# --- 리포트 설정 ---
COLUMNS_TO_KEEP = [
    '상품ID', '상품명', '옵션정보', '수량', '환불수량', '가구매 개수', '결제금액', '환불금액',
    '판매가', '마진율', '광고비율', '이윤율', '가구매 금액', '가구매 비용',
    '개당 가구매 금액', '개당 가구매 비용', '순매출', '매출', '판매마진', '순이익', '리워드'
]

# --- 데이터 처리 설정 ---
CANCEL_OR_REFUND_STATUSES = ['취소완료', '반품요청', '반품완료', '수거중']
