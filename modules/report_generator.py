# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import glob
import re
import logging
import io
import json
from datetime import datetime
from . import config

def normalize_product_id(value):
    """상품ID를 정규화 - 문자열과 숫자 타입 모두 처리"""
    if pd.isna(value):
        return ''
    
    # 이미 문자열인 경우
    if isinstance(value, str):
        return value.strip()
    
    # 숫자 타입인 경우 (int, float)
    if isinstance(value, (int, float)):
        # float인 경우 .0 제거를 위해 int로 변환
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        else:
            return str(value)
    
    # 기타 타입은 문자열로 변환
    return str(value).strip()

def read_protected_excel(file_path, password=None, **kwargs):
    """
    암호로 보호된 Excel 파일을 읽는 함수
    msoffcrypto-tool이 설치되어 있으면 사용하고, 없으면 기본 pandas 사용
    """
    try:
        # 먼저 암호 없이 시도
        return pd.read_excel(file_path, engine='openpyxl', **kwargs)
    except Exception as e:
        if password is None:
            logging.error(f"암호 보호된 파일이지만 암호가 제공되지 않았습니다: {file_path}")
            raise e
        
        # msoffcrypto-tool 사용 시도
        try:
            import msoffcrypto
            
            with open(file_path, 'rb') as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=password)
                
                # 메모리에서 해독된 파일 처리 (최신 버전 호환)
                decrypted = io.BytesIO()
                try:
                    # 최신 버전: decrypt 메서드 사용
                    office_file.decrypt(decrypted)
                except AttributeError:
                    # 이전 버전: save 메서드 사용
                    office_file.save(decrypted)
                
                decrypted.seek(0)
                return pd.read_excel(decrypted, engine='openpyxl', **kwargs)
                
        except ImportError:
            logging.error("msoffcrypto-tool이 설치되지 않았습니다.")
            logging.error("해결 방법: pip install msoffcrypto-tool")
            logging.error("또는 Excel에서 파일을 열어 암호를 제거한 후 저장하세요.")
            raise ImportError("msoffcrypto-tool 라이브러리가 필요합니다. 'pip install msoffcrypto-tool'로 설치하세요.")
        except Exception as decrypt_error:
            logging.error(f"암호 해독 실패: {decrypt_error}")
            logging.error("암호가 올바른지 확인하거나 Excel에서 수동으로 암호를 제거해보세요.")
            raise decrypt_error

def get_reward_for_date_and_product(product_id, date_str):
    """날짜와 상품ID에 해당하는 리워드 값 조회 (안전한 버전)"""
    try:
        reward_file = os.path.join(config.BASE_DIR, '리워드설정.json')
        
        # 파일 존재 확인
        if not os.path.exists(reward_file):
            return 0
        
        # 파일 크기 확인 (빈 파일 체크)
        if os.path.getsize(reward_file) == 0:
            return 0
        
        # JSON 파일 읽기
        with open(reward_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 데이터 구조 검증
        if not isinstance(data, dict) or 'rewards' not in data:
            return 0
        
        rewards_list = data.get('rewards', [])
        if not isinstance(rewards_list, list):
            return 0
        
        # 날짜 파싱 (여러 형식 지원)
        target_date = None
        for date_format in ['%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d']:
            try:
                target_date = datetime.strptime(date_str, date_format).date()
                break
            except ValueError:
                continue
        
        if target_date is None:
            logging.warning(f"날짜 형식을 파싱할 수 없습니다: {date_str}")
            return 0
        
        # 해당 상품과 날짜에 맞는 리워드 찾기
        for reward_entry in rewards_list:
            try:
                # 필수 키 존재 확인
                if not all(k in reward_entry for k in ['start_date', 'end_date', 'product_id', 'reward']):
                    continue
                
                start_date = datetime.strptime(reward_entry['start_date'], '%Y-%m-%d').date()
                end_date = datetime.strptime(reward_entry['end_date'], '%Y-%m-%d').date()
                
                # 상품ID 정규화하여 비교 (JSON의 .0도 제거)
                normalized_entry_id = normalize_product_id(reward_entry['product_id'])
                normalized_target_id = normalize_product_id(product_id)
                
                logging.debug(f"리워드 비교: JSON ID='{normalized_entry_id}' vs 타겟 ID='{normalized_target_id}'")
                
                if (start_date <= target_date <= end_date and 
                    normalized_entry_id == normalized_target_id):
                    reward_value = reward_entry['reward']
                    # 리워드 값이 숫자인지 확인
                    if isinstance(reward_value, (int, float)) and reward_value >= 0:
                        return int(reward_value)
            except (ValueError, KeyError, TypeError) as e:
                # 개별 엔트리 파싱 실패는 로그만 남기고 계속 진행
                continue
        
        return 0  # 설정이 없으면 0
        
    except FileNotFoundError:
        return 0
    except json.JSONDecodeError as e:
        logging.warning(f"JSON 파일 형식 오류: {e}")
        return 0
    except Exception as e:
        logging.warning(f"리워드 조회 중 예상치 못한 오류: {e}")
        return 0

def get_purchase_count_for_date_and_product(product_id, date_str):
    """날짜와 상품ID에 해당하는 가구매 개수 조회 (리워드 방식과 동일)"""
    try:
        purchase_file = os.path.join(config.BASE_DIR, '가구매설정.json')
        
        # 파일 존재 확인
        if not os.path.exists(purchase_file):
            return 0
        
        # 파일 크기 확인 (빈 파일 체크)
        if os.path.getsize(purchase_file) == 0:
            return 0
        
        # JSON 파일 읽기
        with open(purchase_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 데이터 구조 검증
        if not isinstance(data, dict) or 'purchases' not in data:
            return 0
        
        purchases_list = data.get('purchases', [])
        if not isinstance(purchases_list, list):
            return 0
        
        # 날짜 파싱 (여러 형식 지원)
        target_date = None
        for date_format in ['%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d']:
            try:
                target_date = datetime.strptime(date_str, date_format).date()
                break
            except ValueError:
                continue
        
        if target_date is None:
            logging.warning(f"가구매 개수 조회: 날짜 형식을 파싱할 수 없습니다: {date_str}")
            return 0
        
        # 해당 상품과 날짜에 맞는 가구매 개수 찾기
        for purchase_entry in purchases_list:
            try:
                # 필수 키 존재 확인
                if not all(k in purchase_entry for k in ['start_date', 'end_date', 'product_id', 'purchase_count']):
                    continue
                
                start_date = datetime.strptime(purchase_entry['start_date'], '%Y-%m-%d').date()
                end_date = datetime.strptime(purchase_entry['end_date'], '%Y-%m-%d').date()
                
                # 상품ID 정규화하여 비교
                normalized_entry_id = normalize_product_id(purchase_entry['product_id'])
                normalized_target_id = normalize_product_id(product_id)
                
                if (start_date <= target_date <= end_date and 
                    normalized_entry_id == normalized_target_id):
                    purchase_count = purchase_entry['purchase_count']
                    # 가구매 개수가 숫자인지 확인
                    if isinstance(purchase_count, (int, float)) and purchase_count >= 0:
                        return int(purchase_count)
            except (ValueError, KeyError, TypeError) as e:
                # 개별 엔트리 파싱 실패는 로그만 남기고 계속 진행
                continue
        
        return 0  # 설정이 없으면 0
        
    except FileNotFoundError:
        return 0
    except json.JSONDecodeError as e:
        logging.warning(f"가구매 설정 JSON 파일 형식 오류: {e}")
        return 0
    except Exception as e:
        logging.warning(f"가구매 개수 조회 중 예상치 못한 오류: {e}")
        return 0

def generate_individual_reports():
    """개별 스토어의 주문조회 파일을 기반으로 옵션별 통합 리포트를 생성합니다."""
    logging.info("--- 1단계: 주문조회 기반 개별 통합 리포트 생성 시작 ---")
    
    # 마진정보 파일 로드 및 검증
    try:
        margin_df = pd.read_excel(config.MARGIN_FILE, engine='openpyxl')
        logging.info(f"'{os.path.basename(config.MARGIN_FILE)}' 파일을 성공적으로 불러왔습니다.")
        
        # 필수 컬럼 존재 확인
        required_columns = ['상품번호', '상품명', '판매가', '마진율']
        missing_columns = [col for col in required_columns if col not in margin_df.columns]
        if missing_columns:
            raise ValueError(f"마진정보 파일에 필수 컬럼이 없습니다: {missing_columns}")
        
        # 컬럼명 정규화
        margin_df = margin_df.rename(columns={'상품번호': '상품ID'})
        
        # 상품ID 데이터 타입 정규화 (문자열/숫자 모두 처리)
        margin_df['상품ID'] = margin_df['상품ID'].apply(normalize_product_id)
        if margin_df['상품ID'].isna().any():
            logging.warning("마진정보에 빈 상품ID가 있습니다. 해당 행들을 제거합니다.")
            margin_df = margin_df.dropna(subset=['상품ID'])
        
        # 데이터 타입 검증 및 변환
        if not pd.api.types.is_numeric_dtype(margin_df['판매가']):
            logging.warning("판매가 컬럼이 숫자 타입이 아닙니다. 변환을 시도합니다.")
            margin_df['판매가'] = pd.to_numeric(margin_df['판매가'], errors='coerce')
        
        if not pd.api.types.is_numeric_dtype(margin_df['마진율']):
            logging.warning("마진율 컬럼이 숫자 타입이 아닙니다. 변환을 시도합니다.")
            margin_df['마진율'] = pd.to_numeric(margin_df['마진율'], errors='coerce')
        
        # 대표옵션 정보 처리
        if '대표옵션' in margin_df.columns:
            margin_df['대표옵션'] = margin_df['대표옵션'].astype(str).str.upper().isin(['O', 'Y', 'TRUE'])
            rep_price_map = margin_df[margin_df['대표옵션'] == True].set_index('상품ID')['판매가'].to_dict()
            logging.info("대표옵션 판매가 정보를 생성했습니다.")
        else:
            logging.warning(f"경고: '{os.path.basename(config.MARGIN_FILE)}'에 '대표옵션' 컬럼이 없습니다.")
            margin_df['대표옵션'] = False
            rep_price_map = {}
            
        # 옵션정보 정규화 (마진정보) - pandas의 nullable 데이터 처리 모범사례 적용
        def normalize_option_info(value):
            """옵션정보 정규화 - pandas.isna()로 모든 NA 타입 처리"""
            if pd.isna(value):
                return ''
            
            value_str = str(value).strip()
            if value_str == '' or value_str.lower() in ['단일', '기본옵션', '선택안함', 'null', 'none', '없음']:
                return ''
            
            return value_str
            
        if '옵션정보' not in margin_df.columns:
            margin_df['옵션정보'] = ''
        else:
            margin_df['옵션정보'] = margin_df['옵션정보'].apply(normalize_option_info)
            
    except FileNotFoundError:
        logging.error(f"마진정보 파일을 찾을 수 없습니다: {config.MARGIN_FILE}")
        return []
    except PermissionError:
        logging.error(f"마진정보 파일에 접근할 수 없습니다: {config.MARGIN_FILE}")
        return []
    except ValueError as e:
        logging.error(f"마진정보 파일 데이터 검증 실패: {e}")
        return []
    except Exception as e:
        logging.error(f"마진정보 파일 읽기 중 예상치 못한 오류: {e}")
        return []

    # 처리 가능한 파일들 찾기
    all_files = [f for f in os.listdir(config.get_processing_dir()) if f.endswith('.xlsx') and not f.startswith('~')]
    source_files = [f for f in all_files if '통합_리포트' not in f and '마진정보' not in f]

    # 주문조회 파일만 필터링
    order_files = [f for f in source_files if '스마트스토어_주문조회' in f]
    
    if not order_files:
        logging.info("처리할 주문조회 파일이 없습니다.")
        return True

    logging.info(f"총 {len(order_files)}개의 주문조회 파일에 대한 리포트를 생성합니다.")
    processed_groups = []
    
    for order_file in order_files:
        # 파일명에서 스토어명과 날짜 추출
        if '스마트스토어_주문조회' in order_file:
            parts = order_file.split(' 스마트스토어_주문조회_')
            if len(parts) == 2:
                store = parts[0]
                date = parts[1].replace('.xlsx', '')
            else:
                continue
        else:
            continue
            
        output_filename = f'{store}_통합_리포트_{date}.xlsx'
        output_path = os.path.join(config.get_processing_dir(), output_filename)
        
        # 이미 리포트가 존재하는지 확인
        if os.path.exists(output_path):
            logging.info(f"- {store} ({date}) 이미 리포트가 생성되어 있습니다.")
            processed_groups.append((store, date))
            continue
            
        logging.info(f"- {store} ({date}) 주문조회 기반 데이터 처리 시작...")
        
        try:
            # 주문조회 파일 읽기 (암호 보호될 수 있음)
            order_path = os.path.join(config.get_processing_dir(), order_file)
            order_df = read_protected_excel(order_path, password=config.ORDER_FILE_PASSWORD)
            
            # 파일이 비어있는지 확인
            if order_df.empty:
                logging.error(f"-> {store}({date}) 주문조회 파일이 비어있습니다: {order_file}")
                continue
            
            logging.info(f"-> {store}({date}) 주문조회 파일 로드 완료: {len(order_df)}행")
            logging.info(f"-> {store}({date}) 주문조회 파일 컬럼: {list(order_df.columns)}")
            
            # 상품번호 -> 상품ID 변환 (컬럼이 있는 경우에만)
            if '상품번호' in order_df.columns:
                order_df = order_df.rename(columns={'상품번호': '상품ID'})
            
            # 필수 컬럼 존재 확인
            required_cols = ['상품ID']
            missing_cols = [col for col in required_cols if col not in order_df.columns]
            if missing_cols:
                logging.error(f"-> {store}({date}) 필수 컬럼 누락: {missing_cols}")
                continue
            
            # 상품ID 데이터 타입 정규화 (마진정보와 동일한 방식)
            order_df['상품ID'] = order_df['상품ID'].apply(normalize_product_id)
            
            # 옵션정보 정규화 
            def normalize_option_info(value):
                if pd.isna(value) or value == '' or str(value).strip() == '':
                    return ''
                value_str = str(value).strip()
                # '단일', '기본옵션', '선택안함' 등을 빈 문자열로 통일
                if value_str.lower() in ['단일', '기본옵션', '선택안함', 'null', 'none', '없음']:
                    return ''
                return value_str
            
            if '옵션정보' not in order_df.columns:
                order_df['옵션정보'] = ''
            else:
                order_df['옵션정보'] = order_df['옵션정보'].apply(normalize_option_info)
            
            logging.info(f"-> {store}({date}) 옵션정보 정규화 후 샘플: {order_df['옵션정보'].head(5).tolist()}")
            
            # 클레임상태 컬럼 확인 및 환불 관련 처리
            if '클레임상태' not in order_df.columns:
                # 다른 가능한 컬럼명들 확인
                possible_status_cols = ['상태', '주문상태', '처리상태', '배송상태', '주문처리상태', '결제상태']
                status_col = None
                for col in possible_status_cols:
                    if col in order_df.columns:
                        status_col = col
                        break
                
                if status_col:
                    logging.info(f"-> {store}({date}) '{status_col}' 컬럼을 클레임상태로 사용합니다.")
                    order_df['클레임상태'] = order_df[status_col]
                else:
                    logging.warning(f"-> {store}({date}) 클레임상태 컬럼을 찾을 수 없습니다.")
                    order_df['클레임상태'] = '정상'
            
            # 수량 컬럼 확인
            if '수량' not in order_df.columns:
                possible_quantity_cols = ['결제수량', '주문수량', '상품수량', '결제상품수량']
                quantity_col = None
                for col in possible_quantity_cols:
                    if col in order_df.columns:
                        quantity_col = col
                        break
                
                if quantity_col:
                    logging.info(f"-> {store}({date}) '{quantity_col}' 컬럼을 수량으로 사용합니다.")
                    order_df['수량'] = order_df[quantity_col]
                else:
                    logging.warning(f"-> {store}({date}) 수량 컬럼을 찾을 수 없습니다. 기본값 1 사용")
                    order_df['수량'] = 1
            
            # 수량을 숫자형으로 변환
            order_df['수량'] = pd.to_numeric(order_df['수량'], errors='coerce').fillna(1)
            
            # 클레임상태 분포 확인
            status_counts = order_df['클레임상태'].value_counts()
            logging.info(f"-> {store}({date}) 클레임상태 분포: {status_counts.to_dict()}")
            
            # 환불수량 계산
            cancel_mask = order_df['클레임상태'].isin(config.CANCEL_OR_REFUND_STATUSES)
            order_df['환불수량'] = order_df['수량'].where(cancel_mask, 0)
            
            # 환불수량 계산 결과
            total_refund_quantity = order_df['환불수량'].sum()
            refund_rows = (order_df['환불수량'] > 0).sum()
            logging.info(f"-> {store}({date}) 총 환불수량: {total_refund_quantity}, 환불 행 수: {refund_rows}")
            
            # 옵션별 집계 (핵심 로직!) - 상품명도 함께 집계
            logging.info(f"-> {store}({date}) 옵션별 데이터 집계 시작...")
            
            # 상품명 컬럼 확인
            if '상품명' in order_df.columns:
                group_cols = ['상품ID', '상품명', '옵션정보']
                agg_dict = {
                    '수량': 'sum',           # 옵션별 총 판매수량
                    '환불수량': 'sum'        # 옵션별 총 환불수량
                }
            else:
                group_cols = ['상품ID', '옵션정보'] 
                agg_dict = {
                    '수량': 'sum',           # 옵션별 총 판매수량
                    '환불수량': 'sum'        # 옵션별 총 환불수량
                }
                logging.warning(f"-> {store}({date}) 주문조회 파일에 상품명 컬럼이 없습니다.")
            
            # 중복 데이터 검증
            duplicates = order_df.duplicated(group_cols).sum()
            if duplicates > 0:
                logging.warning(f"-> {store}({date}) 주문조회 데이터에 중복된 상품ID-옵션정보 조합이 {duplicates}개 있습니다.")
            
            option_summary = order_df.groupby(group_cols, as_index=False).agg(agg_dict)
            
            logging.info(f"-> {store}({date}) 옵션별 집계 완료: {len(option_summary)}개 옵션")
            
            # 판매가는 마진정보 파일에서만 가져옴 (주문조회 파일에는 판매가 컬럼이 없음)
            logging.info(f"-> {store}({date}) 판매가는 마진정보 파일에서 가져옵니다.")
            
            # 병합 전 데이터 확인
            logging.info(f"-> {store}({date}) 병합 전 주문조회 상품ID 샘플: {option_summary['상품ID'].head(3).tolist()}")
            logging.info(f"-> {store}({date}) 병합 전 주문조회 옵션정보 샘플: {option_summary['옵션정보'].head(3).tolist()}")
            logging.info(f"-> {store}({date}) 병합 전 마진정보 상품ID 샘플: {margin_df['상품ID'].head(3).tolist()}")
            logging.info(f"-> {store}({date}) 병합 전 마진정보 옵션정보 샘플: {margin_df['옵션정보'].head(3).tolist()}")
            
            # 마진정보와 안전한 병합 with 검증
            logging.info(f"-> {store}({date}) 마진정보와 병합 시작...")
            
            # 병합 전 마진정보 중복 검증
            margin_duplicates = margin_df.duplicated(['상품ID', '옵션정보']).sum()
            if margin_duplicates > 0:
                logging.warning(f"-> {store}({date}) 마진정보에 중복된 상품ID-옵션정보 조합이 {margin_duplicates}개 있습니다.")
                # 첫 번째 값만 유지
                margin_df = margin_df.drop_duplicates(['상품ID', '옵션정보'], keep='first')
                logging.info(f"-> {store}({date}) 중복 제거 후 마진정보 행 수: {len(margin_df)}")
            
            # 마진정보에서 상품명 컬럼 제거 (주문조회의 상품명 유지)
            margin_cols_to_use = [col for col in margin_df.columns if col != '상품명']
            margin_df_clean = margin_df[margin_cols_to_use].copy()
            
            try:
                # 안전한 병합 with validation (상품명은 주문조회에서만 사용)
                final_df = pd.merge(
                    option_summary, 
                    margin_df_clean, 
                    on=['상품ID', '옵션정보'], 
                    how='left',
                    validate='many_to_one'  # 마진정보의 각 상품-옵션은 고유해야 함
                )
            except pd.errors.MergeError as e:
                logging.error(f"-> {store}({date}) 병합 검증 실패: {e}")
                # validation 없이 재시도
                final_df = pd.merge(option_summary, margin_df_clean, on=['상품ID', '옵션정보'], how='left')
            
            # 병합 결과 확인
            merged_count = len(final_df)
            margin_matched = final_df['마진율'].notna().sum()
            logging.info(f"-> {store}({date}) 병합 완료: {merged_count}행, 마진 매칭 {margin_matched}행")
            
            # 매칭 실패한 경우 디버깅 정보 및 변드을 통한 대안 매칭 시도
            if margin_matched == 0:
                logging.warning(f"-> {store}({date}) 마진정보 매칭 실패! 디버깅 정보:")
                logging.warning(f"   주문조회 고유 상품ID: {option_summary['상품ID'].unique()[:5]}")
                logging.warning(f"   마진정보 고유 상품ID: {margin_df['상품ID'].unique()[:5]}")
                logging.warning(f"   주문조회 고유 옵션정보: {option_summary['옵션정보'].unique()[:5]}")
                logging.warning(f"   마진정보 고유 옵션정보: {margin_df['옵션정보'].unique()[:5]}")
                
                # 상품ID만으로 대안 매칭 시도 (옵션 무시)
                logging.info(f"-> {store}({date}) 옵션정보 없이 상품ID만으로 대안 매칭 시도...")
                
                # 빈 옵션정보만 필터링하여 대안 매칭 (상품명도 제외)
                margin_df_no_option = margin_df[margin_df['옵션정보'] == ''].copy()
                if len(margin_df_no_option) > 0:
                    # 옵션정보와 상품명 모두 제외
                    alt_cols = margin_df_no_option.columns.difference(['옵션정보', '상품명'])
                    final_df_alt = pd.merge(
                        option_summary, 
                        margin_df_no_option[alt_cols], 
                        on='상품ID', 
                        how='left'
                    )
                    alt_matched = final_df_alt['마진율'].notna().sum()
                    if alt_matched > 0:
                        logging.info(f"-> {store}({date}) 대안 매칭 성공: {alt_matched}개 상품 매칭")
                        # 옵션정보 컬럼 다시 추가
                        final_df_alt['옵션정보'] = option_summary['옵션정보']
                        final_df = final_df_alt
                        margin_matched = alt_matched
            
            # 기본값 설정 및 데이터 타입 검증
            numeric_columns = ['마진율', '판매가', '개당 가구매 비용']
            for col in numeric_columns:
                if col in final_df.columns:
                    # 숫자 타입을 강제로 변환
                    final_df[col] = pd.to_numeric(final_df[col], errors='coerce')
            
            final_df.fillna({
                '마진율': 0.0, 
                '판매가': 0.0,  # 마진정보의 판매가
                '개당 가구매 비용': 0.0, 
                '대표옵션': False
            }, inplace=True)
            
            # 상품명 확인 (마진정보에서 상품명을 제외했으므로 주문조회의 상품명이 유지됨)
            logging.info(f"-> {store}({date}) 상품명 확인 - 현재 컬럼: {list(final_df.columns)}")
            
            if '상품명' not in final_df.columns:
                logging.error(f"-> {store}({date}) 상품명 컬럼을 찾을 수 없습니다!")
                # 응급 처치: 상품ID를 상품명으로 사용
                final_df['상품명'] = final_df['상품ID']
                logging.warning(f"-> {store}({date}) 임시로 상품ID를 상품명으로 사용합니다.")
            else:
                logging.info(f"-> {store}({date}) 상품명 유지 완료 - 샘플: {final_df['상품명'].head(2).tolist()}")
            
            # 기본 계산 필드들
            final_df['결제금액'] = final_df['수량'] * final_df['판매가']
            final_df['환불금액'] = final_df['환불수량'] * final_df['판매가'] 
            final_df['매출'] = final_df['결제금액'] - final_df['환불금액']
            
            # 대표판매가 (가구매 금액 계산용)
            final_df['대표판매가'] = final_df['상품ID'].map(rep_price_map).fillna(0)
            
            # 가구매 개수 적용 (대표옵션에만, GUI에서 설정한 값)
            final_df['가구매 개수'] = 0  # 기본값
            rep_option_mask = final_df['대표옵션'] == True
            
            if rep_option_mask.sum() > 0:
                for product_id in final_df.loc[rep_option_mask, '상품ID'].unique():
                    purchase_count = get_purchase_count_for_date_and_product(product_id, date)
                    final_df.loc[(final_df['상품ID'] == product_id) & rep_option_mask, '가구매 개수'] = purchase_count
                    if purchase_count > 0:
                        logging.info(f"-> {store}({date}) 상품 {product_id} 가구매 개수: {purchase_count}")
            
            # 추가 계산 필드들
            final_df['가구매 수량'] = final_df['가구매 개수']
            final_df['개당 가구매 금액'] = final_df['대표판매가']
            final_df['가구매 금액'] = final_df['개당 가구매 금액'] * final_df['가구매 수량']
            final_df['순매출'] = final_df['매출'] - final_df['가구매 금액']
            final_df['가구매 비용'] = final_df['개당 가구매 비용'] * final_df['가구매 수량']
            
            # 리워드 적용 (대표옵션에만)
            final_df['리워드'] = 0
            if rep_option_mask.sum() > 0:
                for product_id in final_df.loc[rep_option_mask, '상품ID'].unique():
                    reward_value = get_reward_for_date_and_product(product_id, date)
                    final_df.loc[(final_df['상품ID'] == product_id) & rep_option_mask, '리워드'] = reward_value
                    if reward_value > 0:
                        logging.info(f"-> {store}({date}) 상품 {product_id} 리워드: {reward_value}원")
            
            # 안전한 나누기 함수 정의
            def safe_divide(numerator, denominator, fill_value=0.0):
                """안전한 나누기 - 0 나누기와 NaN 처리"""
                with np.errstate(divide='ignore', invalid='ignore'):
                    result = np.where(
                        (denominator == 0) | pd.isna(denominator),
                        fill_value,
                        numerator / denominator
                    )
                return result
            
            # 판매마진 및 비율 계산 (안전한 방식)
            final_df['판매마진'] = final_df['순매출'] * final_df['마진율']
            
            # 광고비율 = (리워드 + 가구매 비용) / 순매출
            final_df['광고비율'] = safe_divide(
                final_df['리워드'] + final_df['가구매 비용'],
                final_df['순매출'],
                fill_value=0.0  # 순매출이 0이면 광고비율은 0%
            )
            
            final_df['이윤율'] = final_df['마진율'] - final_df['광고비율']
            final_df['순이익'] = final_df['판매마진'] - final_df['가구매 비용'] - final_df['리워드']
            
            # 퍼센트 값 변환
            final_df['마진율'] = (final_df['마진율'] * 100).round(1)
            final_df['광고비율'] = (final_df['광고비율'] * 100).round(1)
            final_df['이윤율'] = (final_df['이윤율'] * 100).round(1)
            
            # 결제수, 환불건수 계산 (주문조회 기반)
            if '상품주문번호' in order_df.columns:
                # 결제수 (상품주문번호 개수)
                order_count = order_df.groupby(['상품ID', '옵션정보'])['상품주문번호'].nunique().reset_index()
                order_count.rename(columns={'상품주문번호': '결제수'}, inplace=True)
                final_df = pd.merge(final_df, order_count, on=['상품ID', '옵션정보'], how='left')
                final_df['결제수'] = final_df['결제수'].fillna(0)
                
                # 환불건수 (환불 상태인 주문번호 개수)  
                cancel_orders = order_df[order_df['클레임상태'].isin(config.CANCEL_OR_REFUND_STATUSES)]
                if not cancel_orders.empty:
                    refund_count = cancel_orders.groupby(['상품ID', '옵션정보'])['상품주문번호'].nunique().reset_index()
                    refund_count.rename(columns={'상품주문번호': '환불건수'}, inplace=True)
                    final_df = pd.merge(final_df, refund_count, on=['상품ID', '옵션정보'], how='left')
                    final_df['환불건수'] = final_df['환불건수'].fillna(0)
                else:
                    final_df['환불건수'] = 0
            else:
                final_df['결제수'] = 0
                final_df['환불건수'] = 0
                
            # 최종 컬럼 정리
            final_columns = [col for col in config.COLUMNS_TO_KEEP if col in final_df.columns]
            sorted_df = final_df[final_columns].sort_values(by=['상품명', '옵션정보'])
            
            # 데이터 요약 로깅
            logging.info(f"-> {store}({date}) 최종 데이터 요약:")
            logging.info(f"   - 총 옵션 수: {len(sorted_df)}")
            logging.info(f"   - 총 판매수량: {sorted_df['수량'].sum()}")
            logging.info(f"   - 총 환불수량: {sorted_df['환불수량'].sum()}")
            logging.info(f"   - 총 매출: {sorted_df['매출'].sum():,.0f}원")
            logging.info(f"   - 총 판매마진: {sorted_df['판매마진'].sum():,.0f}원")
            
            # 엑셀 파일 생성
            pivot_quantity = pd.pivot_table(sorted_df, index='상품명', columns='옵션정보', values='수량', aggfunc='sum', fill_value=0)
            pivot_margin = pd.pivot_table(sorted_df, index='상품명', columns='옵션정보', values='판매마진', aggfunc='sum', fill_value=0)
            
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                sorted_df.to_excel(writer, sheet_name='정리된 데이터', index=False)
                pivot_quantity.to_excel(writer, sheet_name='옵션별 판매수량')
                pivot_margin.to_excel(writer, sheet_name='옵션별 판매마진')
                
                # 표 서식 적용
                worksheet = writer.sheets['정리된 데이터']
                (max_row, max_col) = sorted_df.shape
                worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': [{'header': col} for col in sorted_df.columns]})
                for i, col in enumerate(sorted_df.columns):
                    col_len = max(sorted_df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, col_len)
            
            # 생성 완료 확인
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                logging.info(f"-> '{output_filename}' 생성 완료: (파일 크기: {file_size:,} bytes)")
                processed_groups.append((store, date))
            else:
                logging.error(f"-> 파일 생성 실패: {output_path}")
                
        except Exception as e:
            logging.error(f"-> {store}({date}) 처리 중 오류 발생: {e}")
            import traceback
            logging.error(f"-> {store}({date}) 상세 오류: {traceback.format_exc()}")
        finally:
            # 메모리 정리
            try:
                if 'order_df' in locals():
                    del order_df
                if 'final_df' in locals():
                    del final_df
                if 'sorted_df' in locals():
                    del sorted_df
            except:
                pass
    
    logging.info("--- 1단계: 주문조회 기반 개별 통합 리포트 생성 완료 ---")
    return processed_groups

def consolidate_daily_reports():
    """날짜별로 생성된 모든 개별 리포트를 취합하여 전체 통합 리포트를 생성합니다."""
    logging.info("--- 2단계: 전체 통합 리포트 생성 시작 ---")
    all_report_files = [f for f in glob.glob(os.path.join(config.get_processing_dir(), '*_통합_리포트_*.xlsx')) if not os.path.basename(f).startswith('~') and not os.path.basename(f).startswith('전체_')]
    if not all_report_files:
        logging.info("취합할 개별 통합 리포트가 없습니다.")
        return

    date_pattern = re.compile(r'_(\d{4}-\d{2}-\d{2})\.xlsx$')
    unique_dates = set()
    for f in all_report_files:
        match = date_pattern.search(os.path.basename(f))
        if match:
            unique_dates.add(match.group(1))
    
    if not unique_dates:
        logging.info("파일에서 날짜 정보를 찾을 수 없습니다.")
        return

    logging.info(f"총 {len(sorted(list(unique_dates)))}개의 날짜에 대한 전체 리포트를 생성합니다: {sorted(list(unique_dates))}")
    logging.info(f"처리할 개별 리포트 파일 수: {len(all_report_files)}")
    for date in sorted(list(unique_dates)):
        logging.info(f"- {date} 데이터 통합 중...")
        output_file = os.path.join(config.get_processing_dir(), f'전체_통합_리포트_{date}.xlsx')
        daily_files = [f for f in all_report_files if date in f]
        logging.info(f"-> {date} 날짜에 대한 개별 파일 수: {len(daily_files)}")
        daily_dfs = []
        for file_path in daily_files:
            try:
                store_name = os.path.basename(file_path).split('_통합_리포트_')[0]
                df = pd.read_excel(file_path, sheet_name='정리된 데이터', engine='openpyxl')
                df['스토어명'] = store_name
                daily_dfs.append(df)
                logging.info(f"-> '{os.path.basename(file_path)}' 통합 완료: {len(df)}행 데이터 추가")
            except Exception as e:
                logging.error(f"-> '{os.path.basename(file_path)}' 처리 중 오류: {e}")
        
        if daily_dfs:
            total_rows_before = sum(len(df) for df in daily_dfs)
            logging.info(f"-> {date} 날짜 병합 전 총 데이터 행 수: {total_rows_before}")
            
            master_df = pd.concat(daily_dfs, ignore_index=True)
            logging.info(f"-> {date} 날짜 병합 후 데이터 행 수: {len(master_df)}")
            master_df = master_df[['스토어명'] + [col for col in master_df.columns if col != '스토어명']]
            grouping_keys = ['스토어명', '상품ID', '상품명', '옵션정보']
            agg_methods = {
                '수량': 'sum', '판매마진': 'sum', '결제수': 'sum', '결제금액': 'sum',
                '환불건수': 'sum', '환불금액': 'sum', '환불수량': 'sum',
                '가구매 개수': 'sum', '판매가': 'mean', '마진율': 'mean',
                '가구매 비용': 'sum', '순매출': 'sum', '매출': 'sum', '가구매 금액': 'sum',
                '이윤율': 'mean', '광고비율': 'mean', '순이익': 'sum', '리워드': 'sum'
            }
            actual_agg_methods = {k: v for k, v in agg_methods.items() if k in master_df.columns}
            logging.info(f"-> {date} 날짜 집계 전 데이터 행 수: {len(master_df)}, 사용 가능한 집계 컬럼: {list(actual_agg_methods.keys())}")
            
            aggregated_df = master_df.groupby(grouping_keys, as_index=False).agg(actual_agg_methods)
            logging.info(f"-> {date} 날짜 집계 후 데이터 행 수: {len(aggregated_df)}")
            
            # 퍼센트 필드들을 소수점 첫 자리까지 반올림
            for col in ['마진율', '광고비율', '이윤율']:
                if col in aggregated_df.columns:
                    aggregated_df[col] = aggregated_df[col].round(1)
            
            final_columns = ['스토어명'] + [col for col in config.COLUMNS_TO_KEEP if col in aggregated_df.columns]
            logging.info(f"-> {date} 날짜 최종 컬럼 수: {len(final_columns)}, 컬럼: {final_columns[:10]}...")  # 처음 10개만
            
            aggregated_df = aggregated_df[final_columns]
            logging.info(f"-> {date} 날짜 최종 데이터: {len(aggregated_df)}행")
            try:
                with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                    aggregated_df.to_excel(writer, sheet_name='전체 통합 데이터', index=False)
                    worksheet = writer.sheets['전체 통합 데이터']
                    (max_row, max_col) = aggregated_df.shape
                    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': [{'header': col} for col in aggregated_df.columns]})
                    for i, col in enumerate(aggregated_df.columns):
                        col_len = max(aggregated_df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(i, i, col_len)
                if os.path.exists(output_file):
                    file_size = os.path.getsize(output_file)
                    logging.info(f"-> '{os.path.basename(output_file)}' 생성 완료: {output_file} (파일 크기: {file_size:,} bytes)")
                    
                    # 생성된 파일 내용 검증
                    try:
                        verify_df = pd.read_excel(output_file, sheet_name='전체 통합 데이터')
                        logging.info(f"-> 검증: 전체 통합 리포트에 {len(verify_df)}행 데이터 저장됨")
                    except Exception as verify_e:
                        logging.error(f"-> 전체 리포트 검증 중 오류: {verify_e}")
                else:
                    logging.error(f"-> 전체 리포트 생성 실패: {output_file} 파일이 생성되지 않음")
            except Exception as e:
                logging.error(f"-> 최종 파일 저장 중 오류: {e}")
            finally:
                # 메모리 정리
                try:
                    del master_df, aggregated_df
                except:
                    pass
        
        # daily_dfs 메모리 정리
        try:
            del daily_dfs
        except:
            pass
        else:
            logging.warning(f"-> {date} 날짜에 대한 개별 리포트가 없어 전체 리포트를 생성할 수 없습니다.")
            
    logging.info("--- 2단계: 전체 통합 리포트 생성 완료 ---")
