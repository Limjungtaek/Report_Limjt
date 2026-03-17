import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('시트 1과 시트 2의 데이터를 통합 집계 중입니다...'):
        try:
            # 1. 인풋 파일 읽기
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- [시트1 집계 로직] (B, D, E, F열 관련) ---
            data_s1 = df_sheet1.iloc[:, [3, 4]].copy()
            data_s1.columns = ['status', 'item_name']
            data_s1['status'] = data_s1['status'].astype(str).str.strip()
            data_s1['item_name'] = data_s1['item_name'].astype(str).str.strip()
            
            s1_total_counts = data_s1['item_name'].value_counts()
            s1_normal_counts = data_s1[data_s1['status'] == '정상']['item_name'].value_counts()
            s1_closed_counts = data_s1[data_s1['status'] == '폐업']['item_name'].value_counts()

            # --- [시트2 집계 로직] (J, K, L, M열 관련) ---
            # D열(인덱스 3: 항목명), P열(인덱스 15: 상태) 추출
            data_s2 = df_sheet2.iloc[:, [3, 15]].copy()
            data_s2.columns = ['item_name', 'status']
            data_s2['item_name'] = data_s2['item_name'].astype(str).str.strip()
            data_s2['status'] = data_s2['status'].astype(str).str.strip()
            
            s2_total_counts = data_s2['item_name'].value_counts()
            s2_out_counts = data_s2[data_s2['status'] == '출고']['item_name'].value_counts()
            s2_hold_counts = data_s2[data_s2['status'] == '보유']['item_name'].value_counts()

            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # 상단 요약 셀 (기존 로직 유지)
                ws['B7'] = df_sheet1.iloc[5:10000, 3].count()
                ws['D7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '정상').sum()
                ws['F7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '폐업').sum()
                
                ws['J7'] = df_sheet2.iloc[4:10000, 15].count()
                ws['L7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '출고').sum()
                ws['N7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '보유').sum()
                
                # --- [11~15행 상세 기입] ---
                for row_num in range(11, 16):
                    # 왼쪽 영역 (B열 기준 -> D, E, F 채우기)
                    key_b = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                    if key_b:
                        ws[f'D{row_num}'] = s1_total_counts.get(key_b, 0)
                        ws[f'E{row_num}'] = s1_normal_counts.get(key_b, 0)
                        ws[f'F{row_num}'] = s1_closed_counts.get(key_b, 0)
                    
                    # 오른쪽 영역 (J열 기준 -> K, L, M 채우기)
                    key_j = str(ws[f'J{row_num}'].value).strip() if ws[f'J{row_num}'].value else ""
                    if key_j:
                        ws[f'K{row_num}'] = s2_total_counts.get(key_j, 0) # 전체
                        ws[f
