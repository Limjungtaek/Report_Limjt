import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('21~23행의 출고/보유 내역을 추가 집계 중입니다...'):
        try:
            # 1. 인풋 파일 읽기
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- [시트1 집계 로직] ---
            data_s1 = df_sheet1.iloc[:, [3, 4]].copy()
            data_s1.columns = ['status', 'item_name']
            data_s1['status'] = data_s1['status'].astype(str).str.strip()
            data_s1['item_name'] = data_s1['item_name'].astype(str).str.strip()
            
            s1_total_counts = data_s1['item_name'].value_counts()
            s1_normal_counts = data_s1[data_s1['status'] == '정상']['item_name'].value_counts()
            s1_closed_counts = data_s1[data_s1['status'] == '폐업']['item_name'].value_counts()

            # --- [시트2 집계 로직] ---
            # D열(3), N열(13), O열(14), P열(15) 추출
            # 기존 J~M열용 데이터 (D열 기준) + 신규 21~23행용 데이터 (N열 기준)
            data_s2 = df_sheet2.iloc[:, [3, 13, 14, 15]].copy()
            data_s2.columns = ['item_d', 'item_n', 'sum_value', 'status']
            
            # 전처리
            for col in ['item_d', 'item_n', 'status']:
                data_s2[col] = data_s2[col].astype(str).str.strip()
            data_s2['sum_value'] = pd.to_numeric(data_s2['sum_value'], errors='coerce').fillna(0)
            
            # (A) J~M열용 (D열 기준 집계)
            s2_d_total = data_s2['item_d'].value_counts()
            s2_d_out = data_s2[data_s2['status'] == '출고']['item_d'].value_counts()
            s2_d_hold = data_s2[data_s2['status'] == '보유']['item_d'].value_counts()
            s2_d_sum = data_s2.groupby('item_d')['sum_value'].sum()

            # (B) 21~23행용 (N열 기준 집계)
            s2_n_out = data_s2[data_s2['status'] == '출고']['item_n'].value_counts()
            s2_n_hold = data_s2[data_s2['status'] == '보유']['item_n'].value_counts()

            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # --- 상단 요약 (7행) ---
                ws['B7'] = df_sheet1.iloc[5:10000, 3].count()
                ws['D7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '정상').sum()
                ws['F7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '폐업').sum()
                
                ws['I7'] = df_sheet2.iloc[4:10000, 15].count()
                ws['K7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '출고').sum()
                ws['M7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '보유').sum()
                
                # --- [상세 기입 루프 11~16행] ---
                for row_num in range(11, 17):
                    # 왼쪽 영역 (B열 기준 -> D, E, F) : 15행까지만
                    if row_num <= 15:
                        key_b = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                        if key_b:
                            ws[f'D{row_num}'] = s1_normal_counts.get(key_b, 0) + s1_closed_counts.get(key_b, 0) # 전체 (D열)
                            ws[f'E{row_num}'] = s1_normal_counts.get(key_b, 0)
                            ws[f'F{row_num}'] = s1_closed_counts.get(key_b, 0)
                    
                    # 오른쪽 영역 (I열 기준 -> J, K, L, M)
                    key_i = str(ws[f'I{row_num}'].value).strip() if ws[f'I{row_num}'].value else
