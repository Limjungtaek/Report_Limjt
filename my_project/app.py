import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('모든 조건을 "포함" 검색으로 적용하여 집계 중입니다...'):
        try:
            # 1. 인풋 파일 읽기
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- [시트1 집계 로직] ---
            data_s1 = df_sheet1.iloc[:, [3, 4]].copy()
            data_s1.columns = ['status', 'item_name']
            data_s1['status'] = data_s1['status'].astype(str).str.strip()
            data_s1['item_name'] = data_s1['item_name'].astype(str).str.strip()
            
            s1_normal_counts = data_s1[data_s1['status'] == '정상']['item_name'].value_counts()
            s1_closed_counts = data_s1[data_s1['status'] == '폐업']['item_name'].value_counts()

            # --- [시트2 집계 로직] ('출고' 및 '보유' 포함 검색 적용) ---
            data_s2 = df_sheet2.iloc[:, [3, 13, 14, 15]].copy()
            data_s2.columns = ['item_d', 'item_n', 'sum_value', 'status']
            
            # 전처리
            for col in ['item_d', 'item_n', 'status']:
                data_s2[col] = data_s2[col].astype(str).str.strip()
            data_s2['sum_value'] = pd.to_numeric(data_s2['sum_value'], errors='coerce').fillna(0)
            
            # 문자열 포함 여부 판단 (결과파일 L11~16, F21~23 등에 적용)
            is_out = data_s2['status'].str.contains('출고', na=False)
            is_hold = data_s2['status'].str.contains('보유', na=False)

            # (A) J~M열용 (인풋 D열 기준 집계)
            s2_d_total = data_s2['item_d'].value_counts()
            s2_d_out = data_s2[is_out]['item_d'].value_counts()
            s2_d_hold = data_s2[is_hold]['item_d'].value_counts()
            s2_d_sum = data_s2.groupby('item_d')['sum_value'].sum()

            # (B) 21~23행용 (인풋 N열 기준 집계)
            s2_n_out = data_s2[is_out]['item_n'].value_counts()
            s2_n_hold = data_s2[is_hold]['item_n'].value_counts()

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
                
                # 시트2 요약도 '포함' 조건으로 통일
                s2_status_7 = df_sheet2.iloc[4:10000, 15].astype(str)
                ws['I7'] = df_sheet2.iloc[4:10000, 15].count()
                ws['K7'] = s2_status_7.str.contains('출고').sum()
                ws['M7'] = s2_status_7.str.contains('보유').sum()
                
                # --- [상세 기입 루프 11~16행] ---
                for row_num in range(11, 17):
                    # 시트1 영역 (B열 기준)
                    if row_num <= 15:
                        key_b = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                        if key_b:
                            ws[f'D{row_num}'] = s1_normal_counts.get(key_b, 0) + s1_closed_counts.get(key_b, 0)
                            ws[f'E{row_num}'] = s1_normal_counts.get(key_b, 0)
                            ws[f'F{row_num}'] = s1_closed_counts.get(key_b, 0)
                    
                    # 시트2 영역 (I열 기준)
                    key_i = str(ws[f'I{row_num}'].value).strip() if ws[f'I{row_num}'].value else ""
                    if key_i:
                        ws[f'J{row_num}'] = s2_d_total.get(key_i, 0)
                        ws[f'K{row_num}'] = s2_d_out.get(key_i, 0)
                        ws[f'L{row_num}'] = s2_d_hold.get(key_i, 0) # '보유' 포함 카운트 기입
                        ws[f'M{row_num}'] = s2_d_sum.get(key_i, 0)

                # --- [상세 기입 루프 21~23행] ---
                for row_num in range(21, 24):
                    key_b_21 = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                    if key_b_21:
                        ws[f'D{row_num}'] = s2_n_out.get(key_b_21, 0) # '출고' 포함 카운트
                        ws[f'F{row_num}'] = s2_n_hold.get(key_b_21, 0) # '보유' 포함 카운트
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download_name)
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
