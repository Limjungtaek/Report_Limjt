import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('항목별 정상/폐업 데이터를 집계 중입니다...'):
        try:
            # 1. 인풋 파일 읽기
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- 시트1 기초 연산 (B7, D7, F7용) ---
            target_d = df_sheet1.iloc[5:10000, 3] 
            count_total_d = target_d.count()
            clean_d = target_d.astype(str).str.strip()
            count_normal = len(clean_d[clean_d == '정상'])
            count_closed = len(clean_d[clean_d == '폐업'])
            
            # --- 시트2 기초 연산 (J7, L7, N7용) ---
            target_p = df_sheet2.iloc[4:10000, 15] 
            count_total_p = target_p.count()
            clean_p = target_p.astype(str).str.strip()
            count_out = len(clean_p[clean_p == '출고'])
            count_hold = len(clean_p[clean_p == '보유'])
            
            # --- 항목별 상세 집계 로직 (11~15행용) ---
            # D열(인덱스 3: 상태), E열(인덱스 4: 항목명) 추출
            data_subset = df_sheet1.iloc[:, [3, 4]].copy()
            data_subset.columns = ['status', 'item_name']
            data_subset['status'] = data_subset['status'].astype(str).str.strip()
            data_subset['item_name'] = data_subset['item_name'].astype(str).str.strip()
            
            # '정상'인 항목들 개수 집계
            normal_counts = data_subset[data_subset['status'] == '정상']['item_name'].value_counts()
            # '폐업'인 항목들 개수 집계
            closed_counts = data_subset[data_subset['status'] == '폐업']['item_name'].value_counts()

            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # 상단 요약 셀 기입
                ws['B7'], ws['D7'], ws['F7'] = count_total_d, count_normal, count_closed
                ws['J7'], ws['L7'], ws['N7'] = count_total_p, count_out, count_hold
                
                # --- 11행 ~ 15행 상세 기입 ---
                for row_num in range(11, 16):
                    search_key = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                    
                    # E열: 정상 개수 / F열: 폐업 개수
                    ws[f'E{row_num}'] = normal_counts.get(search_key, 0)
                    ws[f'F{row_num}'] = closed_counts.get(search_key, 0)
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download
