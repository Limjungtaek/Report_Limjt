import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('데이터를 처리 중입니다...'):
        try:
            # 1. 엑셀 파일 읽기
            # 첫 번째 시트: D열 연산용, 두 번째 시트: P열 연산용
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- 시트 1 (D열 연산) ---
            target_d = df_sheet1.iloc[5:10000, 3] # D열(3번 인덱스)
            count_total_d = target_d.count()
            clean_d = target_d.astype(str).str.strip()
            count_normal = len(clean_d[clean_d == '정상'])
            count_closed = len(clean_d[clean_d == '폐업'])
            
            # --- 시트 2 (P열 연산) ---
            # P열은 A(0)부터 순서대로 세면 15번째(16번째 열)입니다.
            target_p = df_sheet2.iloc[4:10000, 15] # 5행부터 10000행까지(인덱스 4:10000)
            
            count_total_p = target_p.count() # 전체 값 개수
            clean_p = target_p.astype(str).str.strip()
            count_out = len(clean_p[clean_p == '출고'])
            count_hold = len(clean_p[clean_p == '보유'])
            
            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # 시트 1 결과 기입 (B7, D7, F7)
                ws['B7'] = count_total_d
                ws['D7'] = count_normal
                ws['F7'] = count_closed
                
                # 시트 2 결과 기입 (J7, L7, N7)
                ws['J7'] = count_total_p
                ws['L7'] = count_out
                ws['N7'] = count_hold
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download_name)
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: (데이터 확인 필요) {e}")
