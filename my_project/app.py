import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('데이터를 매칭 중입니다...'):
        try:
            # 1. 인풋 파일 읽기
            df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1, header=None)
            
            # --- 기존 연산 로직 (시트1 & 시트2) ---
            target_d = df_sheet1.iloc[5:10000, 3] # D열
            count_total_d = target_d.count()
            clean_d = target_d.astype(str).str.strip()
            count_normal = len(clean_d[clean_d == '정상'])
            count_closed = len(clean_d[clean_d == '폐업'])
            
            target_p = df_sheet2.iloc[4:10000, 15] # P열
            count_total_p = target_p.count()
            clean_p = target_p.astype(str).str.strip()
            count_out = len(clean_p[clean_p == '출고'])
            count_hold = len(clean_p[clean_p == '보유'])
            
            # --- 추가 연산 로직: VLOOKUP 스타일 카운트 ---
            # 인풋파일 첫 번째 시트의 E열(인덱스 4) 전체를 가져와서 개수를 미리 세어둠
            e_column_counts = df_sheet1.iloc[:, 4].astype(str).str.strip().value_counts()

            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # 기존 셀 기입
                ws['B7'], ws['D7'], ws['F7'] = count_total_d, count_normal, count_closed
                ws['J7'], ws['L7'], ws['N7'] = count_total_p, count_out, count_hold
                
                # --- 신규 셀 기입 (D11 ~ D15) ---
                # B11~B15에 있는 값을 기준으로 E열에서 개수를 찾아 D열에 기입
                for row_num in range(11, 16):
                    search_key = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                    # 찾고자 하는 값이 E열에 있으면 그 개수를, 없으면 0을 입력
                    match_count = e_column_counts.get(search_key, 0)
                    ws[f'D{row_num}'] = match_count
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download_name)
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
