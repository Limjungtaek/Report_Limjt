import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('7행 요약과 상세 내역을 정렬하여 보고서를 생성 중입니다...'):
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
            data_s2 = df_sheet2.iloc[:, [3, 14, 15]].copy()
            data_s2.columns = ['item_name', 'sum_value', 'status']
            data_s2['item_name'] = data_s2['item_name'].astype(str).str.strip()
            data_s2['status'] = data_s2['status'].astype(str).str.strip()
            data_s2['sum_value'] = pd.to_numeric(data_s2['sum_value'], errors='coerce').fillna(0)
            
            s2_total_counts = data_s2['item_name'].value_counts()
            s2_out_counts = data_s2[data_s2['status'] == '출고']['item_name'].value_counts()
            s2_hold_counts = data_s2[data_s2['status'] == '보유']['item_name'].value_counts()
            s2_sum_values = data_s2.groupby('item_name')['sum_value'].sum()

            # 2. 템플릿 처리
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # --- 상단 요약 (7행) 위치 수정 ---
                # 시트1 기준
                ws['B7'] = df_sheet1.iloc[5:10000, 3].count()
                ws['D7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '정상').sum()
                ws['F7'] = (df_sheet1.iloc[5:10000, 3].astype(str).str.strip() == '폐업').sum()
                
                # 시트2 기준 (기존 J, L, N에서 I, K, M으로 한 칸씩 이동)
                ws['I7'] = df_sheet2.iloc[4:10000, 15].count()                   # 전체 (기존 J)
                ws['K7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '출고').sum() # 출고 (기존 L)
                ws['M7'] = (df_sheet2.iloc[4:10000, 15].astype(str).str.strip() == '보유').sum() # 보유 (기존 N)
                
                # --- [11~16행 상세 기입] ---
                for row_num in range(11, 17):
                    # 시트1 기반 (B열 기준 -> D, E, F 채우기)
                    key_b = str(ws[f'B{row_num}'].value).strip() if ws[f'B{row_num}'].value else ""
                    if key_b:
                        ws[f'D{row_num}'] = s1_total_counts.get(key_b, 0)
                        ws[f'E{row_num}'] = s1_normal_counts.get(key_b, 0)
                        ws[f'F{row_num}'] = s1_closed_counts.get(key_b, 0)
                    
                    # 시트2 기반 (I열 기준 -> J, K, L, M 채우기)
                    key_i = str(ws[f'I{row_num}'].value).strip() if ws[f'I{row_num}'].value else ""
                    if key_i:
                        ws[f'J{row_num}'] = s2_total_counts.get(key_i, 0) # 전체
                        ws[f'K{row_num}'] = s2_out_counts.get(key_i, 0)   # 출고
                        ws[f'L{row_num}'] = s2_hold_counts.get(key_i, 0)  # 보유
                        ws[f'M{row_num}'] = s2_sum_values.get(key_i, 0)   # 합계
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
                st.download_button("결과 파일 다운로드", output, download_name)
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("팁: 만약 'MergedCell' 에러가 계속된다면, 데이터가 입력되는 셀(I7, K7, M7 등)이 엑셀에서 다른 셀과 병합되어 있는지 확인해 보세요.")
