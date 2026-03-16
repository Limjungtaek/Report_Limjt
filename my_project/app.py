import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

# 1. 파일 업로드
uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    # 2. 로딩 시각화
    with st.spinner('파일을 연산하고 보고서를 생성 중입니다...'):
        try:
            # 3. 인풋파일(A파일) 읽기 (첫 번째 시트)
            # header=None을 설정하여 행 번호를 우리가 직접 지정할 수 있게 합니다.
            # I열은 9번째 열이므로 index 8을 사용합니다.
            df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # I열(인덱스 8)의 6행(인덱스 5)부터 끝까지 데이터 추출
            i_column_data = df.iloc[5:, 8]
            
            # 공란이 아닌 수들만 연산 (숫자 변환 후 합계 계산)
            numeric_data = pd.to_numeric(i_column_data, errors='coerce')
            result = numeric_data.sum()
            
            # 4. 템플릿 로드
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # B7 셀에 결과값 기록
                ws['B7'] = result
                
                # 5. 메모리에서 엑셀 파일 생성
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                # 6. 다운로드 버튼
                base_name = os.path.splitext(uploaded_file.name)[0]
                download_name = f"{base_name}_Report.xlsx"
                
                st.download_button(
                    label="결과 파일 다운로드",
                    data=output,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("템플릿 파일을 찾을 수 없습니다.")
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
