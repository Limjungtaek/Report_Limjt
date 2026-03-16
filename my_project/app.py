import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import io

st.title("엑셀 데이터 연산 및 다운로드 서비스")

uploaded_file = st.file_uploader("A파일(엑셀)을 업로드하세요", type=["xlsx"])

if uploaded_file:
    with st.spinner('파일을 연산 중입니다...'):
        try:
            # 1. 첫 번째 시트 읽기
            # header=None으로 설정하여 행/열 번호로 직접 접근
            df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # 2. I열(인덱스 8)의 6행(인덱스 5)부터 10000행(인덱스 9999)까지 추출
            # .iloc[5:10000, 8]
            target_data = df.iloc[5:10000, 8]
            
            # 3. COUNTA와 동일하게 동작 (비어 있지 않은 셀 개수 카운트)
            result = target_data.count()
            
            # 4. 템플릿 로드
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # B7 셀에 카운트 결과 기록
                ws['B7'] = result
                
                # 5. 메모리에서 엑셀 파일 생성
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                # 6. 다운로드 버튼
                download_name = f"{os.path.splitext(uploaded_file.name)[0]}_Report.xlsx"
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
