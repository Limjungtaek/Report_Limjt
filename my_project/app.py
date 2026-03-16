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
            # 1. 첫 번째 시트 읽기 (D열은 인덱스 3)
            df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # 2. D열(인덱스 3)의 6행(인덱스 5)부터 10000행(인덱스 9999)까지 추출
            target_range = df.iloc[5:10000, 3]
            
            # 3. '정상'과 '폐업'인 행의 개수 세기
            # astype(str)로 문자로 변환 후 비교하여 정확도 향상
            count_normal = (target_range.astype(str).str.strip() == '정상').sum()
            count_closed = (target_range.astype(str).str.strip() == '폐업').sum()
            
            # 4. 템플릿 로드
            base_path = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_path, "templates", "template_B.xlsx")
            
            if os.path.exists(template_path):
                wb = load_workbook(template_path)
                ws = wb.active
                
                # B7에 '정상' 개수, F7에 '폐업' 개수 기록
                ws['B7'] = count_normal
                ws['F7'] = count_closed
                
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
