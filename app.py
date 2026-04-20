import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import os
import urllib.request
from datetime import datetime

# 1. 폰트 자동 다운로드 함수 추가
def download_font():
    font_file = "NanumGothic.ttf"
    if not os.path.exists(font_file):
        with st.spinner("최초 1회 한글 폰트를 설정 중입니다..."):
            # 구글 폰트 깃허브에서 나눔고딕 TTF 직접 다운로드
            url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
            urllib.request.urlretrieve(url, font_file)

# 2. 한글 출력을 위한 PDF 클래스
class TransactionPDF(FPDF):
    def __init__(self):
        super().__init__()
        # 다운로드된 폰트 파일 연결
        self.add_font("NanumGothic", "", "NanumGothic.ttf")
        self.set_font("NanumGothic", "", 10)

    def header(self):
        self.set_font("NanumGothic", "", 15)
        self.cell(0, 10, '거래명세서', 0, 1, 'C')
        self.ln(5)

def create_pdf(center_name, data_rows):
    pdf = TransactionPDF()
    pdf.add_page()
    
    # 공급자 / 공급받는자 기본 정보
    pdf.set_font("NanumGothic", "", 11)
    pdf.cell(0, 8, f"공급받는자: {center_name}", 0, 1, 'L')
    pdf.cell(0, 8, f"배송일자: {datetime.now().strftime('%Y-%m-%d')}", 0, 1, 'L')
    pdf.ln(5)
    
    # 테이블 헤더 설정
    pdf.set_font("NanumGothic", "", 10)
    headers = ['제품명', '수량(Box/낱개)', '단가', '공급가액']
    col_widths = [90, 30, 35, 35]
    
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
    pdf.ln()
    
    # 데이터 행 추가
    for row in data_rows:
        pdf.cell(col_widths[0], 10, str(row['Product']), 1, 0, 'L')
        pdf.cell(col_widths[1], 10, str(row['Qty']), 1, 0, 'C')
        pdf.cell(col_widths[2], 10, str(row['Price']), 1, 0, 'R')
        pdf.cell(col_widths[3], 10, str(row['Total']), 1, 0, 'R')
        pdf.ln()
        
    return bytes(pdf.output())

def main():
    st.set_page_config(page_title="명세서 PDF 자동화", layout="wide")
    
    # 앱 실행 시 폰트 확인 및 자동 다운로드
    download_font()

    st.title("📦 지점별 거래명세서 PDF 자동 생성기")
    st.write("엑셀 파일을 업로드하면 2번째 시트의 데이터를 바탕으로 지점별 명세서를 일괄 생성합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일 업로드 (.xlsx)", type=['xlsx'])

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            if len(sheet_names) < 2:
                st.error("엑셀 파일에 시트가 2개 이상 존재해야 합니다.")
                return

            st.success(f"파일을 성공적으로 읽었습니다. (참조 시트: {sheet_names[1]})")
            
            # 두 번째 시트 로드
            df_matrix = pd.read_excel(xls, sheet_name=sheet_names[1], header=2) 
            
            # 단가 정보 (필요시 실제 단가로 수정)
            price_dict = {
                "멘소래담 로션 75ml": 3780,
                "멘소래담 로션 100ml": 4590,
                "멘소래담 로션 450ml": 13950
            }

            if st.button("지점별 PDF 생성 및 압축 다운로드"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    center_columns = [col for col in df_matrix.columns if col not in ['제품명', '규격', 'Total'] and not pd.isna(col)]
                    
                    for center in center_columns:
                        data_rows = []
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        for index, row in df_matrix.iterrows():
                            product_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            if pd.notna(qty) and str(qty).replace('.0','').isdigit() and int(float(qty)) > 0:
                                if is_bonus:
                                    unit_price = 0
                                    product_display_name = f"{product_name} (할증분)"
                                else:
                                    unit_price = price_dict.get(product_name, 0)
                                    product_display_name = product_name
                                    
                                total_amount = int(float(qty)) * unit_price
                                
                                data_rows.append({
                                    'Product': product_display_name,
                                    'Qty': int(float(qty)),
                                    'Price': format(unit_price, ','),
                                    'Total': format(total_amount, ',')
                                })
                        
                        if data_rows:
                            safe_center_name = str(center).replace('\n', ' ').replace('/', '_').strip()
                            pdf_bytes = create_pdf(safe_center_name, data_rows)
                            zip_file.writestr(f"거래명세서_{safe_center_name}.pdf", pdf_bytes)
                
                st.download_button(
                    label="전체 명세서 PDF 다운로드 (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"거래명세서_일괄생성_{datetime.now().strftime('%m%d')}.zip",
                    mime="application/zip"
                )
                
                st.balloons()
                st.success("PDF 변환 및 압축이 완벽하게 처리되었습니다!")

        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
