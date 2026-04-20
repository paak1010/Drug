import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
from datetime import datetime

# 한글 출력을 위한 PDF 클래스 (fpdf2 사용)
class TransactionPDF(FPDF):
    def __init__(self):
        super().__init__()
        # 한글 폰트 추가 (폰트 파일명이 정확히 일치해야 하며, 같은 경로에 있어야 합니다)
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
        
    # fpdf2에서는 output()을 호출하면 기본적으로 byte 형식으로 안전하게 반환됩니다.
    return bytes(pdf.output())

def main():
    st.set_page_config(page_title="명세서 PDF 자동화", layout="wide")
    st.title("📦 지점별 거래명세서 PDF 자동 생성기")
    st.write("엑셀 파일을 업로드하면 2번째 시트의 데이터를 바탕으로 지점별 명세서를 일괄 생성합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일 업로드 (.xlsx)", type=['xlsx'])

    if uploaded_file is not None:
        try:
            # 엑셀 파일의 모든 시트 읽기
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            if len(sheet_names) < 2:
                st.error("엑셀 파일에 시트가 2개 이상 존재해야 합니다.")
                return

            st.success(f"파일을 성공적으로 읽었습니다. (참조 시트: {sheet_names[1]})")
            
            # 두 번째 시트 로드 (센터별 물량 데이터, 상단 2줄 헤더 스킵)
            df_matrix = pd.read_excel(xls, sheet_name=sheet_names[1], header=2) 
            
            # 단가 정보 딕셔너리 (필요에 따라 실제 단가로 수정하세요)
            price_dict = {
                "멘소래담 로션 75ml": 3780,
                "멘소래담 로션 100ml": 4590,
                "멘소래담 로션 450ml": 13950
            }

            if st.button("지점별 PDF 생성 및 압축 다운로드"):
                # ZIP 파일을 메모리에 생성
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    # '제품명', '규격', 'Total' 등을 제외한 실제 지점명(열)만 추출
                    center_columns = [col for col in df_matrix.columns if col not in ['제품명', '규격', 'Total'] and not pd.isna(col)]
                    
                    for center in center_columns:
                        data_rows = []
                        
                        # 핵심 로직: 열 이름에 '할증'(또는 오타인 '힐증')이 포함되어 있는지 확인
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        for index, row in df_matrix.iterrows():
                            product_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            # 수량이 존재하고 0보다 큰 경우에만 명세서에 추가
                            if pd.notna(qty) and str(qty).replace('.0','').isdigit() and int(float(qty)) > 0:
                                
                                # 할증 지점인 경우 단가 및 공급가액 0원 처리
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
                        
                        # 해당 지점에 배송할 물품이 1개라도 있다면 PDF 생성 후 ZIP에 추가
                        if data_rows:
                            # 안전한 파일명을 위해 특수문자 치환
                            safe_center_name = str(center).replace('\n', ' ').replace('/', '_').strip()
                            pdf_bytes = create_pdf(safe_center_name, data_rows)
                            zip_file.writestr(f"거래명세서_{safe_center_name}.pdf", pdf_bytes)
                
                # 다운로드 버튼 활성화
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
