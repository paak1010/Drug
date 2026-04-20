import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import os
import urllib.request
from datetime import datetime

# 1. 한글 폰트 자동 설정
def download_font():
    font_file = "NanumGothic.ttf"
    if not os.path.exists(font_file):
        url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
        urllib.request.urlretrieve(url, font_file)

# 2. 명세서 양식을 그대로 그리는 클래스
class TransactionPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("NanumGothic", "", "NanumGothic.ttf")
        self.set_font("NanumGothic", "", 10)

    def draw_template(self, center_name):
        self.add_page()
        # 타이틀
        self.set_font("NanumGothic", "", 20)
        self.cell(0, 15, "거 래 명 세 서", 0, 1, 'C')
        
        self.set_font("NanumGothic", "", 9)
        self.cell(100, 7, f"주문일자: {datetime.now().strftime('%Y-%m-%d')}", 0, 0)
        self.cell(0, 7, f"배송일자: {datetime.now().strftime('%Y-%m-%d')}", 0, 1, 'R')
        
        # 상단 정보 박스 (공급자 vs 공급받는자)
        start_y = self.get_y()
        
        # 왼쪽: 공급자 (멘소래담 정보 고정)
        self.rect(10, start_y, 90, 35) 
        self.set_xy(12, start_y + 2)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "등 록 번 호", 0, 0)
        self.set_font("NanumGothic", "", 10)
        self.cell(0, 6, ": 102-84-02171", 0, 1)
        
        self.set_x(12)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "상      호", 0, 0)
        self.set_font("NanumGothic", "", 9)
        self.cell(0, 6, ": 맨소래덤아시아퍼시픽(주)", 0, 1)
        
        self.set_x(12)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "대  표  자", 0, 0)
        self.set_font("NanumGothic", "", 9)
        self.cell(0, 6, ": 임현정", 0, 1)
        
        self.set_x(12)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "주      소", 0, 0)
        self.set_font("NanumGothic", "", 7)
        self.multi_cell(70, 4, ": 서울시 강남구 역삼동 772\n  동영문화센터빌딩 7층")
        
        # 오른쪽: 공급받는자 (지점명 동적 변경)
        self.rect(105, start_y, 95, 35)
        self.set_xy(107, start_y + 2)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "상      호", 0, 0)
        self.set_font("NanumGothic", "", 11)
        self.cell(0, 8, f": {center_name}", 0, 1)
        
        self.set_xy(10, start_y + 40)
        # 테이블 헤더 (첫 번째 시트 컬럼명 기준)
        self.set_font("NanumGothic", "", 9)
        headers = ['No', '제품명', '입수', 'Box수', '낱개수', '단가', '공급가액']
        widths = [10, 80, 15, 15, 20, 25, 25]
        
        for i, h in enumerate(headers):
            self.cell(widths[i], 10, h, 1, 0, 'C')
        self.ln()
        return widths

def create_formatted_pdf(center_name, data_rows):
    pdf = TransactionPDF()
    widths = pdf.draw_template(center_name)
    
    total_sum = 0
    for i, row in enumerate(data_rows):
        pdf.cell(widths[0], 8, str(i+1), 1, 0, 'C')
        pdf.cell(widths[1], 8, f" {row['Product']}", 1, 0, 'L')
        pdf.cell(widths[2], 8, str(row['InBox']), 1, 0, 'C')
        pdf.cell(widths[3], 8, str(row['BoxQty']), 1, 0, 'C')
        pdf.cell(widths[4], 8, str(row['Qty']), 1, 0, 'C')
        pdf.cell(widths[5], 8, row['Price'], 1, 0, 'R')
        pdf.cell(widths[6], 8, row['Total'], 1, 0, 'R')
        pdf.ln()
        
        # 총 합계 계산을 위해 쉼표 제거 후 더하기
        val = int(row['Total'].replace(',', ''))
        total_sum += val

    # 하단 합계란
    pdf.cell(sum(widths[:6]), 10, "합 계 금 액 (VAT 포함 생략)", 1, 0, 'C')
    pdf.cell(widths[6], 10, format(total_sum, ','), 1, 1, 'R')
    
    return bytes(pdf.output())

def main():
    st.set_page_config(page_title="거래명세서 자동화 시스템", layout="wide")
    download_font()
    
    st.title("📄 양식 정밀 재현 거래명세서 생성기")
    st.info("첫 번째 시트의 '양식'을 참조하고, 두 번째 시트의 '데이터'를 읽어 PDF를 만듭니다.")

    uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=['xlsx'])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            # 첫 번째 시트: 양식 샘플 (여기서 업체 정보 등을 가져올 수도 있지만 일단 고정으로 구현)
            # 두 번째 시트: 배송 데이터 Matrix
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품별 단가 및 입수량 설정 (첫 번째 시트 데이터 기반)
            # 548일 이상 유통기한 기준 등 필요한 로직이 있다면 여기서 필터링 가능합니다.
            product_info = {
                "멘소래담 로션 75ml": {"price": 3780, "inbox": 72},
                "멘소래담 로션 100ml": {"price": 4590, "inbox": 72},
                "멘소래담 로션 450ml": {"price": 13950, "inbox": 20}
            }

            if st.button("지점별 명세서 일괄 생성"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점 열(Column) 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        rows = []
                        # 할증 열 여부 체크 (가격 0원 로직)
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        for _, r in df_matrix.iterrows():
                            p_name = str(r['제품명']).strip()
                            qty = r[center]
                            
                            if pd.notna(qty) and int(float(qty)) > 0:
                                info = product_info.get(p_name, {"price": 0, "inbox": 0})
                                
                                # 할증이면 가격 0원, 아니면 정상가
                                price = 0 if is_bonus else info['price']
                                total = int(float(qty)) * price
                                
                                rows.append({
                                    'Product': f"{p_name}{' (할증분)' if is_bonus else ''}",
                                    'InBox': info['inbox'],
                                    'BoxQty': int(float(qty)) // info['inbox'] if info['inbox'] > 0 else 0,
                                    'Qty': int(float(qty)),
                                    'Price': format(price, ','),
                                    'Total': format(total, ',')
                                })
                        
                        if rows:
                            pdf_data = create_formatted_pdf(str(center).strip(), rows)
                            zip_f.writestr(f"거래명세서_{str(center).strip()}.pdf", pdf_data)
                
                st.download_button("전체 PDF 다운로드 (ZIP)", zip_buffer.getvalue(), "명세서_일괄생성.zip", "application/zip")
                st.success("양식 적용이 완료되었습니다.")

        except Exception as e:
            st.error(f"오류 발생: {e}")

if __name__ == "__main__":
    main()
