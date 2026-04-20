import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import os
import urllib.request
from datetime import datetime

# 1. 한글 폰트 자동 설정
def setup_resources():
    font_file = "NanumGothic.ttf"
    if not os.path.exists(font_file):
        url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
        urllib.request.urlretrieve(url, font_file)

# 2. 첫 번째 시트 양식을 그대로 그리는 클래스
class PerfectTemplatePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("NanumGothic", "", "NanumGothic.ttf")
        self.set_font("NanumGothic", "", 10)

    def draw_excel_header(self, center_name, info_data):
        self.add_page()
        # 타이틀
        self.set_font("NanumGothic", "", 22)
        self.cell(0, 15, "거  래  명  세  서", 0, 1, 'C')
        self.set_font("NanumGothic", "", 9)
        self.cell(0, 5, "(공급자받는자 보관용)", 0, 1, 'C')
        self.ln(5)

        # 날짜 정보
        self.set_font("NanumGothic", "", 9)
        self.cell(25, 8, "주문일자:", 0, 0)
        self.cell(60, 8, info_data.get('order_date', ''), "B", 0)
        self.cell(15, 8, "", 0, 0)
        self.cell(25, 8, "배송일자:", 0, 0)
        self.cell(0, 8, info_data.get('delivery_date', ''), "B", 1)
        self.ln(3)

        # 공급자 및 공급받는자 박스
        curr_y = self.get_y()
        
        # [왼쪽: 공급자]
        self.rect(10, curr_y, 90, 40)
        self.set_xy(12, curr_y + 2)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "등 록 번 호", 0, 0); self.set_font("NanumGothic", "", 10); self.cell(0, 6, ": 102-84-02171", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "상      호", 0, 0); self.set_font("NanumGothic", "", 9); self.cell(0, 6, ": 맨소래덤아시아퍼시픽㈜", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "대  표  자", 0, 0); self.set_font("NanumGothic", "", 9); self.cell(0, 6, ": 임현정", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "주      소", 0, 0); self.set_font("NanumGothic", "", 7); self.multi_cell(70, 4, ": 서울시 강남구 역삼동 772\n  동영문화센터빌딩 7층")

        # [오른쪽: 공급받는자]
        self.rect(105, curr_y, 95, 40)
        self.set_xy(107, curr_y + 2)
        self.set_font("NanumGothic", "", 8); self.cell(15, 6, "상      호", 0, 0); self.set_font("NanumGothic", "", 11); self.cell(0, 6, f": {center_name}", 0, 1)
        
        self.set_xy(10, curr_y + 45)
        
        # 테이블 헤더
        self.set_font("NanumGothic", "", 9)
        headers = ['No', '제품명', '입수', 'Box수', '낱개수', '단가', '공급가액']
        widths = [10, 80, 15, 15, 20, 25, 25]
        for i, h in enumerate(headers):
            self.cell(widths[i], 10, h, 1, 0, 'C')
        self.ln()
        return widths

def main():
    st.set_page_config(page_title="명세서 PDF 변환기", layout="centered")
    setup_resources()
    
    st.title("📄 지점별 거래명세서 일괄 생성기")
    st.write("엑셀 파일을 업로드하면 지점별 명세서 PDF를 생성하여 압축 파일로 다운로드합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx'])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            # 두 번째 시트 로드 (헤더 2줄 스킵)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품 마스터 정보 (입수량, 단가)
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("명세서 PDF 생성 및 다운로드"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점명(Column) 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        pdf = PerfectTemplatePDF()
                        
                        # 상단 날짜 정보 세팅 (오늘 날짜 기준)
                        today_str = datetime.now().strftime('%Y-%m-%d')
                        widths = pdf.draw_excel_header(str(center).strip(), {'order_date': today_str, 'delivery_date': today_str})
                        
                        rows_added = 0
                        
                        # 핵심 로직: 할증 여부 판단
                        is_bonus = '할증' in str(center) or '힐증' in str(center)

                        for idx, row in df_matrix.iterrows():
                            p_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            # 수량이 있는 경우에만 행 추가
                            if pd.notna(qty) and int(float(qty)) > 0:
                                info = product_master.get(p_name, {"inbox": 1, "price": 0})
                                
                                # 할증 지점은 단가 0원
                                price = 0 if is_bonus else info['price']
                                total = int(float(qty)) * price
                                
                                # 제품명에 할증 표기 추가
                                display_name = f"{p_name} (할증분)" if is_bonus else p_name
                                
                                # PDF 표에 데이터 입력
                                pdf.cell(widths[0], 8, str(rows_added + 1), 1, 0, 'C')
                                pdf.cell(widths[1], 8, f" {display_name}", 1, 0, 'L')
                                pdf.cell(widths[2], 8, str(info['inbox']), 1, 0, 'C')
                                pdf.cell(widths[3], 8, str(int(float(qty)) // info['inbox']), 1, 0, 'C') # Box 수
                                pdf.cell(widths[4], 8, str(int(float(qty))), 1, 0, 'C') # 낱개 수
                                pdf.cell(widths[5], 8, format(price, ','), 1, 0, 'R')
                                pdf.cell(widths[6], 8, format(total, ','), 1, 1, 'R')
                                rows_added += 1

                        # 해당 지점에 내역이 하나라도 있으면 ZIP에 추가
                        if rows_added > 0:
                            safe_filename = str(center).replace('\n', ' ').replace('/', '_').strip()
                            zip_f.writestr(f"거래명세서_{safe_filename}.pdf", bytes(pdf.output()))
                
                st.download_button(
                    label="📦 전체 명세서 압축파일(ZIP) 다운로드", 
                    data=zip_buffer.getvalue(), 
                    file_name=f"거래명세서_일괄생성_{datetime.now().strftime('%m%d')}.zip",
                    mime="application/zip"
                )
                st.success("명세서 생성이 완료되었습니다!")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
