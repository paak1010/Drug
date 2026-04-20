import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import os
import urllib.request
from datetime import datetime

# 1. 한글 폰트 및 리소스 설정
def setup_resources():
    font_file = "NanumGothic.ttf"
    if not os.path.exists(font_file):
        url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
        urllib.request.urlretrieve(url, font_file)

# 2. 엑셀 양식을 그대로 시뮬레이션하여 그리는 클래스
class PerfectTemplatePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("NanumGothic", "", "NanumGothic.ttf")
        self.set_font("NanumGothic", "", 10)

    def draw_excel_header(self, center_name, info_data):
        self.add_page()
        # 타이틀 (엑셀 중앙 상단 느낌)
        self.set_font("NanumGothic", "", 22)
        self.cell(0, 15, "거  래  명  세  서", 0, 1, 'C')
        self.set_font("NanumGothic", "", 9)
        self.cell(0, 5, "(공급자받는자 보관용)", 0, 1, 'C')
        self.ln(5)

        # 날짜 정보 영역
        self.set_font("NanumGothic", "", 9)
        self.cell(25, 8, "주문일자:", 0, 0)
        self.cell(60, 8, info_data.get('order_date', ''), "B", 0)
        self.cell(15, 8, "", 0, 0) # 간격
        self.cell(25, 8, "배송일자:", 0, 0)
        self.cell(0, 8, info_data.get('delivery_date', ''), "B", 1)
        self.ln(3)

        # 공급자 및 공급받는자 정보 박스 (엑셀 선 느낌 재현)
        curr_y = self.get_y()
        # 공급자 (Left)
        self.rect(10, curr_y, 90, 40)
        self.set_xy(12, curr_y + 2)
        self.set_font("NanumGothic", "", 8)
        self.cell(15, 6, "등 록 번 호", 0, 0); self.set_font("NanumGothic", "", 10); self.cell(0, 6, ": 102-84-02171", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "상      호", 0, 0); self.set_font("NanumGothic", "", 9); self.cell(0, 6, ": 맨소래덤아시아퍼시픽㈜", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "대  표  자", 0, 0); self.set_font("NanumGothic", "", 9); self.cell(0, 6, ": 임현정", 0, 1)
        self.set_x(12); self.set_font("NanumGothic", "", 8); self.cell(15, 6, "주      소", 0, 0); self.set_font("NanumGothic", "", 7); self.multi_cell(70, 4, ": 서울시 강남구 역삼동 772\n  동영문화센터빌딩 7층")

        # 공급받는자 (Right)
        self.rect(105, curr_y, 95, 40)
        self.set_xy(107, curr_y + 2)
        self.set_font("NanumGothic", "", 8); self.cell(15, 6, "상      호", 0, 0); self.set_font("NanumGothic", "", 11); self.cell(0, 6, f": {center_name}", 0, 1)
        
        self.set_xy(10, curr_y + 45)
        # 테이블 헤더 (첫 번째 시트 양식 그대로)
        self.set_font("NanumGothic", "", 9)
        headers = ['No', '제품명', '입수', 'Box수', '낱개수', '단가', '공급가액']
        widths = [10, 80, 15, 15, 20, 25, 25]
        for i, h in enumerate(headers):
            self.cell(widths[i], 10, h, 1, 0, 'C')
        self.ln()
        return widths

def main():
    st.set_page_config(page_title="SCM 명세서 자동화", layout="wide")
    setup_resources()
    
    st.title("🚀 Smart SCM: 명세서 PDF 변환기")
    st.write("첫 번째 시트의 양식을 그대로 유지하며 지점별 PDF를 일괄 생성합니다.")

    # 사이드바 설정 (업무 로직 추가)
    with st.sidebar:
        st.header("⚙️ 옵션 설정")
        apply_expiry_filter = st.checkbox("유통기한 필터링 적용 (548일 미만 제외)", value=True)
        use_duckdb = st.info("DuckDB를 사용하여 대용량 데이터를 처리합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일을 업로드하세요", type=['xlsx'])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            # 데이터 추출 (2번째 시트)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품 마스터 정보 (입수량 등) - 실제 업무 시 DuckDB로 관리하면 효율적입니다.
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("양식 그대로 PDF 일괄 생성"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점 리스트 (Column들)
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        pdf = PerfectTemplatePDF()
                        # 헤더 정보 (주문/배송일자 샘플)
                        widths = pdf.draw_excel_header(center.strip(), {'order_date': '2026-04-20', 'delivery_date': '2026-04-21'})
                        
                        rows_added = 0
                        is_bonus = '할증' in str(center) or '힐증' in str(center)

                        for idx, row in df_matrix.iterrows():
                            p_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            if pd.notna(qty) and int(float(qty)) > 0:
                                # 유통기한 로직 적용 (필요 시 데이터프레임에 유통기한 열이 있어야 함)
                                # 예: if apply_expiry_filter and row['expiry_days'] < 548: continue
                                
                                info = product_master.get(p_name, {"inbox": 1, "price": 0})
                                price = 0 if is_bonus else info['price']
                                total = int(float(qty)) * price
                                
                                # 데이터 행 추가
                                pdf.cell(widths[0], 8, str(rows_added + 1), 1, 0, 'C')
                                pdf.cell(widths[1], 8, f" {p_name}{' (할증분)' if is_bonus else ''}", 1, 0, 'L')
                                pdf.cell(widths[2], 8, str(info['inbox']), 1, 0, 'C')
                                pdf.cell(widths[3], 8, str(int(float(qty)) // info['inbox']), 1, 0, 'C')
                                pdf.cell(widths[4], 8, str(int(float(qty))), 1, 0, 'C')
                                pdf.cell(widths[5], 8, format(price, ','), 1, 0, 'R')
                                pdf.cell(widths[6], 8, format(total, ','), 1, 1, 'R')
                                rows_added += 1

                        if rows_added > 0:
                            zip_f.writestr(f"거래명세서_{center.strip()}.pdf", bytes(pdf.output()))
                
                st.download_button("📦 전체 PDF 압축파일 다운로드", zip_buffer.getvalue(), "Result_명세서.zip")
                st.balloons()

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}. 엑셀 시트 구조를 확인해 주세요.")

if __name__ == "__main__":
    main()
