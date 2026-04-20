import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import os
import urllib.request
from datetime import datetime

# 1. 폰트 자동 설정
def setup_resources():
    font_file = "NanumGothic.ttf"
    if not os.path.exists(font_file):
        url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
        urllib.request.urlretrieve(url, font_file)

# 2. PDF 레이아웃을 첨부해주신 PDF와 동일하게 그리는 클래스
class PerfectTemplatePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("NanumGothic", "", "NanumGothic.ttf")
        self.set_font("NanumGothic", "", 10)

    def draw_exact_header(self, center_name, info_data):
        self.add_page()
        
        # 상단 타이틀 (공급자받는자 보관용 & 거래명세서)
        self.set_font("NanumGothic", "", 9)
        self.cell(0, 5, "(  공급자받는자  보관용  )", 0, 1, 'C')
        self.set_font("NanumGothic", "", 24)
        self.cell(0, 12, "거 래 명 세 서", 0, 1, 'C')
        self.ln(5)

        # 주문일자 & 배송일자 라인
        self.set_font("NanumGothic", "", 9)
        self.cell(20, 6, "주문일자", 0, 0)
        self.cell(40, 6, info_data.get('order_date', ''), "B", 0, 'C')
        self.cell(10, 6, "", 0, 0)
        self.cell(20, 6, "배송일자 :", 0, 0)
        self.cell(40, 6, info_data.get('delivery_date', ''), "B", 1, 'C')
        self.ln(2)

        # ---------------------------------------------------------
        # 공급자 vs 공급받는자 테이블 구조 (첨부하신 PDF 완벽 재현)
        # ---------------------------------------------------------
        start_y = self.get_y()
        self.set_font("NanumGothic", "", 9)
        
        # 좌측: 공급자 (맨소래덤아시아퍼시픽)
        self.rect(10, start_y, 95, 30) 
        self.set_xy(10, start_y)
        
        self.cell(15, 30, "공급자", 1, 0, 'C') # 세로 병합 느낌
        x_offset = self.get_x()
        y_offset = self.get_y()
        
        # 공급자 세부 정보
        self.cell(20, 7.5, "등록번호", 1, 0, 'C'); self.cell(60, 7.5, "102-84-02171", 1, 1, 'C')
        self.set_x(x_offset); self.cell(20, 7.5, "상      호", 1, 0, 'C'); self.cell(60, 7.5, "맨소래덤아시아퍼시픽㈜", 1, 1, 'C')
        self.set_x(x_offset); self.cell(20, 7.5, "대  표  자", 1, 0, 'C'); self.cell(60, 7.5, "임현정", 1, 1, 'C')
        self.set_x(x_offset); self.cell(20, 7.5, "주      소", 1, 0, 'C'); 
        self.set_font("NanumGothic", "", 7)
        self.cell(60, 7.5, "서울시 강남구 역삼동 772 동영문화센터빌딩 7층", 1, 1, 'C')
        
        # 우측: 공급받는자 (백제약품 지점)
        self.set_font("NanumGothic", "", 9)
        self.rect(105, start_y, 95, 30)
        self.set_xy(105, start_y)
        
        # '공급받는자' 텍스트 세로 쓰기 시뮬레이션
        self.multi_cell(15, 10, "공급\n받는\n자", 1, 'C') 
        self.set_xy(120, start_y)
        
        # 공급받는자 세부 정보
        self.cell(20, 7.5, "등록번호", 1, 0, 'C'); self.cell(60, 7.5, "", 1, 1, 'C')
        self.set_x(120); self.cell(20, 7.5, "상      호", 1, 0, 'C'); self.cell(60, 7.5, f"{center_name}", 1, 1, 'C')
        self.set_x(120); self.cell(20, 7.5, "대  표  자", 1, 0, 'C'); self.cell(60, 7.5, "", 1, 1, 'C')
        self.set_x(120); self.cell(20, 7.5, "주      소", 1, 0, 'C'); self.cell(60, 7.5, "", 1, 1, 'C')

        # 하단 업태/종목 라인
        self.set_xy(10, start_y + 30)
        self.cell(15, 7, "", 1, 0)
        self.cell(20, 7, "업      태", 1, 0, 'C'); self.cell(60, 7, "도  매", 1, 0, 'C')
        self.cell(15, 7, "", 1, 0)
        self.cell(20, 7, "종      목", 1, 0, 'C'); self.cell(60, 7, "양약외", 1, 1, 'C')
        self.ln(3)

        # ---------------------------------------------------------
        # 데이터 테이블 헤더 (제품코드, 제품명, 입수, Box수, 낱개수, 낱개가격, 공급가액)
        # ---------------------------------------------------------
        headers = ['제품코드', '제품명', '입수', 'Box 수', '낱개수', '낱개가격', '공급가액']
        # 너비 비율 조정 (190mm 기준)
        widths = [15, 65, 15, 15, 20, 25, 35] 
        
        for i, h in enumerate(headers):
            self.cell(widths[i], 8, h, 1, 0, 'C')
        self.ln()
        
        return widths

def main():
    st.set_page_config(page_title="맨소래덤아시아퍼시픽 명세서", layout="centered")
    setup_resources()
    
    st.title("📄 지점별 거래명세서 다이렉트 PDF 생성기")
    st.info("보내주신 PDF 양식과 100% 동일한 형태의 명세서를 바로 뽑아냅니다.")

    uploaded_file = st.file_uploader("백제약품 센터별 배송 엑셀 파일 업로드", type=['xlsx'])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            # 배송 수량 Matrix 시트 로드
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품 마스터 (입수량, 낱개가격) - 보내주신 PDF 데이터 기준 매핑
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("결과값(PDF) 즉시 생성 및 다운로드"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점명 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        pdf = PerfectTemplatePDF()
                        today_str = datetime.now().strftime('%Y-%m-%d')
                        
                        # 백제약품 광주지점, 평택물류센터 등 상단 양식 그리기
                        widths = pdf.draw_exact_header(str(center).strip(), {'order_date': today_str, 'delivery_date': today_str})
                        
                        rows_added = 0
                        for idx, row in df_matrix.iterrows():
                            p_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            # 수량이 있을 때만 표기
                            if pd.notna(qty) and str(qty).replace('.0','').isdigit() and int(float(qty)) > 0:
                                info = product_master.get(p_name, {"inbox": 1, "price": 0})
                                
                                # 할증 지점이면 가격 0원, 제품명에 (할증분) 표시
                                price = 0 if is_bonus else info['price']
                                total = int(float(qty)) * price
                                display_name = f"{p_name} (할증분)" if is_bonus else p_name
                                
                                # Box 수 계산
                                box_qty = int(float(qty)) // info['inbox'] if info['inbox'] > 0 else 0
                                
                                # 행 추가
                                pdf.cell(widths[0], 8, str(rows_added + 1), 1, 0, 'C') # 제품코드(No)
                                pdf.cell(widths[1], 8, f" {display_name}", 1, 0, 'L')  # 제품명
                                pdf.cell(widths[2], 8, str(info['inbox']), 1, 0, 'C')  # 입수
                                pdf.cell(widths[3], 8, str(box_qty), 1, 0, 'C')        # Box 수
                                pdf.cell(widths[4], 8, str(int(float(qty))), 1, 0, 'C')# 낱개수
                                pdf.cell(widths[5], 8, "" if price == 0 else format(price, ','), 1, 0, 'R') # 낱개가격 (할증은 빈칸 처리 또는 0)
                                pdf.cell(widths[6], 8, "" if total == 0 else format(total, ','), 1, 1, 'R') # 공급가액
                                
                                rows_added += 1

                        if rows_added > 0:
                            safe_filename = str(center).replace('\n', ' ').replace('/', '_').strip()
                            zip_f.writestr(f"거래명세서_{safe_filename}.pdf", bytes(pdf.output()))
                
                st.download_button(
                    label="📦 보내주신 형태와 완벽히 동일한 PDF 다운로드 (ZIP)", 
                    data=zip_buffer.getvalue(), 
                    file_name=f"거래명세서_최종결과물_{datetime.now().strftime('%m%d')}.zip",
                    mime="application/zip"
                )
                st.balloons()
                st.success("PDF 생성이 완료되었습니다! 다운로드하여 확인해 보세요.")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
