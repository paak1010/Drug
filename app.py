import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
from datetime import datetime

# 필요 라이브러리: pip install streamlit pandas openpyxl

def main():
    st.set_page_config(page_title="명세서 엑셀 템플릿 자동화", layout="wide")
    
    st.title("🎯 원본 엑셀 양식 100% 유지 명세서 자동 생성기")
    st.info("첫 번째 시트의 양식(서식, 병합, 테두리)을 그대로 유지한 채, 지점별 엑셀 파일을 일괄 생성합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx'])

    if uploaded_file:
        try:
            # pandas로 데이터 시트(두 번째 시트)만 읽어오기
            xls = pd.ExcelFile(uploaded_file)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 단가 및 입수량 마스터
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("지점별 명세서(Excel) 일괄 생성 및 다운로드"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점명(Column) 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        data_rows = []
                        for idx, row in df_matrix.iterrows():
                            p_name = str(row['제품명']).strip()
                            qty = row[center]
                            
                            if pd.notna(qty) and str(qty).replace('.0','').isdigit() and int(float(qty)) > 0:
                                info = product_master.get(p_name, {"inbox": 1, "price": 0})
                                price = 0 if is_bonus else info['price']
                                total = int(float(qty)) * price
                                display_name = f"{p_name} (할증분)" if is_bonus else p_name
                                
                                data_rows.append({
                                    'Product': display_name,
                                    'InBox': info['inbox'],
                                    'BoxQty': int(float(qty)) // info['inbox'] if info['inbox'] > 0 else 0,
                                    'Qty': int(float(qty)),
                                    'Price': price,
                                    'Total': total
                                })

                        # ★ 핵심: 해당 지점에 물량이 있으면 엑셀 원본 복사 후 데이터 덮어쓰기
                        if data_rows:
                            # 원본 엑셀 파일을 그대로 메모리에 로드 (서식 100% 유지)
                            wb = openpyxl.load_workbook(uploaded_file)
                            ws = wb.worksheets[0] # 첫 번째 시트 (템플릿)

                            # 1. 공급받는자 상호 입력 (보내주신 CSV 기준 I열 10행 부근)
                            # ※ 엑셀 파일에 따라 셀 위치가 다르면 'I10' 부분을 수정하세요 (예: 'H10', 'J10')
                            ws['I10'] = str(center).strip()
                            
                            # 2. 오늘 날짜로 주문/배송일자 덮어쓰기 (필요시)
                            # ws['C8'] = datetime.now().strftime('%Y-%m-%d') # 주문일자 셀
                            # ws['G8'] = datetime.now().strftime('%Y-%m-%d') # 배송일자 셀

                            # 3. 데이터 리스트 입력 (15행부터 시작한다고 가정)
                            start_row = 15
                            
                            for row_idx, item in enumerate(data_rows):
                                current_row = start_row + row_idx
                                
                                # 각 컬럼 위치에 값 넣기 (엑셀 열 알파벳에 맞춰 수정 가능)
                                ws[f'A{current_row}'] = row_idx + 1         # No
                                ws[f'C{current_row}'] = item['Product']     # 제품명
                                ws[f'I{current_row}'] = item['InBox']       # 입수
                                ws[f'K{current_row}'] = item['BoxQty']      # Box수
                                ws[f'L{current_row}'] = item['Qty']         # 낱개수
                                ws[f'M{current_row}'] = item['Price']       # 단가 (숫자)
                                ws[f'N{current_row}'] = item['Total']       # 공급가액 (숫자)

                            # 완성된 엑셀 객체를 바이트로 변환하여 ZIP에 추가
                            excel_io = io.BytesIO()
                            wb.save(excel_io)
                            excel_io.seek(0)
                            
                            safe_filename = str(center).replace('\n', ' ').replace('/', '_').strip()
                            zip_f.writestr(f"거래명세서_{safe_filename}.xlsx", excel_io.read())
                
                st.download_button(
                    label="📦 전체 명세서 엑셀(ZIP) 다운로드", 
                    data=zip_buffer.getvalue(), 
                    file_name=f"거래명세서_엑셀일괄생성_{datetime.now().strftime('%m%d')}.zip",
                    mime="application/zip"
                )
                st.success("원본 양식이 100% 적용된 엑셀 파일 생성이 완료되었습니다!")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
