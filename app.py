import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
import duckdb

def main():
    st.set_page_config(page_title="Smart SCM 명세서 자동화", layout="wide")
    
    st.title("🎯 원본 엑셀 양식 100% 유지 명세서 생성기")
    st.write("첫 번째 시트의 원본 서식을 그대로 보존하며, 정확한 셀 위치에 데이터만 변경합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx'])

    if uploaded_file:
        try:
            # 1. 두 번째 시트(데이터 매트릭스) 로드
            xls = pd.ExcelFile(uploaded_file)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품별 단가/입수량 마스터
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("명세서 엑셀 일괄 생성"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점명(Column) 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        # DuckDB를 활용한 고속 데이터 필터링
                        query = f"""
                            SELECT 제품명, "{center}" as qty
                            FROM df_matrix
                            WHERE "{center}" IS NOT NULL AND CAST("{center}" AS DOUBLE) > 0
                        """
                        filtered_df = duckdb.query(query).to_df()
                        
                        data_rows = []
                        for _, row in filtered_df.iterrows():
                            p_name = str(row['제품명']).strip()
                            qty = float(row['qty'])
                            
                            info = product_master.get(p_name, {"inbox": 1, "price": 0})
                            
                            # 할증은 단가 0원 처리
                            price = 0 if is_bonus else info['price']
                            total = int(qty) * price
                            display_name = f"{p_name} (할증분)" if is_bonus else p_name
                            
                            data_rows.append({
                                'Product': display_name,
                                'InBox': info['inbox'],
                                'BoxQty': int(qty) // info['inbox'] if info['inbox'] > 0 else 0,
                                'Qty': int(qty),
                                'Price': price,
                                'Total': total
                            })

                        # 해당 지점에 배송할 데이터가 있는 경우에만 엑셀 수정
                        if data_rows:
                            # 원본 엑셀을 도화지로 로드 (서식 100% 유지)
                            uploaded_file.seek(0) 
                            wb = openpyxl.load_workbook(uploaded_file)
                            ws = wb.worksheets[0] 
                            
                            # 공급받는자 상호 셀 수정
                            ws['I10'] = str(center).strip()
                            
                            # 기존 데이터 영역 초기화 (15행부터 30행까지 잔여 데이터 삭제)
                            for r in range(15, 30):
                                for c in ['A', 'C', 'I', 'J', 'K', 'L', 'M']:
                                    ws[f'{c}{r}'] = ""
                            
                            # 데이터 입력 (정확한 좌표로 수정 완료)
                            start_row = 15
                            
                            for row_idx, item in enumerate(data_rows):
                                curr_row = start_row + row_idx
                                
                                ws[f'A{curr_row}'] = row_idx + 1               # 제품코드 (No)
                                ws[f'C{curr_row}'] = item['Product']           # 제품명
                                ws[f'I{curr_row}'] = item['InBox']             # 입수
                                ws[f'J{curr_row}'] = item['BoxQty']            # Box 수 (좌표 수정)
                                ws[f'K{curr_row}'] = item['Qty']               # 낱개수 (좌표 수정)
                                
                                # 할증일 경우 단가/공급가액 빈칸 처리
                                ws[f'L{curr_row}'] = item['Price'] if item['Price'] > 0 else ""  # 낱개가격 (좌표 수정)
                                ws[f'M{curr_row}'] = item['Total'] if item['Total'] > 0 else ""  # 공급가액 (좌표 수정)

                            # 메모리에 엑셀 임시 저장 및 ZIP 추가
                            excel_io = io.BytesIO()
                            wb.save(excel_io)
                            excel_io.seek(0)
                            
                            safe_center = str(center).replace('\n', ' ').replace('/', '_').strip()
                            zip_f.writestr(f"거래명세서_{safe_center}.xlsx", excel_io.read())
                
                st.download_button(
                    label="📦 서식 유지 엑셀 다운로드 (ZIP)", 
                    data=zip_buffer.getvalue(), 
                    file_name="거래명세서_원본유지.zip",
                    mime="application/zip"
                )
                st.success("데이터가 올바른 위치에 삽입된 엑셀 파일이 생성되었습니다!")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
