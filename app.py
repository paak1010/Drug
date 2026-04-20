import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

def main():
    st.set_page_config(page_title="원본 양식 유지 명세서 생성", layout="wide")
    
    st.title("🎯 첫 번째 시트 셀 수정 자동화기")
    st.write("첫 번째 시트의 원본 서식(선, 폰트, 병합 등)을 100% 유지하고 값만 변경합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx'])

    if uploaded_file:
        try:
            # 1. 두 번째 시트(데이터 매트릭스)만 판다스로 읽어오기
            xls = pd.ExcelFile(uploaded_file)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            # 제품별 단가/입수량 마스터
            product_master = {
                "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
                "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
                "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
            }

            if st.button("첫 번째 시트 셀 수정 및 엑셀 일괄 생성"):
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
                            
                            # 수량이 있는 제품만 걸러내기
                            if pd.notna(qty) and str(qty).replace('.0','').isdigit() and int(float(qty)) > 0:
                                info = product_master.get(p_name, {"inbox": 1, "price": 0})
                                
                                # 할증은 단가 0원 처리
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

                        # 해당 지점에 배송할 제품이 있다면 원본 시트 셀 수정 작업 시작
                        if data_rows:
                            # [핵심] 원본 엑셀 파일을 그대로 메모리에 로드하여 도화지로 사용
                            uploaded_file.seek(0) 
                            wb = openpyxl.load_workbook(uploaded_file)
                            ws = wb.worksheets[0] # 첫 번째 시트 선택
                            
                            # 1. 공급받는자 상호 셀 수정 (예: I열 10행 위치)
                            # 👉 실제 엑셀 파일의 공급받는자 위치 셀에 맞게 알파벳과 숫자를 수정하세요!
                            ws['I10'] = str(center).strip()
                            
                            # 2. 제품 내역 셀 수정 (예: 15행부터 시작한다고 가정)
                            # 👉 실제 엑셀 파일의 No. 1 이 시작되는 행 번호로 start_row를 수정하세요!
                            start_row = 15
                            
                            for row_idx, item in enumerate(data_rows):
                                curr_row = start_row + row_idx
                                
                                ws[f'A{curr_row}'] = row_idx + 1               # No
                                ws[f'C{curr_row}'] = item['Product']           # 제품명
                                ws[f'I{curr_row}'] = item['InBox']             # 입수
                                ws[f'K{curr_row}'] = item['BoxQty']            # Box수
                                ws[f'L{curr_row}'] = item['Qty']               # 낱개수
                                
                                # 할증일 경우 단가/공급가액 빈칸 처리 (필요시 0으로 변경)
                                ws[f'M{curr_row}'] = item['Price'] if item['Price'] > 0 else ""
                                ws[f'N{curr_row}'] = item['Total'] if item['Total'] > 0 else ""

                            # 수정된 엑셀 파일을 메모리에 임시 저장하고 ZIP에 추가
                            excel_io = io.BytesIO()
                            wb.save(excel_io)
                            excel_io.seek(0)
                            
                            safe_center = str(center).replace('\n', ' ').replace('/', '_').strip()
                            zip_f.writestr(f"거래명세서_{safe_center}.xlsx", excel_io.read())
                
                # ZIP 파일 다운로드
                st.download_button(
                    label="📦 셀 수정 완료된 엑셀 파일 일괄 다운로드 (ZIP)", 
                    data=zip_buffer.getvalue(), 
                    file_name="거래명세서_엑셀원본유지.zip",
                    mime="application/zip"
                )
                st.success("원본 서식이 100% 유지된 채 데이터만 변경된 엑셀 파일이 생성되었습니다!")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
