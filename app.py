import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
import io
import zipfile
from datetime import datetime

# 1. 병합된 셀에도 안전하게 값을 쓰는 함수
def write_safe(ws, row, col, value):
    target_cell = ws.cell(row=row, column=col)
    if isinstance(target_cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if target_cell.coordinate in merged_range:
                # 병합 범위의 시작점(왼쪽 상단) 셀에 값을 입력
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        target_cell.value = value

def main():
    st.set_page_config(page_title="SCM 명세서 자동화 도구", layout="wide")
    
    st.title("🎯 원본 서식 유지 및 행 자동 삭제 명세서 생성기")
    st.write("첫 번째 시트의 양식을 그대로 사용하여 지점별 맞춤 엑셀을 생성합니다.")

    uploaded_file = st.file_uploader("명세서 엑셀 파일(.xlsx)을 업로드하세요", type=['xlsx'])

    if uploaded_file:
        try:
            # 수량 데이터가 있는 두 번째 시트 읽기
            xls = pd.ExcelFile(uploaded_file)
            df_matrix = pd.read_excel(xls, sheet_name=xls.sheet_names[1], header=2)
            
            if st.button("지점별 엑셀 파일 일괄 생성"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                    # 지점명 추출
                    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]
                    
                    for center in centers:
                        is_bonus = '할증' in str(center) or '힐증' in str(center)
                        
                        # 원본 서식 로드 (data_only=False로 서식 보존)
                        uploaded_file.seek(0)
                        wb = openpyxl.load_workbook(uploaded_file)
                        ws = wb.worksheets[0] 
                        
                        # --- 단계 1: 주요 열 및 행 위치 동적 찾기 ---
                        col_name = col_qty = col_price = col_total = -1
                        header_row = -1
                        footer_row = -1
                        
                        for r in range(1, 40):
                            for c in range(1, 20):
                                val = str(ws.cell(row=r, column=c).value)
                                if "제품명" in val: 
                                    header_row = r; col_name = c
                                elif "낱개수" in val or "날개수" in val: col_qty = c
                                elif "낱개가격" in val or "날개가격" in val: col_price = c
                                elif "공급가액" in val: col_total = c
                                # 배송처 문구 변경
                                if "배송처" in val:
                                    write_safe(ws, r, c, f"배송처: {str(center).strip()}")
                        
                        # 합계 행 찾기
                        for r in range(header_row + 1, 150):
                            for c in range(1, 15):
                                if "합계" in str(ws.cell(row=r, column=c).value):
                                    footer_row = r
                                    break
                            if footer_row != -1: break
                            
                        if header_row == -1 or footer_row == -1:
                            continue

                        # --- 단계 2: 데이터 입력 및 행 삭제 (역순 처리) ---
                        total_sum = 0
                        
                        # 데이터가 들어가는 영역만 역순으로 순회하며 삭제/수정
                        for r in range(footer_row - 1, header_row, -1):
                            p_cell_val = ws.cell(row=r, column=col_name).value
                            if not p_cell_val or str(p_cell_val).strip() == "":
                                ws.delete_rows(r, 1)
                                continue
                            
                            p_name = str(p_cell_val).strip()
                            # 두 번째 시트에서 수량 매칭
                            match = df_matrix[df_matrix['제품명'].astype(str).str.strip() == p_name]
                            
                            qty = 0
                            if not match.empty:
                                val = match.iloc[0][center]
                                if pd.notna(val): qty = int(float(val))
                            
                            if qty > 0:
                                # 낱개수 입력
                                write_safe(ws, r, col_qty, qty)
                                
                                # 기존 낱개가격 읽기
                                try:
                                    price_raw = ws.cell(row=r, column=col_price).value
                                    price = int(float(str(price_raw).replace(',', '')))
                                except:
                                    price = 0
                                
                                # 할증 지점은 0원 처리 및 명칭 변경
                                if is_bonus:
                                    price = 0
                                    write_safe(ws, r, col_price, 0)
                                    write_safe(ws, r, col_name, f"{p_name} (할증분)")
                                
                                # 공급가액 계산 및 입력
                                row_total = qty * price
                                write_safe(ws, r, col_total, row_total)
                                total_sum += row_total
                            else:
                                # 수량이 0인 제품은 행 삭제
                                ws.delete_rows(r, 1)

                        # --- 단계 3: 최종 합계 갱신 ---
                        # 행이 삭제되어 위치가 변한 합계 행을 다시 찾아 총액 기입
                        for r in range(header_row + 1, 100):
                            found = False
                            for c in range(1, 15):
                                if "합계" in str(ws.cell(row=r, column=c).value):
                                    write_safe(ws, r, col_total, total_sum)
                                    found = True
                                    break
                            if found: break

                        # 메모리에 저장
                        excel_out = io.BytesIO()
                        wb.save(excel_out)
                        zip_f.writestr(f"거래명세서_{str(center).strip()}.xlsx", excel_out.getvalue())
                
                st.download_button(
                    label="📦 최종 명세서 일괄 다운로드 (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"거래명세서_완성본_{datetime.now().strftime('%m%d')}.zip",
                    mime="application/zip"
                )
                st.success("모든 요청사항이 반영된 파일 생성이 완료되었습니다!")

        except Exception as e:
            st.error(f"오류 발생: {e}")

if __name__ == "__main__":
    main()
