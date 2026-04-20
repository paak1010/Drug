import pandas as pd
import openpyxl
import win32com.client
import os

def generate_exact_pdfs(excel_filename):
    # 현재 폴더의 절대 경로를 가져옵니다 (엑셀 조종을 위해 필수)
    base_dir = os.getcwd()
    excel_path = os.path.join(base_dir, excel_filename)
    
    print("1. 엑셀의 두 번째 시트(물량 데이터)를 읽어옵니다...")
    df_matrix = pd.read_excel(excel_path, sheet_name=1, header=2)
    
    # 단가 및 입수량 마스터
    product_master = {
        "멘소래담 로션 75ml": {"inbox": 72, "price": 3780},
        "멘소래담 로션 100ml": {"inbox": 72, "price": 4590},
        "멘소래담 로션 450ml": {"inbox": 20, "price": 13950}
    }
    
    # 결과물이 저장될 폴더 생성
    output_dir = os.path.join(base_dir, "출력된_명세서_PDF")
    os.makedirs(output_dir, exist_ok=True)

    print("2. 마이크로소프트 엑셀 엔진을 가동합니다 (화면엔 안 보입니다)...")
    # 실제 MS Excel을 백그라운드에서 실행
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    centers = [c for c in df_matrix.columns if c not in ['제품명', '규격', 'Total'] and not pd.isna(c)]

    for center in centers:
        is_bonus = '할증' in str(center) or '힐증' in str(center)
        
        # [핵심] 원본 엑셀 파일을 매번 그대로 복사해옵니다 (서식, 폰트 100% 유지)
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.worksheets[0] # 첫 번째 시트(도화지) 선택
        
        # 1. 공급받는자 상호 입력 
        # (주의: 실제 엑셀의 상호명 셀 위치에 맞게 알파벳을 수정하세요! 예: I10, J10 등)
        ws['I10'] = str(center).strip() 
        
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
        
        if not data_rows:
            continue # 해당 지점에 배송할 물건이 없으면 패스
            
        # 2. 표 안에 데이터 갈아끼우기 (엑셀 15행부터 시작한다고 가정)
        # (주의: 실제 제품명이 들어가는 시작 행 번호에 맞게 start_row를 수정하세요!)
        start_row = 15 
        for row_idx, item in enumerate(data_rows):
            curr_row = start_row + row_idx
            
            ws[f'A{curr_row}'] = row_idx + 1               # No
            ws[f'C{curr_row}'] = item['Product']           # 제품명
            ws[f'I{curr_row}'] = item['InBox']             # 입수
            ws[f'K{curr_row}'] = item['BoxQty']            # Box수
            ws[f'L{curr_row}'] = item['Qty']               # 낱개수
            
            if item['Price'] == 0:
                ws[f'M{curr_row}'] = "" # 할증이면 0원 대신 빈칸 처리 (필요시 0으로 변경)
                ws[f'N{curr_row}'] = "" 
            else:
                ws[f'M{curr_row}'] = item['Price']         # 낱개가격
                ws[f'N{curr_row}'] = item['Total']         # 공급가액

        # 임시 엑셀 파일로 저장
        safe_center = str(center).replace('\n', ' ').replace('/', '_').strip()
        temp_xlsx_path = os.path.join(base_dir, f"temp_{safe_center}.xlsx")
        wb.save(temp_xlsx_path)
        
        # 3. 엑셀 프로그램이 직접 임시 파일을 열어 완벽한 PDF로 인쇄
        try:
            pdf_path = os.path.join(output_dir, f"거래명세서_{safe_center}.pdf")
            wb_excel = excel.Workbooks.Open(temp_xlsx_path)
            # 0 은 xlTypePDF 옵션입니다.
            wb_excel.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
            wb_excel.Close(False)
            print(f"✅ 생성 완료: 거래명세서_{safe_center}.pdf")
        except Exception as e:
            print(f"❌ PDF 변환 실패 ({safe_center}): {e}")
        finally:
            # 작업이 끝난 임시 엑셀 파일은 삭제하여 깔끔하게 유지
            if os.path.exists(temp_xlsx_path):
                os.remove(temp_xlsx_path)

    excel.Quit()
    print("\n🎉 모든 작업이 완료되었습니다! '출력된_명세서_PDF' 폴더를 확인해주세요.")

if __name__ == "__main__":
    # 선생님의 원본 엑셀 파일명을 정확히 입력해주세요.
    generate_exact_pdfs("거래명세서_백제약품 센터별 (0421).xlsx")
