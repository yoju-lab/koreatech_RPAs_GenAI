from openpyxl import load_workbook

# 호스트 OS와 공유된 Excel 파일 경로
file_path = '/shared/path/to/excel_file.xlsx'

# Excel 파일 열기
wb = load_workbook(file_path)
sheet = wb.active

# B1과 B2의 값을 가져옵니다.
b1_value = sheet['B1'].value
b2_value = sheet['B2'].value

# B1에서 B2를 뺀 값을 계산합니다.
result = b1_value - b2_value

# 결과를 B3 셀에 입력합니다.
sheet['B3'] = result

# 변경 사항 저장
wb.save(file_path)