import os
from openpyxl import load_workbook

# 현재 스크립트 파일의 디렉토리 경로를 가져옵니다.
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_file_path = os.path.join(current_dir, '02_test.xlsx')

# 엑셀 파일 열기
wb = load_workbook(filename=excel_file_path)
sheet = wb.active

# B1과 B2의 값을 가져옵니다.
b1_value = sheet['B1'].value
b2_value = sheet['B2'].value

# B1에서 B2를 뺀 값을 계산합니다.
if isinstance(b1_value, (int, float)) and isinstance(b2_value, (int, float)):
    result = b1_value - b2_value
else:
    raise ValueError("B1 또는 B2의 값이 숫자가 아닙니다.")

# 결과를 B3 셀에 입력합니다.
sheet['B3'] = result

# 엑셀 파일 저장
wb.save(filename=excel_file_path)
wb.close()

print("B1 - B2 =", result)
