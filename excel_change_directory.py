import pandas as pd
import os

# 치환할 문자 선언
old_char1 = '=+https://d.docs.live.net/d2b8db7440f53975/바탕 화면/Zenbook2/DOWIN/inventory/test/'
old_char2 = '=+C:\\Users\\windo\\OneDrive\\바탕 화면\\Zenbook2\\DOWIN\\inventory\\'
new_char = '=+C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/'

# 작업 폴더 경로 설정
folder_path = r'C:\Users\windo\OneDrive\바탕 화면\Zenbook2\DOWIN\inventory\test'
input_file = os.path.join(folder_path, '재고수불부-total_pivot(관리용).xlsx')
output_file = os.path.join(folder_path, '재고수불부-total_pivot(관리용)_updated.xlsx')

# 엑셀 파일 읽기
df = pd.read_excel(input_file)

# 데이터 치환
df = df.applymap(lambda x: x.replace(old_char1, new_char).replace(old_char2, new_char) if isinstance(x, str) else x)

# 변경된 데이터 엑셀 파일로 저장
df.to_excel(output_file, index=False)

print(f"변경된 데이터를 {output_file} 파일로 저장했습니다.")
