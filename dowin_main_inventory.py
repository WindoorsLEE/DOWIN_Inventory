# import remove_password
# import sheet_pivot0
# import sheet_memo
# import sheet_pivot
# import sheet_pivot2data
# import submaterial_pivot
# import partialsheet_pivot
# import total_pivot

import subprocess
if __name__ == "__main__":
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/remove_password.py"])  # remove password
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/sheet_pivot0.py"])  # remove password 파일에서 pivot data 로 전환(메모 제외)
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/sheet_memo.py"])  # remove password 파일에서 메모열 pivot data 로 전환
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/sheet_pivot.py"])  # sheet_pivot0와 sheet_memo파일을 결합하여 sheet pivot data 로 전환
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/sheet_pivot2data.py"])  # pivot 테이블에서 필요한 data만 추출
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/submaterial_pivot.py"])  # 부자재 재고 데이터를 pivot data 로 전환
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/partialsheet_pivot.py"])  # 잔여시트자재 재고 데이터를 pivot data 로 전환
    subprocess.run(["python", "C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/total_pivot.py"])  # 원판자재, 부자재, 잔여자재 데이터 시트 파일을 하나로 합친 파일
