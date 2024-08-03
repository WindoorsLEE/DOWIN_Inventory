import os

# 특정 디렉토리를 지정
folder_path = 'C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/'

# 바꿀 텍스트와 대체할 텍스트를 딕셔너리로 지정
replacements = {
    'C:/Users/windo/OneDrive/바탕 화면/Zenbook2/DOWIN/inventory/': folder_path,
    # 'C:\Users\windo\OneDrive\바탕 화면\Zenbook2\DOWIN\inventory\': folder_path
}

# 폴더 내 모든 파일을 검색하고 파이썬 파일만 처리
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.py'):
            file_path = os.path.join(root, file)
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 여러 텍스트 대체
            for old_text, new_text in replacements.items():
                content = content.replace(old_text, new_text)
            
            # 파일에 다시 저장
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)

print("모든 파일에서 텍스트 대체 완료.")
