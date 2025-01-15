import os
import requests
import time
import tkinter as tk
from tkinter import filedialog

# 구글 스프레드시트 링크에서 XLSX 파일 다운로드
def download_google_sheet_as_xlsx(sheet_url, output_filename):
    # 구글 스프레드시트 URL을 XLSX 형식으로 수정
    xlsx_url = sheet_url.replace('/edit?usp=sharing', '/export?format=xlsx')
    
    # 파일 다운로드 요청
    response = requests.get(xlsx_url)
    
    if response.status_code == 200:
        # 다운로드 받은 파일을 지정된 파일명으로 저장
        with open(output_filename, 'wb') as f:
            f.write(response.content)
        print(f"파일 다운로드 완료: {output_filename}")
    else:
        print("파일 다운로드 실패:", response.status_code)

# 실행 예시
sheet_url = "https://docs.google.com/spreadsheets/d/1om5aH7l1khkN0wk2ET1E6KzsmXv1_BpzstGBrMfB5Uo/edit?usp=sharing"  # 구글 스프레드시트 링크

# Tkinter 윈도우 초기화 (파일 선택 창을 띄우기 위한 GUI)
root = tk.Tk()
root.withdraw()  # 기본 윈도우를 숨김

# 사용자에게 저장할 폴더 경로를 선택하게 하기 위한 창 띄우기
output_directory = filedialog.askdirectory(title="파일을 저장할 폴더를 선택하세요")
if not output_directory:  # 사용자가 경로를 선택하지 않으면 종료
    print("폴더 선택이 취소되었습니다.")
    exit()

# 사용자에게 파일 이름 입력 받기
output_filename = input("저장할 파일 이름을 입력해주세요: ") + ".xlsx"  # 확장자는 자동으로 추가

# 전체 파일 경로 생성
output_path = os.path.join(output_directory, output_filename)

# 경로 출력
print(f"파일이 저장될 경로: {output_path}")
time.sleep(1)

# 다운로드 함수 실행
download_google_sheet_as_xlsx(sheet_url, output_path)

# 완료 후 경로 출력
print(f"파일이 성공적으로 다운로드되었습니다.\n")
print(f"경로: {output_path}")
