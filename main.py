import tkinter as tk
from tkinter import filedialog
import openpyxl
import win32com.client
import time
import win32gui
from pyhwpx import Hwp
import pyautogui
import psutil
import os
import re

def correct_double_spaces_snippet(text: str) -> str:
    """
    문자열에서 연속된 공백을 찾아 해당 부분의 미리보기를 보여주고,
    사용자 선택에 따라 전체 문자열을 수정합니다.

    Args:
        text (str): 검사할 원본 문자열

    Returns:
        str: 연속된 공백이 수정되었거나 수정하지 않은 원본 문자열
    """
    # 'finditer'를 사용해 2개 이상 연속된 모든 공백의 위치(match 객체)를 찾음
    matches = list(re.finditer(r' {2,}', text))

    if matches:
        print(f"\n⚠️ 총 {len(matches)}곳에서 연속된 공백이 발견되었습니다!")
        print("="*50)

        # 발견된 각 위치의 미리보기를 순서대로 보여줌
        for i, match in enumerate(matches):
            # match.start()는 연속 공백이 시작되는 위치
            # match.end()는 연속 공백이 끝나는 위치
            
            # 미리보기의 시작과 끝 지점 계산 (앞뒤 5글자씩)
            preview_start = max(0, match.start() - 5)
            preview_end = min(len(text), match.end() + 5)

            # 원본에서 미리보기 부분만 잘라내기
            original_snippet = text[preview_start:preview_end]
            # 잘라낸 부분에서 연속 공백을 하나로 수정
            modified_snippet = re.sub(r' {2,}', ' ', original_snippet)
            
            # 문자열의 처음이나 끝이 아닐 경우 ...으로 표시
            prefix = "..." if preview_start > 0 else ""
            suffix = "..." if preview_end < len(text) else ""

            print(f"📌 위치 #{i+1}")
            print(f"  [원본]: {prefix}{original_snippet}{suffix}")
            print(f"  [수정]: {prefix}{modified_snippet}{suffix}\n")
        
        print("="*50)
        # 모든 미리보기를 보여준 후, 전체 수정 여부를 한 번에 물어봄
        choice = input("위에 표시된 모든 공백을 한 번에 수정하시겠습니까? (y/N(기본)): ").lower().strip()

        if choice == 'y':
            # 'y'를 입력하면 전체 텍스트를 대상으로 수정 작업 진행
            corrected_text = re.sub(r' {2,}', ' ', text)
            print("\n✅ 전체 문자열의 공백을 수정했습니다.")
            return corrected_text
        else:
            print("\n❌ 원본을 그대로 유지합니다.")
            return text
    else:
        # 연속된 공백이 없는 경우
        print("\n✅ 연속된 공백이 없어 원본을 그대로 반환합니다.")
        return text

def check_if_processes_running():
    """
    한글과 엑셀 프로그램이 현재 실행 중인지 확인하고 결과를 딕셔너리로 반환합니다.

    Returns:
        dict: {'hwp': bool, 'excel': bool}
              'hwp'는 한글, 'excel'은 엑셀의 실행 여부를 나타냅니다.
    """
    process_status = {'hwp': False, 'excel': False}
    
    # 현재 실행 중인 모든 프로세스를 순회합니다.
    for proc in psutil.process_iter(['name']):
        try:
            process_name = proc.info['name'].lower() # 프로세스 이름을 소문자로 변경
            
            # 프로세스 이름으로 한글 또는 엑셀 실행 여부를 확인합니다.
            if 'hwp.exe' in process_name:
                process_status['hwp'] = True
            
            if 'excel.exe' in process_name:
                process_status['excel'] = True

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 일부 프로세스는 접근 권한이 없거나 이미 종료되었을 수 있습니다.
            pass
    return process_status


# 시작 화면 초기화
os.system('cls')

# 루트 윈도우 생성 및 숨기기
root = tk.Tk()
root.withdraw()

print("="*50)
print("한글-엑셀 세부특기사항 수정 도우미 v1.0")
print("="*50)
print("\n이 프로그램은 선생님들의 세부특기사항 수정을 보다 쉽게 도와드리기 위해 제작되었습니다.")
print("made by 성준\n")
print("아래아한글과 엑셀창을 반드시 모두 종료한 뒤 이 프로그램을 사용하여 주시길 바랍니다.\n")

checked_running = check_if_processes_running()
warning_shown = False

if checked_running['hwp']:
    warning_shown = True
    print("="*50)
    print("⚠️ 한글과 컴퓨터 프로그램이 실행중입니다.")
    print("="*50)

if checked_running['excel']:
    warning_shown = True
    print("="*50)
    print("⚠️ 엑셀 프로그램이 실행중입니다.")
    print("="*50)

if warning_shown:
    print()

print("프로그램 사용에 따른 문제나 손실은 모두 사용자에게 있으며")
print("반드시 프로그램 사용중 다른 곳에 저장하는것을 권장드립니다.")
print("계속하시게 된다면 이에 동의하는 것으로 간주됩니다.\n")
input("계속하려면 Enter 키를 누르세요...")

# 파일 선택 화면으로 전환
os.system('cls')
print("="*50)
print("수정을 원하는 엑셀 파일을 선택 해 주세요")
print("="*50)
print("\n파일 선택 창이 열립니다...\n")

# 파일 선택 대화상자 표시
file_path = filedialog.askopenfilename(
    title="엑셀 파일 선택",
    filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")]
)

if not file_path:
    os.system('cls')
    print("\n파일이 선택되지 않았습니다. 프로그램을 종료합니다.\n")
    input("Enter 키를 눌러 종료하세요...")
    exit()

hwp = False

# Excel 애플리케이션 실행 및 파일 열기
os.system('cls')
print("\n엑셀 파일을 여는 중입니다. 잠시만 기다려주세요...\n")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
workbook = excel.Workbooks.Open(file_path)

def bring_excel_to_front():
    pyautogui.press("alt")
    win32gui.ShowWindow(excel.hwnd, 3)
    win32gui.SetForegroundWindow(excel.hwnd)

bring_excel_to_front()

# 작업 메인 화면
os.system('cls')
print("="*50)
print(f"파일: {file_path.split('/')[-1]}")
print("="*50)
print("\n작업을 시작합니다. 수정할 셀을 선택하고 Enter 키를 누르세요.\n")

while True:
    print("="*50)
    print("1. 수정할 셀을 선택한 후 Enter 키를 누르세요.")
    print("2. 종료하려면 '종료' 또는 'exit'를 입력하세요.")
    print("="*50)
    
    # 사용자 입력 받기
    user_input = input("\n명령을 입력하세요: ")
    
    # 종료 조건 확인
    if user_input.lower() in ['종료', 'exit', 'quit', '끝']:
        os.system('cls')
        print("\n작업을 종료합니다. 변경사항이 저장되었습니다.\n")
        break
    
    # 한글 초기화
    if not hwp:
        hwp = Hwp()
    
    try:
        # 현재 선택된 셀 정보 가져오기
        selected_cell = excel.Selection
        cell_value = selected_cell.Text

        # 선택된 셀 정보 출력 및 작업 시작
        os.system('cls')
        print("="*50)
        print(f"선택한 셀: {selected_cell.Address}")
        print("="*50)
        print("\n선택한 셀의 내용:")
        print("-"*50)
        print(f"{cell_value}")
        print("-"*50)
        print("\n한글 프로그램에서 맞춤법 검사를 시작합니다...\n")

        # 한글 새 탭 생성
        hwp.FileNewTab()

        # 텍스트 삽입 및 맞춤법 검사
        hwp.maximize_window()
        hwp.insert_text(cell_value)
        hwp.move_pos(move_id=2)
        hwp.SpellingCheck()

        print("\n한글에서 맞춤법 검사와 수정을 완료하신 후 Enter 키를 누르세요.")
        input("→ 계속하려면 Enter 키를 누르세요: ")

        # 한글에서 수정된 텍스트 가져오기
        confirmed_text = hwp.get_page_text(pgno=0, option=4294967295)

        # 공백 수정 안내
        os.system('cls')
        print("="*50)
        print("맞춤법 검사 완료")
        print("="*50)
        print("\n공백 확인 중...")
        
        # 공백 수정 작업
        confirmed_text = correct_double_spaces_snippet(confirmed_text)
        
        # 엑셀로 텍스트 다시 입력
        excel.Range(selected_cell.Address).Value = confirmed_text
        hwp.clear(option=1)
        hwp.minimize_window()

        bring_excel_to_front()
        
        os.system('cls')
        print("="*50)
        print("작업 완료")
        print("="*50)
        print("\n✅ 텍스트가 성공적으로 엑셀에 업데이트되었습니다.\n")
        print("다음 셀을 선택하거나 종료 명령을 입력하세요.\n")
        
    except Exception as e:
        os.system('cls')
        print("="*50)
        print("오류 발생")
        print("="*50)
        print(f"\n❌ 오류가 발생했습니다: {e}\n")
        print("다시 시도해주세요.\n")
        
        # 오류 발생시에도 한글 문서 정리
        try:
            hwp.clear(option=1)
        except:
            pass

# 프로그램 종료 및 정리
os.system('cls')
print("="*50)
print("프로그램 종료")
print("="*50)
print("\n모든 작업이 완료되었습니다. 감사합니다.\n")

# 정리
try:
    hwp.clear(option=1)
except:
    pass

input("Enter 키를 눌러 창을 닫으세요...")