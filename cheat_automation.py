import pygetwindow as gw
import pywinctl
import pyautogui
import cv2
import numpy as np
import pandas as pd
import pyperclip
import time
import os

# 현재 스크립트 위치 기준으로 상대 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
EXCEL_DIR = os.path.join(BASE_DIR, 'excel_data')

# 치트 엑셀 파일 경로
CHEAT_FILE = os.path.join(BASE_DIR, '치트.xlsx')

class GameCheater:
    def __init__(self):
        self.window = None
        self.cheat_data = None
        self.load_cheat_data()
    
    def load_cheat_data(self):
        """치트 데이터 엑셀 파일에서 불러오기"""
        try:
            self.cheat_data = pd.read_excel(CHEAT_FILE)
            print(f"치트 데이터 로드 완료: {len(self.cheat_data)} 개 항목")
        except Exception as e:
            print(f"치트 데이터 로드 실패: {e}")
            self.cheat_data = None
    
    def select_window(self):
        """활성화된 윈도우 목록에서 게임 윈도우 선택"""
        # 모든 윈도우 목록 가져오기
        active_windows = pywinctl.getAllWindows()
        print(f"총 {len(active_windows)}개 윈도우 감지됨")
        # 테스트를 위해 모든 윈도우를 사용
        # active_windows = [w for w in active_windows if w.visible]
        
        if not active_windows:
            print("활성화된 윈도우가 없습니다.")
            return False
        
        # 윈도우 목록 출력
        print("\n활성화된 윈도우 목록:")
        for i, window in enumerate(active_windows):
            title = window.title if hasattr(window, 'title') else str(window)
            print(f"{i+1}. {title}")
        
        # 사용자 선택
        try:
            selection = int(input("\n게임 윈도우 번호를 선택하세요: ")) - 1
            if 0 <= selection < len(active_windows):
                self.window = active_windows[selection]
                # 윈도우 활성화
                if hasattr(self.window, 'activate'):
                    self.window.activate()
                else:
                    self.window.focus()  # pywinctl에서는 focus() 메서드를 사용
                
                title = self.window.title if hasattr(self.window, 'title') else str(self.window)
                print(f"'{title}' 윈도우 선택됨")
                time.sleep(1)  # 안정성을 위한 대기
                return True
            else:
                print("잘못된 선택입니다.")
                return False
        except ValueError:
            print("숫자를 입력해주세요.")
            return False
    
    def find_image_on_screen(self, template_name, threshold=0.8):
        """화면에서 이미지 찾기"""
        template_path = os.path.join(TEMPLATES_DIR, template_name)
        
        # 화면 캡처
        screenshot = pyautogui.screenshot()
        screenshot = np.array(screenshot)
        screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
        
        # 템플릿 이미지 로드
        template = cv2.imread(template_path)
        
        if template is None:
            print(f"템플릿 이미지를 찾을 수 없습니다: {template_path}")
            return None
        
        # 이미지 매칭
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        if max_val >= threshold:
            # 탐지된 위치의 중앙점 계산
            h, w = template.shape[:2]
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            return (center_x, center_y, max_val)
        
        return None
    
    def open_cheat_menu(self):
        """치트 메뉴 열기 - menu2가 있으면 바로 menu3 클릭, 없으면 menu 먼저 클릭 후 menu3 클릭"""
        # menu2의 존재 여부 확인
        menu2_result = self.find_image_on_screen('menu2.png')
        
        if menu2_result:
            print("menu2 발견: 치트 메뉴가 이미 열려있습니다.")
            # menu3 클릭
            menu3_result = self.find_image_on_screen('menu3.png')
            if menu3_result:
                pyautogui.click(menu3_result[0], menu3_result[1])
                print("menu3 클릭 완료")
                time.sleep(0.5)
                return True
            else:
                print("menu3를 찾을 수 없습니다.")
                return False
        else:
            # menu 클릭
            menu_result = self.find_image_on_screen('menu.png')
            if menu_result:
                pyautogui.click(menu_result[0], menu_result[1])
                print("menu 클릭 완료")
                time.sleep(0.5)
                
                # menu3 클릭
                menu3_result = self.find_image_on_screen('menu3.png')
                if menu3_result:
                    pyautogui.click(menu3_result[0], menu3_result[1])
                    print("menu3 클릭 완료")
                    time.sleep(0.5)
                    return True
                else:
                    print("menu3를 찾을 수 없습니다.")
                    return False
            else:
                print("menu를 찾을 수 없습니다.")
                return False
    
    def show_available_cheats(self):
        """사용 가능한 치트 목록 출력"""
        if self.cheat_data is None:
            print("치트 데이터가 로드되지 않았습니다.")
            return None
        
        print("\n사용 가능한 치트 목록:")
        cheat_names = self.cheat_data['Unnamed: 1'].dropna().tolist()
        cheat_codes = self.cheat_data['Unnamed: 2'].dropna().tolist()
        
        for i, (name, code) in enumerate(zip(cheat_names, cheat_codes)):
            if pd.notna(name) and pd.notna(code):
                print(f"{i+1}. {name} - {code}")
        
        try:
            selection = int(input("\n사용할 치트 번호를 선택하세요: ")) - 1
            if 0 <= selection < len(cheat_codes):
                return cheat_codes[selection]
            else:
                print("잘못된 선택입니다.")
                return None
        except ValueError:
            print("숫자를 입력해주세요.")
            return None
    
    def execute_cheat(self, cheat_code):
        """치트 실행 - 코드 입력 및 실행"""
        # 코드를 클립보드에 복사
        pyperclip.copy(cheat_code)
        print(f"치트 코드 '{cheat_code}' 복사됨")
        
        # code 클릭
        code_result = self.find_image_on_screen('code.png')
        if not code_result:
            print("code 버튼을 찾을 수 없습니다.")
            return False
        
        pyautogui.click(code_result[0], code_result[1])
        print("code 클릭 완료")
        time.sleep(0.5)
        
        # code2 클릭
        code2_result = self.find_image_on_screen('code2.png')
        if not code2_result:
            print("code2 버튼을 찾을 수 없습니다.")
            return False
        
        pyautogui.click(code2_result[0], code2_result[1])
        print("code2 클릭 완료")
        time.sleep(0.5)
        
        # 코드 붙여넣기
        pyautogui.hotkey('ctrl', 'v')
        print("코드 붙여넣기 완료")
        time.sleep(0.5)
        
        # code3 클릭
        code3_result = self.find_image_on_screen('code3.png')
        if not code3_result:
            print("code3 버튼을 찾을 수 없습니다.")
            return False
        
        pyautogui.click(code3_result[0], code3_result[1])
        print("code3 클릭 완료")
        time.sleep(0.5)
        
        # code4 클릭
        code4_result = self.find_image_on_screen('code4.png')
        if not code4_result:
            print("code4 버튼을 찾을 수 없습니다.")
            return False
        
        pyautogui.click(code4_result[0], code4_result[1])
        print("code4 클릭 완료, 치트 실행됨")
        return True
    
    def run(self):
        """전체 치트 프로세스 실행"""
        print("게임 치트 자동화 프로그램 시작")
        
        # 윈도우 선택
        if not self.select_window():
            return
        
        # 치트 메뉴 열기
        if not self.open_cheat_menu():
            return
        
        # 치트 선택
        cheat_code = self.show_available_cheats()
        if not cheat_code:
            return
        
        # 치트 실행
        self.execute_cheat(cheat_code)

# 메인 실행
if __name__ == "__main__":
    try:
        cheater = GameCheater()
        cheater.run()
    except KeyboardInterrupt:
        print("\n프로그램이 종료되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")
