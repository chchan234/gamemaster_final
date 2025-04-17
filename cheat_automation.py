import pygetwindow as gw
import pywinctl
import pyautogui
import cv2
import numpy as np
import pandas as pd
import pyperclip
import time
import os
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import json

# 현재 스크립트 위치 기준으로 상대 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
EXCEL_DIR = os.path.join(BASE_DIR, 'excel_data')

# 치트 엑셀 파일 경로
CHEAT_FILE = os.path.join(BASE_DIR, 'cheat.xlsx')

class GameCheaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("게임 치트 자동화 프로그램")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.window = None
        self.cheat_data = None
        self.active_windows = []
        self.window_titles = []
        self.threshold = 0.6
        self.current_category = None
        self.cheat_categories = {}  # 치트 카테고리 저장할 딕셔너리
        
        self.create_gui()
        self.load_cheat_categories()
        self.get_window_list()
        
    def create_gui(self):
        # 메인 프레임 설정
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 탭 컨트롤 생성
        self.tab_control = ttk.Notebook(self.main_frame)
        
        # 세 개의 탭 생성
        self.window_tab = ttk.Frame(self.tab_control)
        self.cheat_tab = ttk.Frame(self.tab_control)
        self.log_tab = ttk.Frame(self.tab_control)
        
        # 탭 추가
        self.tab_control.add(self.window_tab, text="윈도우 선택")
        self.tab_control.add(self.cheat_tab, text="치트 카테고리")
        self.tab_control.add(self.log_tab, text="로그")
        
        self.tab_control.pack(expand=1, fill=tk.BOTH)
        
        # 각 탭 설정
        self.setup_window_tab()
        self.setup_cheat_tab()
        self.setup_log_tab()
    
    def setup_window_tab(self):
        # 윈도우 선택 탭 설정
        window_frame = ttk.LabelFrame(self.window_tab, text="게임 윈도우 선택", padding="10")
        window_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 윈도우 리스트 설명
        ttk.Label(window_frame, text="아래 목록에서 게임 윈도우를 선택해주세요:").pack(anchor=tk.W, padx=10, pady=10)
        
        # 윈도우 목록 프레임
        list_frame = ttk.Frame(window_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 윈도우 목록 리스트박스
        self.window_listbox = tk.Listbox(list_frame, width=70, height=15)
        self.window_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.window_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        button_frame = ttk.Frame(window_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        refresh_btn = ttk.Button(button_frame, text="새로고침", command=self.get_window_list)
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        apply_btn = ttk.Button(button_frame, text="선택 적용", command=self.apply_selected_window_and_switch_tab)
        apply_btn.pack(side=tk.RIGHT, padx=5)
        
        # 임계값 설정 프레임
        threshold_frame = ttk.LabelFrame(window_frame, text="이미지 인식 설정", padding="10")
        threshold_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(threshold_frame, text="인식 임계값:").pack(side=tk.LEFT, padx=5)
        self.threshold_var = tk.DoubleVar(value=self.threshold)
        threshold_scale = ttk.Scale(threshold_frame, from_=0.1, to=1.0, orient=tk.HORIZONTAL, 
                                  length=200, variable=self.threshold_var, command=self.update_threshold)
        threshold_scale.pack(side=tk.LEFT, padx=5)
        
        self.threshold_label = ttk.Label(threshold_frame, text=f"{self.threshold:.1f}")
        self.threshold_label.pack(side=tk.LEFT, padx=5)
        
        # 디버그 버튼
        debug_btn = ttk.Button(threshold_frame, text="템플릿 디버그", command=self.debug_templates)
        debug_btn.pack(side=tk.LEFT, padx=20)
        
    def setup_cheat_tab(self):
        # 치트 카테고리 탭 설정
        cheat_frame = ttk.Frame(self.cheat_tab, padding="10")
        cheat_frame.pack(fill=tk.BOTH, expand=True)
        
        # 선택된 윈도우 표시
        self.window_info_label = ttk.Label(cheat_frame, text="선택된 윈도우: 없음", font=("Arial", 10, "bold"))
        self.window_info_label.pack(anchor=tk.W, padx=10, pady=5)
        
        # 카테고리 선택 영역
        category_frame = ttk.LabelFrame(cheat_frame, text="치트 카테고리 선택", padding="10")
        category_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 카테고리 드롭다운
        ttk.Label(category_frame, text="카테고리:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(category_frame, textvariable=self.category_var, 
                                      width=60, state="readonly")
        self.category_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 카테고리 목록 설정
        self.category_combo['values'] = list(self.cheat_categories.keys())
        
        # 카테고리 선택 이벤트 바인딩
        self.category_combo.bind("<<ComboboxSelected>>", self.on_category_selected)
        
        # 치트 선택 영역
        cheat_select_frame = ttk.LabelFrame(cheat_frame, text="치트 선택", padding="10")
        cheat_select_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(cheat_select_frame, text="치트:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 치트 드롭다운
        self.cheat_var = tk.StringVar()
        self.cheat_combo = ttk.Combobox(cheat_select_frame, textvariable=self.cheat_var, 
                                   width=60, state="readonly")
        self.cheat_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 치트 선택 이벤트 바인딩
        self.cheat_combo.bind("<<ComboboxSelected>>", self.on_cheat_selected)
        
        # 파라미터 입력 프레임 (중괄호 포함 치트 코드용)
        self.param_frame = ttk.Frame(cheat_select_frame)
        self.param_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
        self.param_entries = {}  # 파라미터 입력 필드를 저장할 딕셔너리
        
        # 실행 버튼 영역
        button_frame = ttk.Frame(cheat_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        execute_btn = ttk.Button(button_frame, text="치트 실행", command=self.execute_selected_cheat)
        execute_btn.pack(side=tk.RIGHT, padx=5)
    
        # 설명 표시 영역
        description_frame = ttk.LabelFrame(cheat_frame, text="치트 설명", padding="10")
        description_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.description_text = tk.Text(description_frame, wrap=tk.WORD, width=70, height=10, 
                                  font=("Courier", 10), state=tk.DISABLED)
        self.description_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def on_category_selected(self, event):
        """카테고리가 선택되었을 때 호출되는 함수"""
        category = self.category_var.get()
        if not category:
            return
            
        self.select_category(category)
    
    def select_category(self, category):
        """카테고리 선택 시 해당 카테고리의 치트만 표시"""
        self.current_category = category
        
        # 치트 콤보박스 업데이트 - 코드 부분을 제외한 이름만 표시
        cheat_display_names = []
        self.full_cheat_data = {}  # 치트 이름을 키로, 전체 치트 문자열을 값으로 저장
        
        self.log(f"카테고리 '{category}' 치트 목록 처리 시작")
        
        # 치트 목록 가져오기
        cheat_list = self.cheat_categories.get(category, [])
        self.log(f"카테고리 '{category}'에서 {len(cheat_list)}개 치트 로드됨")
        
        for cheat in cheat_list:
            if " — " in cheat:
                display_name = cheat.split(" — ")[0]  # "HP,MP 전체 회복" 부분만 추출
                cheat_display_names.append(display_name)
                self.full_cheat_data[display_name] = cheat
                self.log(f"치트 등록: '{display_name}' -> '{cheat}'")
            else:
                cheat_display_names.append(cheat)
                self.full_cheat_data[cheat] = cheat
                self.log(f"치트 등록(코드 없음): '{cheat}'")
                
        self.cheat_combo['values'] = cheat_display_names
        if len(cheat_display_names) > 0:
            self.cheat_combo.current(0)  # 첫 번째 항목 선택
            self.on_cheat_selected(None)  # 처음 선택된 치트에 대한 파라미터 필드 생성
        
        # 설명 텍스트 업데이트
        self.update_description()
        
        self.log(f"카테고리 '{category}' 선택됨, {len(cheat_list)}개 치트 표시")
    
    def update_description(self):
        """현재 선택된 치트에 대한 설명 업데이트"""
        selected_cheat_display = self.cheat_var.get()
        
        # 설명 텍스트 업데이트
        self.description_text.config(state=tk.NORMAL)
        self.description_text.delete(1.0, tk.END)
        
        if selected_cheat_display:
            # 전체 치트 문자열 가져오기 (코드 포함)
            full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
            
            # 설명 텍스트 생성
            description = f"선택된 치트: {selected_cheat_display}\n\n"
            
            # 코드 부분 추출 (GT.로 시작하는 부분)
            if " — GT." in full_cheat:
                cheat_code = full_cheat.split(" — ")[-1]
                description += f"실행될 코드: {cheat_code}\n"
            
            self.description_text.insert(tk.END, description)
        
        self.description_text.config(state=tk.DISABLED)
    
    def load_cheat_categories(self):
        """치트 카테고리 데이터 로드 - 엑셀에서만 로드"""
        try:
            # 초기화 - 빈 카테고리 딕셔너리
            self.cheat_categories = {}
            self.use_excel_data = False
            
            # 엑셀 파일 로드
            if os.path.exists(CHEAT_FILE):
                try:
                    # 위치 기반 접근을 위해 헤더 없이 로드
                    self.cheat_data = pd.read_excel(CHEAT_FILE, header=None)
                    self.log(f"치트 데이터 로드 완료: {len(self.cheat_data)} 개 항목")
                    
                    # 엑셀 파일의 내용 일부 출력 (디버깅)
                    for i in range(min(5, len(self.cheat_data))):
                        row_data = [str(x) if not pd.isna(x) else "NaN" for x in self.cheat_data.iloc[i]]
                        self.log(f"행 {i}: {row_data}")
                    
                    # 엑셀 데이터 처리 - 직접 행과 열 지정 (엑셀 파일 구조 기반)
                    current_category = None
                    
                    # 파일 구조 파악 (컬럼 헤더가 있는지)
                    header_row = -1
                    for i in range(min(10, len(self.cheat_data))):
                        row = self.cheat_data.iloc[i]
                        row_data = [str(x).strip() for x in row if not pd.isna(x)]
                        if any(header in row_data for header in ['치트명', '치트키', '이름', '코드']):
                            header_row = i
                            self.log(f"컬럼 헤더 행 발견: {header_row}")
                            break
                    
                    # 각 행 처리
                    for i in range(len(self.cheat_data)):
                        row = self.cheat_data.iloc[i]
                        
                        # 빈 행이면 건너뛰기
                        if all(pd.isna(x) for x in row):
                            continue
                            
                        # 헤더 행이면 건너뛰기    
                        if i == header_row:
                            continue
                            
                        # 카테고리 행 확인 (첫 번째 열에 값이 있고 나머지는 대부분 비어있음)
                        if not pd.isna(row[0]) and len(str(row[0]).strip()) > 0:
                            # 다른 열이 대부분 비어있으면 카테고리로 간주
                            non_empty_cells = sum(1 for x in row if not pd.isna(x))
                            if non_empty_cells <= 2:  # 카테고리 이름과 설명 정도만 있을 수 있음
                                current_category = str(row[0]).strip()
                                if current_category not in self.cheat_categories:
                                    self.cheat_categories[current_category] = []
                                self.log(f"카테고리 발견: '{current_category}'")
                                continue
                        
                        # 치트 항목 처리 (두 번째, 세 번째 열에 이름과 코드가 있음)
                        if not pd.isna(row[1]) and not pd.isna(row[2]):
                            # 현재 카테고리가 없으면 "기타" 카테고리에 추가
                            if current_category is None:
                                current_category = "기타"
                                if current_category not in self.cheat_categories:
                                    self.cheat_categories[current_category] = []
                                    
                            cheat_name = str(row[1]).strip()
                            cheat_code = str(row[2]).strip()
                            
                            # 사용 예시가 있으면 포함
                            example = ""
                            if len(row) > 3 and not pd.isna(row[3]):
                                example = str(row[3]).strip()
                            
                            # 치트 정보 구성
                            full_cheat = f"{cheat_name} — {cheat_code}"
                            if example:
                                full_cheat += f" — {example}"
                                
                            # 치트 데이터 추가
                            self.cheat_categories[current_category].append(full_cheat)
                            self.log(f"치트 추가: '{cheat_name}' -> '{cheat_code}'")
                    
                    # 엑셀 데이터 로드 성공
                    if self.cheat_categories and sum(len(cheats) for cheats in self.cheat_categories.values()) > 0:
                        self.log(f"엑셀에서 {len(self.cheat_categories)} 개의 카테고리와 {sum(len(cheats) for cheats in self.cheat_categories.values())} 개의 치트 로드됨")
                        self.use_excel_data = True
                        
                        # 카테고리 콤보박스 업데이트
                        self.category_combo['values'] = list(self.cheat_categories.keys())
                        
                        # 첫 번째 카테고리 선택
                        if len(self.cheat_categories) > 0:
                            first_category = list(self.cheat_categories.keys())[0]
                            self.category_combo.set(first_category)
                            self.select_category(first_category)
                            return
                    else:
                        self.log("엑셀 파일에서 유효한 치트 데이터를 찾을 수 없습니다.")
                        raise ValueError("치트 데이터가 없습니다.")
                        
                except Exception as e:
                    self.log(f"엑셀 파일 처리 중 오류 발생: {e}")
                    import traceback
                    self.log(traceback.format_exc())
                    raise
            else:
                self.log(f"오류: 치트 엑셀 파일을 찾을 수 없습니다: {CHEAT_FILE}")
                raise FileNotFoundError(f"엑셀 파일이 없습니다: {CHEAT_FILE}")
                
        except Exception as e:
            self.log(f"치트 데이터 로드 실패: {e}")
            import traceback
            self.log(traceback.format_exc())
            
            # 적어도 빈 카테고리라도 생성해서 프로그램이 실행될 수 있게 함
            self.log("빈 카테고리 생성 중...")
            self.cheat_categories = {"오류: 데이터 없음": ["치트 데이터 로드 실패 — 엑셀 파일 확인 필요"]}
            self.category_combo['values'] = list(self.cheat_categories.keys())
            self.category_combo.set("오류: 데이터 없음")
            self.select_category("오류: 데이터 없음")
    
    def execute_selected_cheat(self):
        """선택된 치트 실행 버튼 핸들러"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return
            
        # 치트 선택 확인
        selected_cheat_display = self.cheat_var.get()
        if not selected_cheat_display:
            self.log("경고: 실행할 치트를 선택해주세요.")
            return
        
        # 전체 치트 문자열 가져오기 (코드 포함)
        full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
            
        # 먼저 치트 메뉴 열기
        self.log("치트 메뉴 열기 시도 중...")
        if not self.open_cheat_menu():
            self.log("경고: 치트 메뉴를 열지 못했습니다.")
            return
        
        # "— GT." 문자열을 기준으로 코드 추출
        if " — GT." in full_cheat:
            cheat_code = full_cheat.split(" — ")[-1]
        else:
            cheat_code = full_cheat
        
        # 중괄호가 있는지 확인하고 파라미터 값 적용
        if '{' in cheat_code and '}' in cheat_code and self.param_entries:
            import re
            
            # 각 파라미터에 대해 입력된 값 적용
            for param, entry_var in self.param_entries.items():
                value = entry_var.get()
                if not value:  # 값이 비어있으면 알림
                    self.log(f"경고: '{param}' 값이 입력되지 않았습니다.")
                    if not messagebox.askyesno("파라미터 없음", f"'{param}' 값이 입력되지 않았습니다. 계속 진행하시겠습니까?"):
                        self.log("치트 실행이 취소되었습니다.")
                        return
                
                # 중괄호와 함께 파라미터를 사용자 입력으로 교체
                cheat_code = cheat_code.replace(f"{{{param}}}", value)
                self.log(f"파라미터 '{param}'에 '{value}' 값이 적용되었습니다.")
        
        # 치트 실행
        self.execute_cheat(cheat_code)
    
    def process_cheat_code_with_params(self, cheat_code):
        """중괄호({})가 포함된 치트 코드에서 사용자 입력을 받아 처리"""
        import re
        
        # 중괄호 내부의 파라미터 추출 (예: {RATE}, {VALUE} 등)
        params = re.findall(r'{([^}]+)}', cheat_code)
        
        if not params:
            return cheat_code
            
        self.log(f"치트 코드에 {len(params)}개의 파라미터가 필요합니다.")
        
        # 각 파라미터에 대해 사용자 입력 받기
        for param in params:
            param_prompt = f"'{param}' 값을 입력하세요:"
            user_input = simpledialog.askstring("파라미터 입력", param_prompt)
            
            if user_input is None:  # 사용자가 취소한 경우
                self.log(f"'{param}' 입력이 취소되었습니다.")
                return None
                
            # 중괄호와 함께 파라미터를 사용자 입력으로 교체
            cheat_code = cheat_code.replace(f"{{{param}}}", user_input)
            self.log(f"파라미터 '{param}'에 '{user_input}' 값이 입력되었습니다.")
        
        return cheat_code
    
    def open_cheat_menu(self):
        """치트 메뉴 열기 버튼 핸들러"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return False
        
        # 템플릿 매칭으로 메뉴 접근 시도
        self.log("템플릿 매칭으로 메뉴 접근 시도")
        
        # menu2의 존재 여부 확인
        menu2_result = self.find_image_on_screen('menu2.png')
        
        if menu2_result:
            self.log("menu2 발견: 치트 메뉴가 이미 열려있습니다.")
            # menu3 클릭
            menu3_result = self.find_image_on_screen('menu3.png')
            if menu3_result:
                pyautogui.click(menu3_result[0], menu3_result[1])
                self.log("menu3 클릭 완료 (이미지 매칭)")
                time.sleep(0.5)
                return True
            else:
                self.log("menu3를 찾을 수 없습니다.")
                return False
        else:
            # menu 클릭
            menu_result = self.find_image_on_screen('menu.png', report_max_val=True)
            if menu_result:
                pyautogui.click(menu_result[0], menu_result[1])
                self.log("menu 클릭 완료 (이미지 매칭)")
                time.sleep(0.5)
                
                # menu3 클릭
                menu3_result = self.find_image_on_screen('menu3.png')
                if menu3_result:
                    pyautogui.click(menu3_result[0], menu3_result[1])
                    self.log("menu3 클릭 완료 (이미지 매칭)")
                    time.sleep(0.5)
                    return True
                else:
                    self.log("menu3를 찾을 수 없습니다.")
                    return False
            else:
                self.log("menu를 찾을 수 없습니다.")
                return False
    
    def execute_cheat(self, cheat_code):
        """치트 실행 - 코드 입력 및 실행"""
        # 코드를 클립보드에 복사
        pyperclip.copy(cheat_code)
        self.log(f"치트 코드 '{cheat_code}' 복사됨")
        
        # 먼저 code2가 화면에 있는지 확인
        code2_result = self.find_image_on_screen('code2.png', report_max_val=True)
        if code2_result:
            self.log("code2 이미 존재함, code 버튼 클릭 단계 건너뜀")
            # code2 바로 클릭
            pyautogui.click(code2_result[0], code2_result[1])
            self.log(f"code2 클릭 완료 (이미지 매칭 위치: {code2_result[0]}, {code2_result[1]})")
        else:
            # code2가 없으면 일반적인 방법으로 진행
            if not self.click_button('code'):
                self.log("경고: code 버튼을 찾을 수 없습니다.")
                return False
            time.sleep(0.2)  # 대기 시간 변경
            
            if not self.click_button('code2'):
                self.log("경고: code2 버튼을 찾을 수 없습니다.")
                return False
        
        time.sleep(0.2)  # 대기 시간 변경
        
        # 코드 붙여넣기
        pyautogui.hotkey('ctrl', 'v')
        self.log("코드 붙여넣기 완료")
        time.sleep(0.2)  # 대기 시간 변경
        
        self.log("code3 클릭 시도...")
        if not self.click_button('code3'):
            self.log("경고: code3 버튼을 찾을 수 없습니다.")
            return False
        time.sleep(0.2)  # 대기 시간 변경
        
        # code5 먼저 클릭
        self.log("code5 클릭 시도...")
        code5_result = self.find_image_on_screen('code5.png', report_max_val=True)
        if code5_result:
            self.log(f"code5 발견! 위치: ({code5_result[0]}, {code5_result[1]})")
            pyautogui.click(code5_result[0], code5_result[1])
            self.log("code5 클릭 완료")
        else:
            self.log("code5 버튼을 찾을 수 없어 건너뜁니다.")
        
        time.sleep(0.2)  # 대기 시간 변경
        
        # 클릭 전 조금 더 대기
        self.log("code4 클릭 준비 중...")
        time.sleep(0.2)  # 대기 시간 변경
        
        # code4 버튼에 대해서는 이미지 검색과 클릭을 직접 처리
        code4_result = self.find_image_on_screen('code4.png', report_max_val=True)
        if code4_result:
            self.log(f"code4 발견! 위치: ({code4_result[0]}, {code4_result[1]})")
            # 두 번 클릭 시도
            pyautogui.click(code4_result[0], code4_result[1])
            time.sleep(0.2)  # 대기 시간 변경
            pyautogui.click(code4_result[0], code4_result[1])
            self.log("code4 클릭 완료 (두 번 시도)")
        else:
            self.log("경고: code4 버튼을 찾을 수 없습니다.")
            return False
        
        time.sleep(0.2)  # 대기 시간 변경
        self.log("치트 실행 완료")
        self.log("성공: 치트가 성공적으로 실행되었습니다.")
        return True

    def setup_log_tab(self):
        # 로그 탭 설정
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 로그 영역
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=30)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_area.config(state=tk.DISABLED)
        
        # 로그 제어 버튼
        button_frame = ttk.Frame(log_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        clear_log_btn = ttk.Button(button_frame, text="로그 지우기", command=self.clear_log)
        clear_log_btn.pack(side=tk.RIGHT, padx=5)
    
    def clear_log(self):
        """로그 영역 지우기"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.log("로그가 지워졌습니다.")
    
    def update_threshold(self, value):
        """임계값 업데이트"""
        self.threshold = float(value)
        self.threshold_label.config(text=f"{self.threshold:.1f}")
    
    def debug_templates(self):
        """템플릿 이미지 디버그"""
        if not self.window:
            self.log("경고: 윈도우를 먼저 선택하고 적용해주세요.")
            return
            
        self.log("=== 템플릿 디버그 시작 ===")
        
        # 각 템플릿 파일 테스트
        template_files = ['menu.png', 'menu2.png', 'menu3.png', 
                         'code.png', 'code2.png', 'code3.png', 'code4.png', 'code5.png']
        
        for template_file in template_files:
            # 템플릿 파일 확인
            template_path = os.path.join(TEMPLATES_DIR, template_file)
            if not os.path.exists(template_path):
                self.log(f"템플릿 파일 없음: {template_file}")
                continue
                
            # 이미지 매칭 시도
            result = self.find_image_on_screen(template_file, report_max_val=True)
            if result:
                self.log(f"템플릿 '{template_file}' 매칭 성공: 위치 ({result[0]}, {result[1]}), 정확도: {result[2]:.2f}")
            else:
                self.log(f"템플릿 '{template_file}' 매칭 실패")
        
        self.log("=== 템플릿 디버그 완료 ===")
    
    def log(self, message):
        """로그 영역에 메시지 추가"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        print(message)  # 콘솔에도 출력
    
    def get_window_list(self):
        """활성화된 윈도우 목록 가져오기"""
        try:
            # 모든 윈도우 목록 가져오기
            all_windows = pywinctl.getAllWindows()
            self.log(f"총 {len(all_windows)}개 윈도우 감지됨")
            
            # 보이는 창만 필터링 (타이틀이 있고 visible 속성이 True인 창)
            visible_windows = []
            for window in all_windows:
                title = window.title if hasattr(window, 'title') else ""
                is_visible = hasattr(window, 'visible') and window.visible
                
                # 타이틀이 있고 보이는 창만 추가
                if title and title.strip() and is_visible:
                    visible_windows.append(window)
            
            self.active_windows = visible_windows
            self.log(f"보이는 창 {len(self.active_windows)}개 필터링됨")
            
            # 리스트박스 업데이트
            self.window_titles = []
            self.window_listbox.delete(0, tk.END)
            
            for window in self.active_windows:
                title = window.title if hasattr(window, 'title') else str(window)
                if title.strip():  # 빈 타이틀은 제외
                    self.window_titles.append(title)
                    self.window_listbox.insert(tk.END, title)
            
            if not self.window_titles:
                self.log("활성화된 윈도우가 없습니다.")
                
        except Exception as e:
            self.log(f"오류: 윈도우 목록 가져오기 실패: {e}")
    
    def apply_selected_window_and_switch_tab(self):
        """윈도우를 선택하고 치트 탭으로 전환"""
        result = self.select_window()
        if result:
            # 선택된 윈도우 정보 업데이트
            self.window_info_label.config(text=f"선택된 윈도우: {self.window_titles[self.window_listbox.curselection()[0]]}")
            
            # 치트 카테고리 탭으로 전환
            self.tab_control.select(1)  # 두 번째 탭(index 1)으로 이동
            
            self.log("성공: 선택한 윈도우가 적용되었습니다.")
        else:
            self.log("경고: 윈도우 적용에 실패했습니다.")
    
    def select_window(self):
        """리스트박스에서 선택된 윈도우 활성화"""
        if not self.active_windows:
            self.log("활성화된 윈도우가 없습니다.")
            return False
        
        try:
            selected_indices = self.window_listbox.curselection()
            if not selected_indices:
                self.log("윈도우를 선택해주세요.")
                return False
                
            selected_index = selected_indices[0]
            if selected_index >= 0 and selected_index < len(self.window_titles):
                selected_title = self.window_titles[selected_index]
                
                # 타이틀로 윈도우 찾기
                for window in self.active_windows:
                    title = window.title if hasattr(window, 'title') else str(window)
                    if title == selected_title:
                        self.window = window
                        
                        # 윈도우 활성화
                        if hasattr(self.window, 'activate'):
                            self.window.activate()
                        else:
                            self.window.focus()  # pywinctl에서는 focus() 메서드를 사용
                        
                        self.log(f"'{selected_title}' 윈도우 선택됨")
                        time.sleep(0.2)  # 안정성을 위한 대기 (시간 변경)
                        return True
                
                self.log(f"선택한 윈도우를 찾을 수 없습니다: {selected_title}")
                return False
            else:
                self.log("윈도우를 선택해주세요.")
                return False
                
        except Exception as e:
            self.log(f"윈도우 선택 실패: {e}")
            return False
    
    def find_image_on_screen(self, template_name, threshold=None, report_max_val=False):
        """화면에서 이미지 찾기"""
        if threshold is None:
            threshold = self.threshold
            
        template_path = os.path.join(TEMPLATES_DIR, template_name)
        
        # 화면 캡처
        screenshot = pyautogui.screenshot()
        screenshot = np.array(screenshot)
        screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
        
        # 템플릿 이미지 로드
        template = cv2.imread(template_path)
        
        if template is None:
            self.log(f"템플릿 이미지를 찾을 수 없습니다: {template_path}")
            return None
        
        # 이미지 매칭
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # 매칭 결과 로깅
        if report_max_val:
            self.log(f"템플릿 '{template_name}' 매칭 값: {max_val:.2f}, 임계값: {threshold:.2f}")
        
        if max_val >= threshold:
            # 탐지된 위치의 중앙점 계산
            h, w = template.shape[:2]
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            return (center_x, center_y, max_val)
        
        return None
    
    def click_button(self, button_name):
        """버튼 클릭 - 이미지 매칭 방법만 사용"""
        # 템플릿 매칭으로 시도
        result = self.find_image_on_screen(f'{button_name}.png', report_max_val=True)
        if result:
            self.log(f"{button_name} 클릭 시도 (위치: {result[0]}, {result[1]})")
            pyautogui.click(result[0], result[1])
            time.sleep(0.2)  # 클릭 후 잠시 대기
            self.log(f"{button_name} 클릭 완료 (이미지 매칭 위치: {result[0]}, {result[1]})")
            return True
        
        self.log(f"{button_name} 버튼을 찾을 수 없습니다.")
        return False
    
    def on_cheat_selected(self, event):
        """치트가 선택되었을 때 호출되는 함수"""
        selected_cheat_display = self.cheat_var.get()
        if not selected_cheat_display:
            return
            
        self.log(f"치트 선택됨: '{selected_cheat_display}'")
            
        # 설명 업데이트
        self.update_description()
        
        # 파라미터 입력 필드 업데이트
        self.update_parameter_fields()
    
    def update_parameter_fields(self):
        """선택된 치트에 필요한 파라미터 입력 필드 생성"""
        # 기존 파라미터 입력 필드 삭제
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        self.param_entries.clear()
        
        selected_cheat_display = self.cheat_var.get()
        if not selected_cheat_display:
            return
            
        # 전체 치트 문자열 가져오기 (코드 포함)
        full_cheat = self.full_cheat_data.get(selected_cheat_display, selected_cheat_display)
        self.log(f"전체 치트 문자열: '{full_cheat}'")
        
        # 중괄호 안의 파라미터 추출
        import re
        params = []
        param_options = {}  # 파라미터별 옵션을 저장할 딕셔너리
        
        # 치트 코드 부분 (GT.로 시작하는 부분) 추출
        if " — GT." in full_cheat:
            # GT. 뒤의 모든 코드 부분 추출
            cheat_code_parts = full_cheat.split(" — GT.")
            for part in cheat_code_parts[1:]:  # 첫 번째는 이름이므로 건너뛰기
                # 각 코드 부분에서 중괄호 파라미터 검색
                param_matches = re.finditer(r'{([^}]+)}', part)
                for match in param_matches:
                    param_text = match.group(1)
                    
                    # 파이프(|)가 있는지 확인 (예: ON|OFF)
                    if '|' in param_text:
                        param_name = param_text.split('|')[0].split(':')[0].strip()
                        options = [opt.strip() for opt in param_text.split('|')]
                        # 파라미터 이름이 옵션에 포함되어 있으면 제거
                        if ':' in param_name:
                            param_name = param_name.split(':')[0].strip()
                            options[0] = options[0].split(':')[1].strip()
                        
                        params.append(param_name)
                        param_options[param_name] = options
                    else:
                        params.append(param_text)
        
        self.log(f"찾은 파라미터: {params}")
        
        if not params:
            self.log("중괄호 파라미터가 발견되지 않았습니다.")
            return
            
        # 중복 제거
        unique_params = []
        for p in params:
            if p not in unique_params:
                unique_params.append(p)
        params = unique_params
        self.log(f"중복 제거 후 파라미터: {params}")
            
        # 파라미터 라벨 추가
        ttk.Label(self.param_frame, text="파라미터:", font=("Arial", 9, "bold")).grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 각 파라미터에 대한 입력 필드 생성
        for i, param in enumerate(params):
            ttk.Label(self.param_frame, text=f"{param}:").grid(
                row=i+1, column=0, padx=5, pady=2, sticky=tk.W)
            
            # 토글/선택 옵션이 있는 경우 콤보박스 사용
            if param in param_options:
                self.log(f"파라미터 '{param}'에 옵션 선택 필드 생성: {param_options[param]}")
                combo_var = tk.StringVar()
                combo = ttk.Combobox(self.param_frame, textvariable=combo_var, 
                                  width=15, state="readonly")
                combo['values'] = param_options[param]
                combo.current(0)  # 첫 번째 옵션 선택
                combo.grid(row=i+1, column=1, padx=5, pady=2, sticky=tk.W)
                self.param_entries[param] = combo_var
            else:
                # 일반 입력 필드
                entry_var = tk.StringVar()
                entry = ttk.Entry(self.param_frame, textvariable=entry_var, width=30)
                entry.grid(row=i+1, column=1, padx=5, pady=2, sticky=tk.W)
                self.param_entries[param] = entry_var
            
        self.log(f"{len(params)}개의 파라미터 입력 필드가 생성되었습니다.")

# 메인 실행
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = GameCheaterGUI(root)
        root.mainloop()
    except KeyboardInterrupt:
        print("\n프로그램이 종료되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")
        messagebox.showerror("치명적 오류", f"프로그램 실행 중 오류 발생: {e}")
