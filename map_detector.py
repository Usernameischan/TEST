import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import time
import win32gui
from PIL import Image
import os
import keyboard

# templates 폴더 생성
TEMPLATES_DIR = 'templates'
if not os.path.exists(TEMPLATES_DIR):
    os.makedirs(TEMPLATES_DIR)

class MapControlDetector:
    def __init__(self):
        # 컨트롤 버튼들의 상대 위치
        self.control_positions = {
            "5단계1": (40, 70),    
            "5단계2": (90, 70),
            "10단계": (140, 70),
            "기업": (190, 70),
            "뉴스": (230, 70),
            "Tick": (270, 70),
            "회원사": (320, 70),
            "PR": (370, 70),
            "가격대": (410, 70),
            "투자자": (460, 70),
            "외인": (510, 70),
            "News": (560, 70),
            "공식": (610, 70),
            "경쟁사": (660, 70),
            "ELW": (710, 70),
            "선물": (760, 70),
            "CFD": (810, 70)
        }

    def capture_map_window(self):
        """맵 창 캡처"""
        def enum_child_windows(hwnd, results):
            def callback(child_hwnd, _):
                text = win32gui.GetWindowText(child_hwnd)
                if any(pattern in text for pattern in ["모의실행"]):
                    results.append(child_hwnd)
            win32gui.EnumChildWindows(hwnd, callback, None)

        try:
            formde_windows = gw.getWindowsWithTitle("FORM Designer")
            if not formde_windows:
                raise Exception("FORM Designer 창을 찾을 수 없습니다.")

            formde_window = formde_windows[0]
            formde_hwnd = win32gui.FindWindow(None, "FORM Designer")
            
            child_windows = []
            enum_child_windows(formde_hwnd, child_windows)

            if not child_windows:
                raise Exception("맵 창을 찾을 수 없습니다.")

            map_hwnd = child_windows[0]
            rect = win32gui.GetWindowRect(map_hwnd)
            return rect  # (left, top, right, bottom)

        except Exception as e:
            print(f"캡처 중 오류 발생: {str(e)}")
            return None


### 상대 좌표 + 창의 시작 위치 = 절대 좌표
    def calculate_absolute_positions(self, window_rect):
        """상대 좌표를 절대 좌표로 변환"""
        if not window_rect:
            return {}

        base_x, base_y = window_rect[0], window_rect[1]
        absolute_positions = {}

        for control, rel_pos in self.control_positions.items():
            abs_x = base_x + rel_pos[0]
            abs_y = base_y + rel_pos[1]
            absolute_positions[control] = (abs_x, abs_y)
            print(f"컨트롤 '{control}' 위치: ({abs_x}, {abs_y})")

        return absolute_positions

    def measure_exact_position(self):
        """마우스 위치를 정확하게 측정"""
        print("\n컨트롤 위치 정밀 측정을 시작합니다.")
        print("각 컨트롤의 정확한 중앙을 클릭해주세요.")
        print("ESC를 누르면 측정을 종료합니다.")
        print("Space를 누르면 현재 마우스 위치를 확인할 수 있습니다.\n")
        print("2초 후에 측정이 시작됩니다...")
        
        # 2초 대기 추가
        for i in range(2, 0, -1):
            print(f"{i}초...")
            time.sleep(1)
        print("측정을 시작합니다!\n")

        measured_positions = {}
        
        for control_name in self.control_positions.keys():
            print(f"\n'{control_name}' 컨트롤의 위치를 클릭해주세요...")
            
            try:
                # 사용자가 클릭할 때까지 대기
                while True:
                    if keyboard.is_pressed('esc'):
                        print("\n측정이 취소되었습니다.")
                        return None
                        
                    if keyboard.is_pressed('space'):
                        x, y = pyautogui.position()
                        print(f"현재 마우스 위치: ({x}, {y})")
                        time.sleep(0.2)  # 스페이스바 연속 입력 방지
                        
                    if pyautogui.mouseDown():
                        x, y = pyautogui.position()
                        measured_positions[control_name] = (x, y)
                        print(f"측정된 위치: ({x}, {y})")
                        time.sleep(0.5)  # 더블클릭 방지
                        break
                        
                    time.sleep(0.1)
                    
            except KeyboardInterrupt:
                print("\n측정이 취소되었습니다.")
                return None
        
        return measured_positions

def main():
    try:
        detector = MapControlDetector()
        
        # 수동 측정 모드 선택
        response = input("컨트롤 위치를 수동으로 측정하시겠습니까? (y/n): ")
        if response.lower() == 'y':
            positions = detector.measure_exact_position()
            if positions:
                detector.control_positions = positions
        
        # 맵 창 위치 찾기
        window_rect = detector.capture_map_window()
        if not window_rect:
            return
            
        # 캡처 및 저장
        x, y, right, bottom = window_rect
        width = right - x
        height = bottom - y
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        
        # templates 폴더에 이미지 저장
        screenshot_path = os.path.join(TEMPLATES_DIR, "captured_map.png")
        screenshot.save(screenshot_path)
        print(f"맵 이미지가 {screenshot_path}로 저장되었습니다.")
        
        # 절대 좌표 계산
        positions = detector.calculate_absolute_positions(window_rect)
        
        # 좌표 정보 파일도 templates 폴더에 저장
        positions_path = os.path.join(TEMPLATES_DIR, 'control_positions.py')
        try:
            with open(positions_path, 'w', encoding='utf-8') as f:
                f.write("# 컨트롤 좌표 정보\n\n")
                f.write("CONTROL_POSITIONS = {\n")
                for control, pos in positions.items():
                    f.write(f"    '{control}': {pos},\n")
                f.write("}\n")
            print(f"좌표 정보가 {positions_path} 파일에 저장되었습니다.")
        except Exception as e:
            print(f"좌표 정보 저장 중 오류 발생: {str(e)}")
        
        print("\n창 정보:")
        print(f"Left: {x}")
        print(f"Top: {y}")
        print(f"Width: {width}")
        print(f"Height: {height}")
        
    except KeyboardInterrupt:
        print("\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"예상치 못한 오류 발생: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()