import pyautogui
import time
import json
import logging
import os
from pathlib import Path
import win32gui
import pygetwindow as gw
import keyboard

# templates 폴더 경로 설정
TEMPLATES_DIR = 'templates'
CONTROL_POSITIONS_PATH = os.path.join(TEMPLATES_DIR, 'control_positions.py')

# control_positions.py 파일이 없을 경우 에러 처리
if not os.path.exists(CONTROL_POSITIONS_PATH):
    raise FileNotFoundError(f"'{CONTROL_POSITIONS_PATH}' 파일을 찾을 수 없습니다. map_detector.py를 먼저 실행해주세요.")

# control_positions.py 파일에서 좌표 정보 로드
import sys
sys.path.append(TEMPLATES_DIR)
from control_positions import CONTROL_POSITIONS

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('control_clicker.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class ControlClicker:
    def __init__(self):
        self.load_config()
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = self.config['click_delay']
        self.positions = CONTROL_POSITIONS

    def load_config(self):
        """설정 파일 로드"""
        default_config = {
            'click_delay': 0.5,
            'max_retries': 3,
            'retry_delay': 1.0,
            'move_speed': 0.2
        }
        
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                logging.info("설정 파일을 로드했습니다.")
            else:
                self.config = default_config
                with open('config.json', 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, indent=4)
                logging.info("기본 설정 파일을 생성했습니다.")
        except Exception as e:
            logging.error(f"설정 파일 로드 중 오류: {str(e)}")
            self.config = default_config

    def validate_position(self, position):
        """좌표 유효성 검사"""
        try:
            x, y = position
            screen_width, screen_height = pyautogui.size()
            return 0 <= x <= screen_width and 0 <= y <= screen_height
        except Exception as e:
            logging.error(f"좌표 검증 중 오류: {str(e)}")
            return False

    def check_window_exists(self):
        """FORM Designer 창이 존재하는지 확인"""
        try:
            formde_windows = gw.getWindowsWithTitle("FORM Designer")
            return bool(formde_windows)
        except Exception as e:
            logging.error(f"창 확인 중 오류: {str(e)}")
            return False

    def safe_click(self, position, control_name):
        """정밀한 클릭 수행"""
        for attempt in range(self.config['max_retries']):
            try:
                if not self.validate_position(position):
                    logging.warning(f"유효하지 않은 좌표: {control_name} - {position}")
                    return False
                
                # 마우스를 천천히 정확한 위치로 이동
                x, y = position
                pyautogui.moveTo(x, y, duration=0.2)
                
                # 미세 조정을 위한 주변 픽셀 확인
                for offset_x, offset_y in [(0,0), (1,0), (-1,0), (0,1), (0,-1)]:
                    check_x = x + offset_x
                    check_y = y + offset_y
                    # 여기서 필요한 경우 픽셀 색상 등을 확인할 수 있습니다
                
                # 클릭 전 잠시 대기
                time.sleep(0.1)
                
                # 정밀 클릭 수행
                pyautogui.click(x, y)
                
                # 클릭 확인을 위한 대기
                time.sleep(0.1)
                
                logging.info(f"'{control_name}' 클릭 성공 ({x}, {y})")
                return True
                
            except Exception as e:
                logging.warning(f"'{control_name}' 클릭 실패 (시도 {attempt + 1}/{self.config['max_retries']}): {str(e)}")
                time.sleep(self.config['retry_delay'])
        
        return False

    def show_progress(self, current, total):
        """진행 상황 표시"""
        percent = (current / total) * 100
        print(f"\r진행률: {percent:.1f}% ({current}/{total})", end="", flush=True)

    def click_controls(self):
        """컨트롤 클릭 실행"""
        if not self.check_window_exists():
            logging.error("FORM Designer 창을 찾을 수 없습니다.")
            return

        logging.info("컨트롤 클릭을 시작합니다...")
        print("중지하려면 마우스를 화면 왼쪽 상단 모서리로 이동하세요.")
        print("각 클릭 사이에 'space'를 누르면 일시 정지합니다.")

        try:
            total_controls = len(self.positions)
            successful_clicks = 0

            for i, (control_name, position) in enumerate(self.positions.items(), 1):
                # 현재 마우스 위치 저장
                current_pos = pyautogui.position()
                
                # 사용자 일시 정지 확인
                if keyboard.is_pressed('space'):
                    input("일시 정지됨. 계속하려면 Enter를 누르세요...")
                
                # 정밀 클릭 수행
                if self.safe_click(position, control_name):
                    successful_clicks += 1
                
                # 원래 마우스 위치로 복귀
                pyautogui.moveTo(current_pos)
                
                # 진행 상황 표시
                self.show_progress(i, total_controls)
                
                # 대기
                time.sleep(self.config['click_delay'])

            print("\n클릭 완료 통계:")
            print(f"총 컨트롤 수: {total_controls}")
            print(f"성공한 클릭: {successful_clicks}")
            print(f"실패한 클릭: {total_controls - successful_clicks}")
            
            logging.info(f"클릭 작업 완료 (성공: {successful_clicks}, 실패: {total_controls - successful_clicks})")

        except pyautogui.FailSafeException:
            logging.info("사용자에 의해 중지되었습니다.")
            print("\n사용자에 의해 중지되었습니다.")
        except Exception as e:
            logging.error(f"오류 발생: {str(e)}")
            print(f"\n오류 발생: {str(e)}")

def main():
    try:
        clicker = ControlClicker()
        
        # 시작 전 확인
        print("3초 후 클릭을 시작합니다...")
        print("준비하세요...")
        time.sleep(3)
        
        # 컨트롤 클릭 실행
        clicker.click_controls()
        
    except KeyboardInterrupt:
        logging.info("프로그램이 사용자에 의해 중단되었습니다.")
        print("\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        logging.error(f"예상치 못한 오류 발생: {str(e)}")
        print(f"예상치 못한 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()