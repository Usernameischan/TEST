import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import time
import win32gui
from PIL import Image
import os
from typing import Dict, Tuple, Optional
from dataclasses import dataclass
import logging
from ctypes import windll
from config import Config

@dataclass
class WindowRect:
    x: int
    y: int
    right: int
    bottom: int
    
    @property
    def width(self) -> int:
        return self.right - self.x
        
    @property
    def height(self) -> int:
        return self.bottom - self.y

class MapControlDetector:
    def __init__(self):
        """맵 컨트롤 디텍터 초기화"""
        self.config = Config()
        self._init_templates_dir()
        self.control_positions = self._get_default_control_positions()
        self.dpi_scale = self._get_dpi_scale()

    def _get_dpi_scale(self) -> float:
        """시스템의 DPI 스케일링 팩터를 반환"""
        try:
            if self.config.DPI_AWARE:
                awareness = windll.shcore.GetProcessDpiAwareness(0)
                if awareness != 2:  # DPI 인식 설정
                    windll.shcore.SetProcessDpiAwareness(2)
                dpi = windll.user32.GetDpiForSystem()
                return dpi / 96.0  # 기본 DPI는 96
        except Exception as e:
            logging.warning(f"DPI 스케일 획득 실패: {str(e)}")
        return 1.0

    @staticmethod
    def _init_templates_dir() -> None:
        """템플릿 디렉토리 초기화"""
        if not os.path.exists(Config.TEMPLATES_DIR):
            os.makedirs(Config.TEMPLATES_DIR)

    @staticmethod
    def _get_default_control_positions() -> Dict[str, Tuple[int, int]]:
        """기본 컨트롤 위치 정의"""
        return {
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

    def _adjust_coordinates(self, x: int, y: int) -> Tuple[int, int]:
        """DPI 스케일링과 윈도우 테두리 등을 고려한 좌표 보정"""
        adjusted_x = int((x - self.config.BORDER_OFFSET_X) * self.dpi_scale)
        adjusted_y = int((y - self.config.BORDER_OFFSET_Y) * self.dpi_scale)
        return adjusted_x, adjusted_y

    def verify_position(self, x: int, y: int, window_rect: WindowRect) -> bool:
        """좌표가 유효한 범위 내에 있는지 검증"""
        return (window_rect.x <= x <= window_rect.right and 
                window_rect.y <= y <= window_rect.bottom)

    def refine_position_with_image_recognition(self, 
                                             screenshot: Image, 
                                             x: int, 
                                             y: int, 
                                             template_name: str) -> Tuple[int, int]:
        """이미지 매칭을 통한 좌표 미세 조정"""
        try:
            if not self.config.USE_IMAGE_RECOGNITION:
                return x, y

            template_path = os.path.join(self.config.TEMPLATES_DIR, f"{template_name}.png")
            if not os.path.exists(template_path):
                return x, y

            screenshot_np = np.array(screenshot)
            template = cv2.imread(template_path, cv2.IMREAD_COLOR)
            
            search_area = screenshot_np[
                max(0, y-self.config.SEARCH_AREA_SIZE):min(screenshot_np.shape[0], y+self.config.SEARCH_AREA_SIZE),
                max(0, x-self.config.SEARCH_AREA_SIZE):min(screenshot_np.shape[1], x+self.config.SEARCH_AREA_SIZE)
            ]
            
            result = cv2.matchTemplate(search_area, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= self.config.RECOGNITION_THRESHOLD:
                refined_x = x - self.config.SEARCH_AREA_SIZE + max_loc[0]
                refined_y = y - self.config.SEARCH_AREA_SIZE + max_loc[1]
                return refined_x, refined_y
                
        except Exception as e:
            logging.error(f"이미지 인식 보정 실패: {str(e)}")
        return x, y

    def capture_map_window(self) -> Optional[WindowRect]:
        """맵 창의 위치 정보를 캡처"""
        try:
            formde_window = self._find_main_window()
            map_hwnd = self._find_child_window(formde_window)
            rect = win32gui.GetWindowRect(map_hwnd)
            return WindowRect(*rect)
        except Exception as e:
            logging.error(f"캡처 중 오류 발생: {str(e)}")
            return None

    def _find_main_window(self) -> int:
        """메인 윈도우 찾기"""
        formde_windows = gw.getWindowsWithTitle(self.config.WINDOW_TITLE)
        if not formde_windows:
            raise Exception(f"{self.config.WINDOW_TITLE} 창을 찾을 수 없습니다.")
        return win32gui.FindWindow(None, self.config.WINDOW_TITLE)

    def _find_child_window(self, parent_hwnd: int) -> int:
        """자식 윈도우 찾기"""
        child_windows = []
        
        def callback(child_hwnd, _):
            if self.config.CHILD_WINDOW_PATTERN in win32gui.GetWindowText(child_hwnd):
                child_windows.append(child_hwnd)
                
        win32gui.EnumChildWindows(parent_hwnd, callback, None)
        
        if not child_windows:
            raise Exception("맵 창을 찾을 수 없습니다.")
        return child_windows[0]

    def calculate_absolute_positions(self, window_rect: WindowRect) -> Dict[str, Tuple[int, int]]:
        """상대 좌표를 절대 좌표로 변환"""
        screenshot = pyautogui.screenshot(region=(
            window_rect.x, 
            window_rect.y, 
            window_rect.width, 
            window_rect.height
        ))
        
        absolute_positions = {}
        for control, (rel_x, rel_y) in self.control_positions.items():
            # 기본 좌표 계산
            abs_x = window_rect.x + rel_x
            abs_y = window_rect.y + rel_y
            
            # 좌표 보정 적용
            adjusted_x, adjusted_y = self._adjust_coordinates(abs_x, abs_y)
            
            # 이미지 인식을 통한 미세 조정
            refined_x, refined_y = self.refine_position_with_image_recognition(
                screenshot, adjusted_x, adjusted_y, control
            )
            
            # 좌표 검증
            if not self.verify_position(refined_x, refined_y, window_rect):
                logging.warning(f"컨트롤 '{control}'의 좌표가 창 범위를 벗어났습니다.")
                continue
                
            absolute_positions[control] = (refined_x, refined_y)
            logging.info(f"컨트롤 '{control}' 위치: ({refined_x}, {refined_y})")
            
        return absolute_positions

    def save_screenshot(self, window_rect: WindowRect) -> None:
        """스크린샷 저장"""
        screenshot = pyautogui.screenshot(region=(
            window_rect.x, 
            window_rect.y, 
            window_rect.width, 
            window_rect.height
        ))
        screenshot_path = os.path.join(self.config.TEMPLATES_DIR, self.config.SCREENSHOT_FILENAME)
        screenshot.save(screenshot_path)
        logging.info(f"맵 이미지가 {screenshot_path}로 저장되었습니다.")

    def save_positions(self, positions: Dict[str, Tuple[int, int]]) -> None:
        """좌표 정보 저장"""
        positions_path = os.path.join(self.config.TEMPLATES_DIR, self.config.POSITIONS_FILENAME)
        try:
            with open(positions_path, 'w', encoding='utf-8') as f:
                f.write("# 컨트롤 좌표 정보\n\n")
                f.write("CONTROL_POSITIONS = {\n")
                for control, pos in positions.items():
                    f.write(f"    '{control}': {pos},\n")
                f.write("}\n")
            logging.info(f"좌표 정보가 {positions_path} 파일에 저장되었습니다.")
        except Exception as e:
            logging.error(f"좌표 정보 저장 중 오류 발생: {str(e)}")

    def print_window_info(self, window_rect: WindowRect) -> None:
        """창 정보 출력"""
        logging.info("\n창 정보:")
        logging.info(f"Left: {window_rect.x}")
        logging.info(f"Top: {window_rect.y}")
        logging.info(f"Width: {window_rect.width}")
        logging.info(f"Height: {window_rect.height}")

def main():
    """메인 실행 함수"""
    try:
        # 로깅 설정
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        detector = MapControlDetector()
        
        # 맵 창 위치 찾기
        window_rect = detector.capture_map_window()
        if not window_rect:
            return
            
        # 스크린샷 저장
        detector.save_screenshot(window_rect)
        
        # 좌표 계산 및 저장 (이미지 인식 보정 포함)
        positions = detector.calculate_absolute_positions(window_rect)
        
        # 좌표 검증
        validated_positions = {
            control: pos for control, pos in positions.items()
            if detector.verify_position(*pos, window_rect)
        }
        
        detector.save_positions(validated_positions)
        detector.print_window_info(window_rect)
        
    except Exception as e:
        logging.error(f"예상치 못한 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())

if __name__ == "__main__":
    main()