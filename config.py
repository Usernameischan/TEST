from dataclasses import dataclass

@dataclass
class Config:
    # 기본 설정
    TEMPLATES_DIR: str = 'templates'
    WINDOW_TITLE: str = "FORM Designer"
    CHILD_WINDOW_PATTERN: str = "모의실행"
    SCREENSHOT_FILENAME: str = "captured_map.png"
    POSITIONS_FILENAME: str = 'control_positions.py'

    # 좌표 보정 관련 설정
    DPI_AWARE: bool = True
    BORDER_OFFSET_X: int = 8
    BORDER_OFFSET_Y: int = 31
    
    # 이미지 인식 관련 설정
    USE_IMAGE_RECOGNITION: bool = True
    RECOGNITION_THRESHOLD: float = 0.8
    SEARCH_AREA_SIZE: int = 20