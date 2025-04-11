import sys
import os
from datetime import datetime

# 'utils' 폴더 경로를 sys.path에 추가
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'utils')))

# 'kiwoom' 모듈 임포트
from kiwoom import Kiwoom

if __name__ == "__main__":
    kiwoom_instance = Kiwoom()
    try:
        print(f"[🕒 프로그램 실행 시작] {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        kiwoom_instance.run()
    except KeyboardInterrupt:
        print("\n[🔴 강제 종료 요청]")
        kiwoom_instance.shutdown()
        sys.exit()
