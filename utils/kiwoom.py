# -----------------------------------
# 🔵 1. 필수 라이브러리 로드
# -----------------------------------
import os
import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import configparser
from datetime import datetime, time as dtime
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QEventLoop, QTimer
from PyQt5.QAxContainer import QAxWidget
from openpyxl import Workbook
from PyQt5.QtWidgets import QMessageBox

# -----------------------------------
# 🔵 2. Kiwoom 클래스 정의 (메인)
# -----------------------------------
class Kiwoom:
    def __init__(self):
        """프로그램 초기화 (PyQt, API 연결, 설정 파일 로드, 내부 변수 초기화)"""
        print("[🟢 프로그램 초기화 중...]")

        # PyQt5 애플리케이션 생성 (필수)
        self.app = QApplication(sys.argv)

        # 키움 API 연결 객체 생성
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # 키움 이벤트 핸들러 등록
        self.ocx.OnEventConnect.connect(self._on_login)
        self.ocx.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.ocx.OnReceiveRealData.connect(self._on_receive_real_data)
        self.ocx.OnReceiveChejanData.connect(self._on_receive_chejan_data)

        # 설정 파일 로드 (config.ini)
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        # 사용자 설정값 저장
        self.account_pw = config['USER']['account_pw']
        self.max_profit_rate = float(config['TRADING']['max_profit_rate'])
        self.max_loss_rate = float(config['TRADING']['max_loss_rate'])
        self.max_holding_count = int(config['TRADING']['max_holding_count'])
        self.target_stocks = eval(config['TRADING']['target_list'])
        self.max_stock_ratio = float(config['TRADING']['max_stock_ratio'])
        self.buy_split_count = int(config['TRADING']['buy_split_count'])
        self.restart_after_close = config.getboolean('TRADING', 'restart_after_close')

        # 내부 상태 변수
        self.account_number = None
        self.available_cash = 0
        self.login_event_loop = None
        self.tr_event_loop = None
        self.macd_data = {}
        self.own_stocks = {}
        self.trade_log = []
        self.logged_realtime_codes = set()
        self.current_screen_no = 2000
        self.screen_by_code = {}
        self.real_time_success = False     # 실시간 등록 성공 여부
        self.daily_chart_success = False   # 일봉 데이터 수신 성공 여부

        # 장 상태 체크 타이머
        self.check_timer = QTimer()
        self.check_timer.timeout.connect(self.check_market_status)

        # 잔액 조회 타이머
        self.balance_timer = QTimer()
        self.balance_timer.timeout.connect(self.check_balance)

        print("[✅ 프로그램 초기화 완료]")

# -----------------------------------
# 🔵 3. 로그인 처리
# -----------------------------------
    def login(self):
        """키움 서버 로그인 요청"""
        print("[🔐 로그인 요청 중...]")
        self.ocx.dynamicCall("CommConnect()")   # 키움 로그인창 띄우기
        self.login_event_loop = QEventLoop()    # 로그인 완료될 때까지 대기
        self.login_event_loop.exec_()

    def _on_login(self, err_code):
        """로그인 완료 이벤트 수신"""
        if err_code == 0:
            print("[✅ 로그인 성공]")
            self.account_number = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO").split(';')[0]
            # 서버 종류 체크
            server_type = self.ocx.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
            if server_type == "1":
                print("[🧪 모의투자 서버 접속 감지]")
            else:
                print("[🏦 실서버 접속 감지]")
        else:
            print(f"[❌ 로그인 실패] 에러코드: {err_code}")
        self.login_event_loop.exit()

# -----------------------------------
# 🔵 4. 잔액 조회 (초기 현금 확보)
# -----------------------------------
    def check_balance(self):
        """계좌 잔액 조회 요청"""
        print("[💰 초기에 잔액 조회 요청 중...]")
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_number)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.account_pw)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opw00018_req", "opw00018", 0, "2000")
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next):
        """TR 데이터 수신 이벤트"""
        if rqname == "opw00018_req":
            self.handle_balance(trcode, rqname)
        elif rqname == "opt10081_req":
            self.handle_daily_chart(trcode, rqname, screen_no)

    def handle_balance(self, trcode, rqname):
        """잔액 조회 결과 저장"""
        try:
            cash_raw = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, 0, "출금가능금액").strip()
            if cash_raw and cash_raw.lstrip('-').isdigit():
                self.available_cash = abs(int(cash_raw))
                print(f"[💰 현재 출금 가능 금액]: {self.available_cash:,}원")
            else:
                self.available_cash = 0
                print("[⚠️ 출금 가능 금액 없음]")
        except Exception as e:
            print(f"[❌ 잔액 조회 실패]: {e}")
            self.save_error_log(str(e))
        finally:
            self.tr_event_loop.exit()

# -----------------------------------
# 🔵 5. 관심 종목 일봉 데이터 요청
# -----------------------------------
    def request_daily_chart(self, code):
        """특정 종목 코드에 대해 일봉 데이터 요청"""
        print(f"[📈 {code}] 일봉 데이터 요청")
        self.current_screen_no += 1
        screen_no = str(self.current_screen_no)
        self.screen_by_code[code] = screen_no

        today = datetime.now().strftime("%Y%m%d")
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "기준일자", today)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10081_req", "opt10081", 0, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

# -----------------------------------
# 🔵 6. 일봉 데이터 수신 및 분석
# -----------------------------------
    def handle_daily_chart(self, trcode, rqname, screen_no):
        """서버로부터 받은 일봉 데이터 처리 (MACD, EMA, 5일 이평선 계산)"""
        print(f"[📥 {screen_no}] 일봉 데이터 수신 처리 시작")

        try:
            count = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
            closes = []

            # 종가 데이터 수집
            for i in range(count):
                close = self.ocx.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", trcode, rqname, i, "현재가"
                ).strip()
                if close and close.lstrip('-').isdigit():
                    closes.append(abs(int(close)))

            if len(closes) < 50:
                print(f"[⚠️ {screen_no}] 데이터 부족: {len(closes)}개 → 종목 제외")
                self.tr_event_loop.exit()
                return

            closes.reverse()  # 최신순 → 과거순 변환
            closes = pd.Series(closes)

            # 기술적 지표 계산
            ema5 = closes.ewm(span=5, adjust=False).mean()
            ema12 = closes.ewm(span=12, adjust=False).mean()
            ema26 = closes.ewm(span=26, adjust=False).mean()
            macd_line = ema12 - ema26
            signal_line = macd_line.ewm(span=9, adjust=False).mean()

            code = [k for k, v in self.screen_by_code.items() if v == screen_no][0]

            # 종목별 데이터 저장
            self.macd_data[code] = {
                "ema5": ema5,
                "macd": macd_line,
                "signal": signal_line,
                "closes": closes
            }
            self.daily_data_success = True  # 일봉 데이터 수신 성공 기록

        except Exception as e:
            self.save_error_log(str(e))
            print(f"[❌ 일봉 데이터 처리 실패]: {e}")
            self.daily_data_success = False  # 실패 기록
        finally:
            self.tr_event_loop.exit()

# -----------------------------------
# 🔵 7. 실시간 체결 감시 등록
# -----------------------------------
    def start_real_time_monitoring(self):
        """관심 종목에 대해 실시간 체결 감시 등록 (10개씩 나눠서 등록)"""
        print("[📡 실시간 체결 감시 등록 시작]")

        fids = "10"  # 체결가격 FID
        codes_list = list(self.target_stocks.keys())
        batch_size = 10  # 한 번에 10종목씩 등록
        self.real_time_success = True

        for idx in range(0, len(codes_list), batch_size):
            batch_codes = codes_list[idx:idx + batch_size]
            code_str = ";".join(batch_codes)
            screen_no = str(5000 + idx // batch_size)

            try:
                self.ocx.dynamicCall(
                    "SetRealReg(QString, QString, QString, QString)",
                    screen_no, code_str, fids, "0"
                )
                print(f"[✅ 실시간 등록 완료] 화면번호 {screen_no} → 종목 {batch_codes}")
            except Exception as e:
                self.save_error_log(str(e))
                print(f"[⚠️ 실시간 등록 실패]: {e}")
                self.real_time_success = False

# -----------------------------------
# 🔵 8. 매수 조건 판단
# -----------------------------------
    def predict_trading(self, code):
        """성공/실패 상황에 따라 매수 전략 선택"""
        print(f"[⚙️ {code}] 매수 전략 판단 시작")

        # 일봉+실시간 데이터가 모두 성공한 경우 → 기존 MACD 전략 사용
        if self.daily_data_success and self.real_time_success:
            self.predict_by_macd_strategy(code)
        else:
            # 실패 시 → 5일 이평선 돌파 전략 사용
            self.predict_by_ema5_breakout(code)

    def predict_by_macd_strategy(self, code):
        """MACD + Signal Line 전략"""
        print(f"[🔵 {code}] MACD 전략 적용")

        data = self.macd_data.get(code)
        if not data:
            print(f"[⚠️ {code}] 데이터 없음")
            return

        macd_now = data["macd"].iloc[-1]
        macd_prev = data["macd"].iloc[-2]
        signal_now = data["signal"].iloc[-1]
        signal_prev = data["signal"].iloc[-2]

        is_golden_cross = macd_prev < signal_prev and macd_now > signal_now

        if is_golden_cross:
            print(f"[🌟 {code}] MACD 골든크로스 감지 → 매수 시도")
            self.try_buy(code)
        else:
            print(f"[⚪ {code}] MACD 골든크로스 없음 → 매수 보류")

    def predict_by_ema5_breakout(self, code):
        """5일 이평선 돌파 전략"""
        print(f"[🟡 {code}] 5일 이평선 돌파 전략 적용")

        data = self.macd_data.get(code)
        if not data:
            print(f"[⚠️ {code}] 데이터 없음")
            return

        close_today = data["closes"].iloc[-1]
        ema5_today = data["ema5"].iloc[-1]

        # 오늘 종가가 5일 이평선 돌파
        if close_today > ema5_today:
            print(f"[🌟 {code}] 종가 5일선 돌파 감지 → 매수 시도")
            self.try_buy(code)
        else:
            print(f"[⚪ {code}] 5일선 돌파 아님 → 매수 보류")

    def _on_receive_real_data(self, code, real_type, real_data):
        """실시간 체결 데이터 수신 이벤트"""
        if real_type != "주식체결":
            return

        try:
            price_raw = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 10).strip()
            price = abs(int(price_raw))
        except (ValueError, AttributeError) as e:
            print(f"[❌ 실시간 데이터 변환 에러] 종목: {code} / 에러: {e}")
            return

        if code not in self.own_stocks:
            # 보유 안 한 종목 → 매수 판단
            self.predict_trading(code)
        else:
            # 보유한 종목 → 매도 판단
            self.try_sell(code, price)

    def _on_receive_chejan_data(self, gubun, item_cnt, fid_list):
        """체결/잔고 데이터 수신 이벤트"""
        print(f"[📩 체결 데이터 수신] 구분: {gubun} / 항목 수: {item_cnt}")

        if gubun == "0":  # 0: 주문체결
            code = self.ocx.dynamicCall("GetChejanData(int)", 9001).strip()  # 종목코드
            order_status = self.ocx.dynamicCall("GetChejanData(int)", 913).strip()  # 주문상태
            filled_qty = self.ocx.dynamicCall("GetChejanData(int)", 911).strip()  # 체결수량
            price = self.ocx.dynamicCall("GetChejanData(int)", 910).strip()  # 체결가격
            print(f"[체결완료] {code} / 상태: {order_status} / 체결수량: {filled_qty} / 체결가격: {price}")

    def check_market_status(self):
        """장 종료 감지 후 프로그램 종료"""
        now = datetime.now().time()
        if now < dtime(8, 0) or now > dtime(18, 0):
            print("[🚪 장 종료 감지] 프로그램 종료 시작")
            self.shutdown()


# -----------------------------------
# 🔵 9. 매수 실행 (분할 매수 + 투자비율 제한)
# -----------------------------------
    def try_buy(self, code):
        """실제 매수 실행 (분할 매수 + 종목당 투자비율 제한)"""
        if code in self.own_stocks:
            print(f"[🚫 {code}] 이미 보유 중 → 추가 매수 금지")
            return

        # 현재가 조회
        raw_price = self.ocx.dynamicCall("GetMasterLastPrice(QString)", code)
        try:
            price = abs(int(raw_price.strip()))
        except Exception as e:
            self.save_error_log(str(e))
            print(f"[❌ {code}] 현재가 조회 실패: {e}")
            return

        # 투자 금액 계산
        if self.available_cash < price:
            print(f"[⚠️ {code}] 잔액 부족 → 매수 불가")
            return

        max_invest_amount = self.available_cash * (self.max_stock_ratio / 100)
        split_amount = max_invest_amount / self.buy_split_count

        print(f"[🛒 {code}] 최대 {max_invest_amount:,.0f}원 / 1회 {split_amount:,.0f}원 매수 시작")

        total_quantity = 0

        for i in range(self.buy_split_count):
            if self.available_cash < split_amount:
                print(f"[⚠️ {code}] 잔액 부족 → {i+1}회차 중단")
                break

            quantity = int(split_amount // price)
            if quantity < 1:
                print(f"[⚠️ {code}] {i+1}회차 매수 실패 (수량 0)")
                continue

            self.send_order(code, 1, quantity)
            self.available_cash -= quantity * price
            total_quantity += quantity

            print(f"[🛒 {code}] {i+1}회차 {quantity}주 매수 완료")

        if total_quantity > 0:
            self.own_stocks[code] = {
                "buy_price": price,
                "quantity": total_quantity,
                "highest_price": price
            }
            print(f"[✅ {code}] 총 {total_quantity}주 매수 완료")
        else:
            print(f"[⚠️ {code}] 최종 매수 실패")

# -----------------------------------
# 🔵 10. 매도 조건 체크 (손익/손절 우선)
# -----------------------------------
    def try_sell(self, code, current_price):
        """매도 조건 체크 (손익 우선, 손절 우선)"""
        stock = self.own_stocks.get(code)
        if not stock:
            print(f"[🚫 {code}] 보유하지 않음 → 매도 무시")
            return

        buy_price = stock['buy_price']
        quantity = stock['quantity']
        highest_price = stock['highest_price']

        if current_price > highest_price:
            stock['highest_price'] = current_price
            highest_price = current_price

        profit_rate = self.calculate_profit_rate(buy_price, current_price)
        trailing_stop_price = highest_price * 0.97

        if profit_rate >= self.max_profit_rate:
            print(f"[🚀 {code}] 목표 수익률 도달 → 매도")
            self.show_alert(f"[익절] {code} 수익률 {profit_rate:.2f}% 도달!")
            self._sell_stock(code, quantity, current_price)
        elif profit_rate <= self.max_loss_rate:
            print(f"[🛑 {code}] 손절 기준 도달 → 매도")
            self.show_alert(f"[손절] {code} 수익률 {profit_rate:.2f}% 도달!")
            self._sell_stock(code, quantity, current_price)
        elif current_price < trailing_stop_price:
            print(f"[🚨 {code}] 트레일링 스탑 발동 → 매도")
            self.show_alert(f"[트레일링 스탑] {code} 가격 하락 → 매도")
            self._sell_stock(code, quantity, current_price)
        else:
            print(f"[⚪ {code}] 매도 조건 미충족 (수익률 {profit_rate:.2f}%)")


# -----------------------------------
# 🔵 11. 실제 매도 실행
# -----------------------------------
    def _sell_stock(self, code, quantity, price):
        """매도 주문 실행 및 매매 기록"""
        print(f"[📈 {code}] 매도 {quantity}주 @ {price}원")
        self.send_order(code, 2, quantity)
        self.record_trade(code, "매도", quantity, price)
        self.own_stocks.pop(code, None)
        self.check_balance()  # 매도 후 잔액 재조회
        
    def record_trade(self, code, trade_type, quantity, price):
        """매매 기록 추가"""
        date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.trade_log.append([date, code, trade_type, quantity, price])
        print(f"[📝 매매 기록 추가] {date} / {code} / {trade_type} / {quantity}주 / {price}원")


# -----------------------------------
# 🔵 12. 주문 전송 함수 (공통)
# -----------------------------------
    def send_order(self, code, order_type, quantity):
        """키움 서버에 주문 전송 (1: 매수, 2: 매도)"""
        order_type_str = 1 if order_type == 1 else 2  # 1: 신규매수, 2: 신규매도

        res = self.ocx.dynamicCall(
            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
            "주문", "5000", self.account_number, order_type_str, code,
            quantity, 0, "03", ""  # 03: 시장가 주문
        )

        if res == 0:
            print(f"[✅ 주문 성공] {code} {quantity}주 {'매도' if order_type == 2 else '매수'}")
        else:
            print(f"[❌ 주문 실패] {code} (결과 코드: {res})")


# -----------------------------------
# 🔵 13. 수익률 계산
# -----------------------------------
    def calculate_profit_rate(self, buy_price, current_price):
        """수익률 계산 함수"""
        try:
            return ((current_price - buy_price) / buy_price) * 100
        except ZeroDivisionError:
            return 0

# -----------------------------------
# 🔵 14. 매매 기록 저장 (엑셀로 저장)
# -----------------------------------
    def save_trade_log(self):
        """매매 기록을 엑셀 파일로 저장"""
        print("[💾 매매 기록 저장 시도]")

        # logs 폴더 자동 생성
        log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
        os.makedirs(log_dir, exist_ok=True)

        # 엑셀 워크북 생성
        wb = Workbook()
        ws = wb.active
        ws.title = "Trade Log"

        # 첫 번째 행 제목
        ws.append(["날짜", "종목코드", "매매구분", "수량", "가격"])

        # 매매 기록 추가
        for log in self.trade_log:
            ws.append(log)

        # 파일 저장
        file_path = os.path.join(log_dir, f"trade_log_{datetime.now().strftime('%Y%m%d')}.xlsx")
        wb.save(file_path)

        print(f"[✅ 매매 기록 저장 완료]: {file_path}")

# -----------------------------------
# 🔵 15. 수익률 그래프 저장
# -----------------------------------
    def draw_profit_graph(self):
        """매매 결과를 바탕으로 수익률 그래프를 그리고 저장"""
        print("[📊 수익률 그래프 그리기 시작]")

        # 매도한 거래만 집계
        profits = [(log[0][:10], float(log[4])) for log in self.trade_log if log[2] == "매도"]

        if not profits:
            print("[⚠️ 수익 데이터 없음 → 그래프 생략]")
            return

        # 데이터프레임 생성
        df = pd.DataFrame(profits, columns=["date", "profit"])
        df_grouped = df.groupby("date").sum().reset_index()

        plt.figure(figsize=(10, 6))
        plt.plot(df_grouped['date'], df_grouped['profit'], marker='o', label='일별 수익')
        plt.axhline(y=self.max_profit_rate, color='green', linestyle='--', label=f'익절 목표 {self.max_profit_rate}%')
        plt.axhline(y=self.max_loss_rate, color='red', linestyle='--', label=f'손절 기준 {self.max_loss_rate}%')

        plt.title('📈 일별 수익률 그래프')
        plt.xlabel('날짜')
        plt.ylabel('수익(원)')
        plt.legend()
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()

        # 저장
        graph_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
        os.makedirs(graph_dir, exist_ok=True)
        file_path = os.path.join(graph_dir, f"profit_graph_{datetime.now().strftime('%Y%m%d')}.png")
        plt.savefig(file_path)

        print(f"[✅ 수익률 그래프 저장 완료]: {file_path}")

# -----------------------------------
# 🔵 16. 실시간 감시 해제
# -----------------------------------
    def stop_real_time_monitoring(self):
        """프로그램 종료 전 실시간 체결 감시 해제"""
        try:
            self.ocx.dynamicCall("DisconnectRealData(QString)", "5000")
            print("[🛑 실시간 체결 감시 해제 완료]")
        except Exception as e:
            self.save_error_log(str(e))
            print(f"[⚠️ 실시간 감시 해제 중 에러]: {e}")

    def show_alert(self, message):
        """PyQt5 알림창"""
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("📢 알림")
        msg_box.setText(message)
        msg_box.exec_()
        
    def save_error_log(self, error_message):
        """에러 메시지를 파일로 저장"""
        log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
        os.makedirs(log_dir, exist_ok=True)
        file_path = os.path.join(log_dir, f"error_log_{datetime.now().strftime('%Y%m%d')}.txt")
        
        with open(file_path, 'a', encoding='utf-8') as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{now}] {error_message}\n")

        print(f"[⚠️ 에러 저장 완료]: {file_path}")

    
# -----------------------------------
# 🔵 17. 프로그램 안전 종료
# -----------------------------------
    def shutdown(self):
        """
        프로그램 안전 종료 처리
        - 매매 기록 저장
        - 수익률 그래프 저장
        - 실시간 감시 해제
        - PyQt 애플리케이션 종료
        """
        print("[🛑 프로그램 종료 - 매매 기록 저장 중...]")
        try:
            self.save_trade_log()
            self.draw_profit_graph()
            self.stop_real_time_monitoring()
            print("[✅ 매매 기록 저장, 그래프 저장, 감시 해제 완료]")
        except Exception as e:
            self.save_error_log(str(e))
            print(f"[❌ 종료 중 에러 발생]: {e}")

        self.app.quit()
        print("[✅ 프로그램 완전 종료]")


    def run(self):
        """로그인 → 관심 종목 일봉 조회 → 실시간 감시 시작 → 타이머 작동"""
        self.login()
        if not self.account_number:
            print("[❌ 로그인 실패. 프로그램 종료]")
            return

        # 관심 종목 일봉 데이터 요청
        for code in self.target_stocks.keys():
            self.request_daily_chart(code)

        # 실시간 체결 감시 시작
        self.start_real_time_monitoring()

        # 타이머 시작
        self.check_timer.start(5000)          # 5초마다 장 종료 여부 확인
        self.balance_timer.start(60 * 60 * 1000)  # 1시간마다 잔액 조회

        print(f"[✅ 프로그램 준비 완료] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.app.exec_()
