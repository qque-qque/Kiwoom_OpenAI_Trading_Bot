import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import configparser
from datetime import datetime, time as dtime
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QEventLoop, QTimer
from PyQt5.QAxContainer import QAxWidget
from openpyxl import Workbook

class Kiwoom:
    def __init__(self):
        print("[🟢 프로그램 초기화 중...]")
        
        self.app = QApplication(sys.argv)
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._on_login)
        self.ocx.OnReceiveTrData.connect(self._on_receive_tr_data)
        self.ocx.OnReceiveRealData.connect(self._on_receive_real_data)

        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')  # 인코딩을 UTF-8로 설정

        self.account_pw = config['USER']['account_pw']
        self.max_profit_rate = float(config['TRADING']['max_profit_rate'])
        self.max_loss_rate = float(config['TRADING']['max_loss_rate'])
        self.max_trade_amount = int(config['TRADING']['max_trade_amount'])
        self.max_holding_count = int(config['TRADING']['max_holding_count'])
        self.target_stocks = eval(config['TRADING']['target_list'])

        self.login_event_loop = None
        self.tr_event_loop = None
        self.account_number = None
        self.histories = {}
        self.own_stocks = {}
        self.trade_log = []
        self.macd_data = {}
        self.banned_stocks = set()
        self.can_buy = True
        
        self.current_screen_no = 2001
        self.screen_by_code = {}
        self.closes_by_code = {}

        self.check_timer = QTimer()
        self.check_timer.timeout.connect(self.check_market_status)

        self.balance_timer = QTimer()
        self.balance_timer.timeout.connect(self.check_balance)

    def login(self):
        print("[🔐 로그인 요청 중...]")
        self.ocx.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _on_login(self, err_code):
        if err_code == 0:
            print("[✅ 로그인 성공]")
            self.account_number = self.ocx.dynamicCall("GetLoginInfo(QString)", "ACCNO").split(';')[0]
        else:
            print(f"[❌ 로그인 실패] 에러코드: {err_code}")
        self.login_event_loop.exit()

    def request_daily_chart(self, code):
        print(f"[📈 {self.target_stocks.get(code, code)}({code}) 일봉 데이터 요청 중...]")
        self.screen_by_code[code] = str(self.current_screen_no)
        self.current_screen_no += 1

        today = datetime.now().strftime("%Y%m%d")
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "기준일자", today)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10081_req", "opt10081", 0, self.screen_by_code[code])

        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def _on_receive_tr_data(self, screen_no, rqname, trcode, recordname, prev_next):
        if rqname == "opt10081_req":
            for code, screen in self.screen_by_code.items():
                if screen == screen_no:
                    self.handle_daily_chart(trcode, rqname, code)
                    break
        elif rqname == "opw00018_req":
            self.handle_balance(trcode, rqname)

    def handle_daily_chart(self, trcode, rqname, code):
        count = self.ocx.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        closes = []
        for i in range(count):
            close = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i, "현재가").strip()
            try:
                closes.append(abs(int(close)))
            except (ValueError, TypeError) as e:
                print(f"[❌ 데이터 변환 에러] {code} {i}번째 데이터 변환 실패: {e}")
                continue

        if not closes:
            print(f"[⚠️ 데이터 없음] {code} 종목 일봉 데이터가 없습니다.")
            self.tr_event_loop.exit()
            return

        closes.reverse()
        closes = pd.Series(closes)
        ema12 = closes.ewm(span=12, adjust=False).mean()
        ema26 = closes.ewm(span=26, adjust=False).mean()
        macd_line = ema12 - ema26
        signal_line = macd_line.ewm(span=9, adjust=False).mean()

        self.macd_data[code] = (macd_line, signal_line)
        self.closes_by_code[code] = closes

        self.predict_trading(code, macd_line, signal_line)

        self.tr_event_loop.exit()

    def predict_trading(self, code, macd_line, signal_line):
        macd_now, macd_prev = macd_line.iloc[-1], macd_line.iloc[-2]
        signal_now, signal_prev = signal_line.iloc[-1], signal_line.iloc[-2]

        if macd_prev < signal_prev and macd_now > signal_now:
            print(f"[📈 예측 매수] {self.target_stocks.get(code, code)}({code})")
            self.try_buy(code)

        if macd_prev > signal_prev and macd_now < signal_now:
            print(f"[📉 예측 매도] {self.target_stocks.get(code, code)}({code})")
            self.try_sell(code)

    def handle_balance(self, trcode, rqname):
        출금가능금액_raw = self.ocx.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, 0, "출금가능금액")

        try:
            출금가능금액 = abs(int(출금가능금액_raw.strip()))
        except (ValueError, AttributeError) as e:
            print(f"[❌ 출금가능금액 변환 에러] 원본 데이터: {출금가능금액_raw} / 에러: {e}")
            출금가능금액 = 0

        print(f"[💰 출금 가능 금액: {출금가능금액:,}원]")

        self.can_buy = 출금가능금액 >= self.max_trade_amount
        self.tr_event_loop.exit()

    def start_real_time_monitoring(self):
        fids = "10"
        codes = ";".join(self.target_stocks.keys())
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", "5000", codes, fids, "0")
        print("[📡 실시간 체결 감시 시작]")

    def _on_receive_real_data(self, code, real_type, real_data):
        if real_type != "주식체결":
            return

        try:
            price_raw = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 10).strip()
            price = abs(int(price_raw))
        except (ValueError, AttributeError) as e:
            print(f"[❌ 실시간 데이터 변환 에러] 종목: {code} / 원본 데이터: {price_raw if 'price_raw' in locals() else '없음'} / 에러: {e}")
            return

        if code not in self.own_stocks:
            self.own_stocks[code] = {"buy_price": price, "quantity": 1}

        stock_info = self.own_stocks.get(code)
        if stock_info:
            buy_price = stock_info['buy_price']
            profit_rate = (price - buy_price) / buy_price * 100
            print(f"[📈 수익률 확인] {self.target_stocks.get(code, code)}({code}) 매수가: {buy_price}원 / 현재가: {price}원 / 수익률: {profit_rate:.2f}%")

        if code in self.own_stocks:
            self.try_sell(code, price)

    def try_buy(self, code):
        if code in self.banned_stocks or not self.can_buy:
            return
        price = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 10).strip()
        try:
            price = abs(int(price))
        except (ValueError, AttributeError) as e:
            print(f"[❌ 실시간 가격 가져오기 실패] {code} - {e}")
            return

        print(f"[🛒 매수 시도] {self.target_stocks.get(code, code)}({code}) {price}원")
        self.send_order(code, 1, 1)

    def try_sell(self, code, price):
        stock = self.own_stocks.get(code)
        if not stock:
            return
        buy_price = stock['buy_price']
        quantity = stock['quantity']
        profit_rate = (price - buy_price) / buy_price * 100

        if profit_rate >= self.max_profit_rate or profit_rate <= self.max_loss_rate:
            print(f"[🛒 매도 시도] {self.target_stocks.get(code, code)}({code}) 수익률: {profit_rate:.2f}%")
            self.send_order(code, 2, quantity)
            self.record_trade(code, "매도", quantity, price, profit_rate)
            self.own_stocks.pop(code, None)

    def send_order(self, code, order_type, quantity):
        res = self.ocx.dynamicCall(
            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
            ["주문", "0101", self.account_number, order_type, code, quantity, 0, "03", ""]
        )
        if res == 0:
            print(f"[✅ 주문 성공] {self.target_stocks.get(code, code)}({code})")
        else:
            print(f"[❌ 주문 실패] {self.target_stocks.get(code, code)}({code})")
            self.banned_stocks.add(code)

    def record_trade(self, code, trade_type, quantity, price, profit_rate=None):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        profit_str = f"{profit_rate:.2f}" if profit_rate is not None else "-"
        self.trade_log.append([now, code, trade_type, quantity, price, profit_str])

    def save_trade_log(self):
         # 로그 디렉토리 경로
        log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
    
        # 디렉토리가 없으면 생성
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        wb = Workbook()
        ws = wb.active
        ws.append(["날짜", "종목코드", "매매구분", "수량", "가격", "수익률(%)"])
        for log in self.trade_log:
            ws.append(log)
        
        file_path = os.path.join(log_dir, f"trade_log_{datetime.now().strftime('%Y%m%d')}.xlsx")
        wb.save(file_path)
        print(f"[📜 매매 기록 저장 완료] {file_path}")

    def draw_profit_graph(self):
        profits = [(log[0], float(log[5])) for log in self.trade_log if log[2] in ["매도"] and log[5] != "-"]
        if not profits:
            return
        df = pd.DataFrame(profits, columns=["date", "profit"])
        df_grouped = df.groupby("date").sum().reset_index()

        plt.figure(figsize=(10, 6))
        plt.plot(df_grouped['date'], df_grouped['profit'], marker='o')
        plt.title('일별 수익률')
        plt.xlabel('날짜')
        plt.ylabel('수익률(%)')
        plt.grid()
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(f"profit_graph_{datetime.now().strftime('%Y%m%d')}.png")

    def check_balance(self):
        print("[💰 출금 가능 금액 조회 요청 중...]")
        self.can_buy = False
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_number)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.account_pw)
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opw00018_req", "opw00018", 0, "2000")

        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def check_market_status(self):
        now = datetime.now().time()
        if now < dtime(8, 0) or now > dtime(18, 0):
            print("[🚪 장 종료 감지] 프로그램 종료")
            self.shutdown()

    def run(self):
        self.login()
        if not self.account_number:
            print("[❌ 로그인 실패. 프로그램 종료]")
            return

        for code in self.target_stocks.keys():
            self.request_daily_chart(code)

        self.start_real_time_monitoring()
        self.check_timer.start(5000)
        self.balance_timer.start(60 * 60 * 1000)
        print("[✅ 프로그램 준비 완료]")
        self.app.exec_()

    def shutdown(self):
        print("[🛑 프로그램 종료] 매매 기록 저장 후 종료합니다...")
        self.save_trade_log()
        self.draw_profit_graph()
        self.app.quit()