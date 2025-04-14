# 🚀 Kiwoom OpenAI Trading Bot

**Kiwoom OpenAI Trading Bot**은 Kiwoom OpenAPI를 사용하여 Python으로 개발된 **주식 자동매매 프로그램**입니다.  
MACD 기반의 기술적 분석과 자동 주문 기능을 통해 매수/매도 시점을 예측하고, 전략을 실행합니다.

> **⚠️ 사용 전 주의사항**  
> 본 프로그램은 **학습용** 및 **모의투자용**으로 개발되었습니다.  
> 실제 투자에 사용 시 발생하는 모든 손실은 사용자 본인에게 책임이 있습니다.

---

## 🛡️ 책임 면제 (Disclaimer)

- 이 프로그램은 투자 수익을 보장하지 않으며, 투자에 따른 결과는 전적으로 사용자 책임입니다.
- 실제 투자에 사용하기 전에 **모의투자 환경**에서 충분히 테스트할 것을 강력히 권장합니다.
- **전문 투자 자문 없이 실거래 사용을 금지합니다.**

---

## ✨ 주요 기능 (Features)

- **실시간 주식 데이터 수집** : Kiwoom OpenAPI를 통해 데이터 수집
- **MACD 기반 자동매매** : 기술적 분석으로 매수/매도 신호 생성
- **자동 주문 실행** : 조건 충족 시 자동으로 주문 전송
- **수익률 기록 및 분석** : 매매 내역 저장, 수익률 그래프 생성
- **보유 주식 관리** : 매수/매도 현황 실시간 관리

---

## 📦 설치 방법 (Installation)

### 1. 프로젝트 클론

```bash
git clone https://github.com/your-username/Kiwoom_OpenAI_Trading_Bot.git
cd Kiwoom_OpenAI_Trading_Bot
```

### 2. Python 의존성 설치

```bash
pip install -r requirements.txt
```

**`requirements.txt` 예시**

```plaintext
PyQt5>=5.15.9
openpyxl>=3.1.2
pandas>=2.2.1
numpy>=1.26.4
matplotlib>=3.8.3
```

### 3. Kiwoom OpenAPI 설치

- [키움증권 OpenAPI+ 다운로드](https://www.kiwoom.com)
- 반드시 키움증권 계좌를 개설하고, 모의투자 계좌를 발급받아야 합니다.

### 4. 설정 파일(config.ini) 작성

`config.ini` 예시:

```ini
[USER]
account_pw = your_account_password   # 계좌 비밀번호 입력

[TRADING]
max_profit_rate = 0.0                # 최대 수익률 (%) 예: 5% 도달 시 익절
max_loss_rate = -0.0                 # 최대 손실률 (%) 예: -3% 도달 시 손절
max_stock_ratio = 0.0               # 종목당 투자비율 (%) 예: 총 잔액의 10%
max_holding_count = 0                # 최대 보유 종목 수
buy_split_count = 0                  # 분할 매수 횟수
restart_after_close = False          # 장 종료 후 자동 재시작 여부
target_list = {'종목코드': '종목명'}  # 매매할 종목
```

---

## 🖥️ 사용 방법 (Usage)

1. `config.ini` 설정 완료
2. 프로그램 실행

```bash
python src/main.py
```

3. **Kiwoom 로그인 창**이 나타나면 로그인
4. 프로그램이 자동으로 종목을 감시하고, 조건 만족 시 매수/매도 수행
5. 거래 기록 (`logs/trade_log_YYYYMMDD.xlsx`) 및 수익률 그래프 (`logs/profit_graph_YYYYMMDD.png`) 자동 저장

---

## 🗂️ 프로젝트 구조 (Project Structure)

```
Kiwoom_OpenAI_Trading_Bot/
├── src/
│   └── main.py                # 프로그램 실행 파일
├── utils/
│   └── kiwoom.py              # Kiwoom API 연동 모듈
├── config.ini                  # 설정 파일
├── requirements.txt            # 의존성 목록
├── LICENSE                     # 라이선스 파일
└── README.md                   # 프로젝트 설명
```

---

## 🤝 기여 방법 (Contributing)

1. 이 저장소를 Fork 합니다.
2. 새 브랜치를 생성합니다 (`git checkout -b feature/새로운기능`).
3. 코드를 커밋합니다 (`git commit -m "Add 새로운 기능"`).
4. 브랜치에 푸시합니다 (`git push origin feature/새로운기능`).
5. Pull Request를 생성합니다.

---

## 📜 라이선스 (License)

이 프로젝트는 [MIT License](LICENSE)를 따릅니다.

자유롭게 수정 및 배포 가능하나, 사용에 따른 책임은 사용자 본인에게 있습니다.
