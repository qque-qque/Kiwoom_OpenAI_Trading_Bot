# Kiwoom OpenAI Trading Bot

이 프로젝트는 Kiwoom OpenAPI를 사용하여 주식 자동 매매를 구현한 봇입니다. Python을 사용하여 주식 데이터를 분석하고, 예측된 매수 및 매도 시점에 따라 자동으로 매매를 수행합니다. 이를 통해 주식 거래 전략을 자동화하는 시스템을 개발할 수 있습니다.

## 책임 면제 및 모의투자 권장

본 프로그램은 Kiwoom OpenAPI를 사용하여 자동 매매를 구현한 도구입니다. 이 프로그램은 주식 거래에 대한 전략을 자동화하는 데 도움을 주며, 투자와 관련된 데이터 분석 및 예측을 기반으로 매수 및 매도 시점을 자동으로 결정하고 주문을 실행합니다. 그러나 본 프로그램은 교육적인 목적 또는 모의투자 환경에서만 사용되며, 실제 투자에 대한 책임을 지지 않습니다.

## 책임 면제

이 프로그램을 사용함으로써 발생하는 모든 투자 손실 및 이익에 대해서는 개발자나 본 프로그램의 제공자는 일체의 책임을 지지 않습니다. 사용자는 본 프로그램을 사용하여 이루어지는 모든 투자 활동의 결과에 대해 전적으로 자신의 책임하에 진행해야 합니다. 프로그램을 사용하는 것은 위험을 감수하는 행동이며, 실제 투자에 적용할 경우에는 신중하게 검토해야 합니다.

## 모의투자 권장

본 프로그램은 모의투자 환경에서 사용을 권장합니다. 실제 시장에서의 투자에 사용하기 전에 가상 자금을 사용하여 모의투자를 통해 프로그램의 성능을 테스트하고, 전략이 어떻게 작동하는지 파악하는 것이 중요합니다. 모의투자를 통해 시스템의 정확성과 위험성을 미리 확인한 후, 실제 자금으로 투자하는 것이 좋습니다.

## 주의 사항

본 프로그램은 실시간 주식 시장 데이터와 자동 매매 기능을 제공하지만, 예측의 정확성에 한계가 있습니다. 시장은 예측할 수 없는 변수를 포함하고 있으므로, 프로그램을 사용한 투자에 있어서도 시장 변화에 따른 리스크를 충분히 인식해야 합니다.

전문적인 투자 상담을 받지 않고 본 프로그램을 사용하여 투자 결정을 내리는 것은 큰 리스크를 동반할 수 있으며, 투자 손실을 초래할 수 있습니다. 따라서, 투자 전문가의 조언을 받는 것이 중요합니다.

## 주요 기능

- **주식 데이터 수집**: Kiwoom OpenAPI를 사용하여 실시간 주식 데이터를 수집합니다.
- **MACD 기반 매매 전략**: 이동 평균 수렴 확산(MACD) 지표를 사용하여 매수/매도 신호를 생성합니다.
- **자동 매매**: 예측된 매수/매도 시점에 자동으로 주문을 실행합니다.
- **주식 정보 관리**: 보유 주식과 거래 내역을 추적합니다.
- **수익률 분석**: 매매 기록을 기반으로 일별 수익률 그래프를 생성합니다.

## 설치 방법

### 1. 프로젝트 클론

```bash
git clone https://github.com/your-username/Kiwoom_OpenAI_Trading_Bot.git
cd Kiwoom_OpenAI_Trading_Bot
```

### 2. 의존성 설치

필요한 Python 라이브러리들을 설치합니다.

```bash
pip install -r requirements.txt
```

`requirements.txt` 파일은 아래와 같이 작성할 수 있습니다:

```txt
pandas
matplotlib
PyQt5
openpyxl
```

### 3. Kiwoom OpenAPI 설정

Kiwoom OpenAPI를 사용하려면 `config.ini` 파일에 계좌 정보 및 거래 관련 설정을 입력해야 합니다. `config.ini` 파일을 다음과 같은 형식으로 설정하십시오:

```ini
[USER]
account_pw = my_secure_password   # 사용자 계좌 비밀번호

[TRADING]
max_profit_rate = 0.0            # 최대 수익률 (매매 시 #% 수익이 나면 자동으로 매도)
max_loss_rate = -0.0              # 최대 손실률 (매매 시 -#% 손실이 나면 자동으로 손절)
max_trade_amount = 0        # 한 번의 매매에서 최대 #원까지만 거래
max_holding_count = 0             # 최대 보유 종목 개수
target_list = "{'종목 코드': '종목명'}"  # 거래할 종목 목록

```

### 4. Kiwoom OpenAPI 라이브러리 설치

Kiwoom OpenAPI를 설치하려면 Kiwoom에서 제공하는 `KHOpenAPI`를 설치하고 PyQt5 라이브러리가 설치된 상태여야 합니다. 해당 라이브러리를 설치하려면 아래 링크에서 제공된 방법을 따라 설치하세요.

- [Kiwoom OpenAPI 설치 가이드](https://www.kiwoom.com)

## 사용 방법

1. **`config.ini` 파일 설정**: 위의 설정 방법을 참고하여 `config.ini` 파일을 설정합니다.
2. **`main.py` 실행**: 설정이 완료되면 `main.py` 파일을 실행하여 주식 자동 매매 프로그램을 시작할 수 있습니다.

```bash
python src/main.py
```

3. **로그인**: 프로그램을 실행하면 Kiwoom OpenAPI에 로그인하여 계좌 정보를 가져옵니다.
4. **매매 실행**: 매수/매도 시점이 되면 자동으로 주문이 실행됩니다.
5. **결과 확인**: 거래 기록과 수익률을 `trade_log_YYYYMMDD.xlsx`와 `profit_graph_YYYYMMDD.png` 파일로 저장합니다.

## 프로젝트 구조

```
Kiwoom_OpenAI_Trading_Bot/
├── src/
│   ├── main.py               # 프로그램 실행 파일
│   └── kiwoom.py             # Kiwoom OpenAPI 연동 코드
├── utils/
│   └── kiwoom.py             # Kiwoom OpenAPI 연동 코드 (모듈화된 버전)
├── config.ini                # 설정 파일 (계좌 정보, 거래 설정)
├── requirements.txt          # 의존성 목록
├── LICENSE                   # 라이선스 파일
└── README.md                 # 프로젝트 설명
```

## 기여

이 프로젝트에 기여하려면, 먼저 포크한 후, 새 브랜치를 만들어 작업하십시오. 작업이 완료되면 풀 리퀘스트(PR)를 제출해 주세요.

## 라이선스

이 프로젝트는 [MIT 라이선스](LICENSE) 하에 배포됩니다.
