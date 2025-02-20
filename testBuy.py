from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer, QObject
import sys


class KiwoomAPI(QObject):  # QObject를 상속
    def __init__(self):
        super().__init__()  # QObject 초기화
        self.app = QApplication(sys.argv)
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        if not self.kiwoom.control():
            print("Kiwoom OpenAPI 컨트롤 로드 실패")
            sys.exit(0)

        print("Kiwoom API 객체 생성 완료")

        # 로그인 이벤트 연결
        self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.OnEventConnect.connect(self.on_login)

        # 타이머 설정: 1초마다 현재가 확인
        self.timer = QTimer(self)  # self는 이제 QObject를 상속하므로 QTimer가 동작함
        self.timer.timeout.connect(self.check_price)
        self.timer.start(1000)  # 1000ms마다 호출 (1초 간격)

        self.app.exec_()  # PyQt 이벤트 루프 실행

    def on_login(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print(f"로그인 실패: {err_code}")

    def check_price(self):
        stock_code = "005930"  # 삼성전자 코드
        print(f"현재가 확인: {stock_code}")  # 로그 추가
        current_price = self.get_stock_price(stock_code)

        print(f"현재 {stock_code}의 시세: {current_price}")

        # 예시: 특정 가격 이상이면 매수
        if current_price < 55000:
            print(f"매수 조건 충족! {stock_code}를 매수합니다.")
            self.buy_stock(stock_code, current_price, 10)  # 70000원 이상일 때 10주 매수

    def get_stock_price(self, stock_code):
        price = self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", stock_code)
        return int(price)

    def buy_stock(self, stock_code, price, quantity):
        order_type = 1  # 1: 매수, 2: 매도
        order_number = "0000"  # 주문번호
        account_number = self.kiwoom.dynamicCall("GetAccountList()")  # 계좌번호 가져오기
        self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString) ",
         "매수주문",  # 주문명
         account_number,  # 계좌번호
         order_type,  # 매수
         stock_code,  # 주식 코드
         quantity,  # 수량
         price,  # 가격
         "55000",  # 가격조건
         "")  # 옵션 (비워두면 기본)
        print(f"{stock_code} 매수 주문을 요청했습니다.")


if __name__ == "__main__":
    KiwoomAPI()
