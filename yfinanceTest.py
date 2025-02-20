from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QEventLoop
from PyQt5.QAxContainer import QAxWidget
import time
import numpy as np

class KiwoomAPI(QAxWidget):  # QAxWidget 상속
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # Kiwoom OpenAPI 객체를 설정
        self.OnEventConnect.connect(self._login_slot)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)  # 오타 수정

    def dynamicCall(self, method_name, *args):
        """ 동적으로 호출하는 메서드 """
        return super().dynamicCall(method_name, *args)

    def comm_connect(self):
        """ 로그인 연결 메소드 """
        self.dynamicCall("CommConnect()")

class Kiwoom(KiwoomAPI):  # KiwoomAPI 상속
    def __init__(self):
        super().__init__()  # KiwoomAPI 초기화
        self._set_event_handlers()  # 이벤트 핸들러 설정

        if not QApplication.instance():
            self.app = QApplication([])

        self.tr_event_loop = QEventLoop()  # 이벤트 루프
        self.login_event_loop = QEventLoop()
        self.top_volume_stocks = []  # 거래량 상위 종목을 저장할 리스트
        self.stock_data = {}  # 종목 데이터 저장할 딕셔너리

    def _set_event_handlers(self):
        self.OnEventConnect.connect(self._login_slot)
        self.OnReceiveTrData.connect(self._on_receive_tr_data)  # 오타 수정

    def _login_slot(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print(f"로그인 실패, 에러코드: {err_code}")
        self.login_event_loop.quit()

    def get_top_volume_stocks(self):
        """ 거래량 상위 100개 종목 가져오기 """
        try:
            self.dynamicCall("SetInputValue(QString, QString)", "시장구분", "000")
            self.dynamicCall("SetInputValue(QString, QString)", "정렬기준", "1")
            self.dynamicCall("SetInputValue(QString, QString)", "조회개수", "100")
            self.dynamicCall("CommRqData(QString, QString, int, QString)", "거래량순위조회", "opt10030", 0, "0001")

            self.tr_event_loop.exec_()
            print(f"거래량 상위 100개 종목: {self.top_volume_stocks}")
        except Exception as e:
            print(f"Error in get_top_volume_stocks: {e}")

    def request_stock_data(self, code):
        """ 개별 종목 차트 데이터 요청 """
        try:
            print(f"[차트 요청] 종목코드: {code}")
            self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", "20240220")
            self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
            self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식일봉조회", "opt10081", 0, "0002")
            self.tr_event_loop.exec_()
        except Exception as e:
            print(f"Error in request_stock_data for {code}: {e}")

    def _on_receive_tr_data(self, scr_no, rq_name, tr_code, record_name, prev_next):
        if rq_name == "거래량순위조회":
            for i in range(100):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "종목코드").strip()
                if not code:
                    break
                self.top_volume_stocks.append(code)
            print(f"거래량 상위 100개 종목: {self.top_volume_stocks}")
            self.tr_event_loop.quit()

        elif rq_name == "주식일봉조회":
            code = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, 0, "종목코드").strip()
            print(f"[데이터 수신] 종목코드: {code}")

            prices = []
            for i in range(20):
                price = self.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, i, "현재가").strip()
                if price:
                    print(f" - {i + 1}일 전 가격: {price}")
                    prices.append(abs(int(price)))

            if len(prices) >= 20:
                ma5 = np.mean(prices[:5])
                ma20 = np.mean(prices[:20])
                std_dev = np.std(prices[:20])
                bb_upper = ma20 + (std_dev * 2)
                bb_lower = ma20 - (std_dev * 2)
            else:
                ma5, ma20, bb_upper, bb_lower = None, None, None, None

            print(f"[계산 완료] {code} - MA5: {ma5}, MA20: {ma20}, BB_Upper: {bb_upper}, BB_Lower: {bb_lower}")
            self.stock_data[code] = {"MA5": ma5, "MA20": ma20, "BB_Upper": bb_upper, "BB_Lower": bb_lower}
            self.tr_event_loop.quit()

    def send_order(self, order_type, stock_code, qty, price, hoga_type="03"):
        """ 주식 매수/매도 주문을 처리하는 메소드 """
        """
        order_type: 'buy' (매수) / 'sell' (매도)
        stock_code: 종목 코드
        qty: 거래 수량
        price: 거래 가격
        hoga_type: 주문 방식 (시장가: 03, 지정가: 00)
        """
        # 주문 타입 매핑
        order_type_dict = {"buy": 1, "sell": 2}

        # 주문 실행을 위한 파라미터 설정
        account_number = "80972740"  # 실제 계좌번호로 바꿔야 해
        screen_number = "0000"  # 화면 번호 설정

        # SendOrder 호출 (주문 전송)
        order_result = self.dynamicCall(
            "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
            ["매매", account_number, screen_number, order_type_dict[order_type], stock_code, qty, price, hoga_type, "00"]
        )

        # 주문 결과 출력
        if order_result == 0:
            print(f"{order_type.capitalize()} 주문 전송 성공: 종목코드: {stock_code}, 수량: {qty}, 가격: {price}")
        else:
            print(f"{order_type.capitalize()} 주문 전송 실패: 에러코드 {order_result}")

    def run(self):
        """ 전체 실행 """
        self.comm_connect()  # 로그인 요청
        self.login_event_loop.exec_()  # 로그인 완료 대기
        self.get_top_volume_stocks()

        for idx, code in enumerate(self.top_volume_stocks):
            print(f"[{idx + 1}/100] {code} 차트 데이터 요청 중...")
            self.request_stock_data(code)
            self.tr_event_loop.exec_()
            time.sleep(0.3)  # API 제한 대응

            # 매매 로직 추가 (MA5, MA20, 볼린저 밴드 기반 매수/매도)
            stock = self.stock_data.get(code)
            if stock:
                ma5 = stock["MA5"]
                ma20 = stock["MA20"]
                bb_upper = stock["BB_Upper"]
                bb_lower = stock["BB_Lower"]
                current_price = int(self.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10081", "주식일봉조회", 0, "현재가").strip())

                if ma5 and ma20:
                    if current_price < bb_lower:  # 매수 조건: 가격이 볼린저 밴드 하한선 아래일 때
                        print(f"[매수 추천] {code} - 현재가: {current_price}, 매수 조건: 가격이 BB 하한선 아래")
                        self.send_order("buy", code, 1, current_price)  # 예시로 1주 매수
                    elif current_price > bb_upper:  # 매도 조건: 가격이 볼린저 밴드 상한선 위일 때
                        print(f"[매도 추천] {code} - 현재가: {current_price}, 매도 조건: 가격이 BB 상한선 위")
                        self.send_order("sell", code, 1, current_price)  # 예시로 1주 매도

if __name__ == "__main__":
    app = QApplication([])  # QApplication 인스턴스 생성
    kiwoom = Kiwoom()
    kiwoom.run()
