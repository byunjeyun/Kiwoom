from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtWidgets import QApplication
import sys

class KiwoomAPI:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        if not self.kiwoom.control():
            print("Kiwoom OpenAPI 컨트롤 로드 실패")
            sys.exit(0)

        print("Kiwoom API 객체 생성 완료")

        # 로그인 이벤트 연결
        self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.OnEventConnect.connect(self.on_login)

        self.app.exec_()  # PyQt 이벤트 루프 실행

    def on_login(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print(f"로그인 실패: {err_code}")
        self.app.quit()

if __name__ == "__main__":
    KiwoomAPI()