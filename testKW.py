import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget

class KiwoomAPI:
    def __init__(self):
        self.app = QApplication(sys.argv)

        # 1. Kiwoom API 객체 생성
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        print("Kiwoom API 객체 생성 완료")

        # 2. 이벤트 핸들러 연결
        self.kiwoom.OnEventConnect.connect(self.on_login)

        # 3. 로그인 요청
        self.kiwoom.dynamicCall("CommConnect()")
        print("로그인 요청 완료")

        # 4. PyQt 이벤트 루프 실행
        self.app.exec_()

    def on_login(self, err_code):
        """로그인 이벤트 콜백 함수"""
        if err_code == 0:
            print("로그인 성공!")
        else:
            print(f"로그인 실패 (에러 코드: {err_code})")
        self.app.quit()  # 로그인 후 프로그램 종료

if __name__ == "__main__":
    KiwoomAPI()
