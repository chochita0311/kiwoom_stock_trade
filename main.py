import sys
import cgitb

import GetLoginInfo as GetLoginInfo

from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import QTimer, QTime

from pywinauto import application
from pywinauto import timings

import time
import random


cgitb.enable(format='text')


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("my trading system")
        self.setGeometry(50, 50, 920, 700)

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")

        self.timer = QTimer(self)  # Status Timer
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self._app = application.Application()
        self.dlg0 = timings.WaitUntilPasses(20, 0.5, lambda: self._app.window_(title="Open API Login"))

        time.sleep(1)
        self.btn_ctrl = self.dlg0.Button1
        self.btn_ctrl.Click()

        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
        self.kiwoom.OnReceiveConditionVer.connect(self.receive_cond_list)
        self.kiwoom.OnReceiveRealCondition.connect(self.receive_real_cond)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_chejan)
        self.kiwoom.OnReceiveRealData.connect(self.receive_realdata)
        self.kiwoom.OnReceiveMsg.connect(self.on_receive_msg)

        self.req_index = 0
        self.order_index = 0

        self.base_bet = 1000000
        self.compound_interest = 1

        self.che_price_list = {}

        self._searched_list = set()
        self._minute_searched_list = []

        # [편입시각(f), 종목코드, 현재가, 주문수량, 매도가, 손절가, 매도접수번호, 매도잔량] + [최우선매수호가]
        self._element = []
        self._list = []

        bt_apiconn = QPushButton("API 접속", self)  # API Connect
        bt_apiconn.move(20, 20)  # ui_stat
        bt_apiconn.clicked.connect(self.bt_apiconn_clicked)

        general_stat_label = QLabel('일반: ', self)
        general_stat_label.move(20, 100)  # ui_stat
        self.general_stat = QTextEdit(self)  # General Status
        self.general_stat.setGeometry(20, 130, 280, 500)  # ui_stat
        self.general_stat.setReadOnly(True)

        trade_stat_label = QLabel('매매 내역: ', self)
        trade_stat_label.move(310, 100)  # ui_stat
        self.trade_stat = QTextEdit(self)  # Trade Status
        self.trade_stat.setGeometry(310, 130, 280, 500)  # ui_stat
        self.trade_stat.setReadOnly(True)

        order_stat_label = QLabel('주문 정보: ', self)
        order_stat_label.move(600, 100)  # ui_stat
        self.order_stat = QTextEdit(self)  # Order Status
        self.order_stat.setGeometry(600, 130, 280, 500)  # ui_stat
        self.order_stat.setReadOnly(True)

        bt2 = QPushButton("계좌 가져오기", self)  # 계좌 얻기
        bt2.move(310, 20)  # ui_stat
        bt2.clicked.connect(self.bt2_clicked)

        self.accounts_listbox = QComboBox(self)  # 계좌 목록
        self.accounts_listbox.move(310, 60)  # ui_stat
        self.accounts_listbox.setStyleSheet(
            "QComboBox{background-color: white;}")

        bt2_1 = QPushButton("비밀번호 등록", self)  # 비밀번호 등록
        bt2_1.move(430, 20)  # ui_stat
        bt2_1.clicked.connect(self.bt2_1_clicked)

    def timeout(self):
        current_time = QTime.currentTime().toString("hh:mm:ss")
        time_msg = "현재시각: " + current_time
        self.statusBar().showMessage(time_msg)

        if len(self._list) != 0:  # 자동 청산
            # element : [편입시각(f), 종목코드, 현재가, 주문수량, 매도가, 손절가, 매도접수번호, 매도잔량, 최우선매수호가]
            for element in self._list:
                if len(element) == 9:
                    if element[7] > 0:
                        if int(element[8][1:]) <= element[5]:  # 손절 룰
                            self.trade_stat.append(
                                "'" + element[1] + "' 손절 매도 주문을 넣겠습니다. - 매도 취소")  # 매도 취소
                            self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [
                                                    "sell_order_" + self.gen_uuid() + "_%d" % self.order_index, "0101", self.accounts_listbox.currentText(), 4, element[1], element[7], 0, "00", element[6]])
                            self.order_index += 1
                            self._list.remove(element)
                        elif current_time[:5] != time.ctime(element[0]).split()[3][:5]:  # 이탈 매도 룰
                            self.trade_stat.append(
                                "'" + element[1] + "' 시간 경과 매도 주문을 넣겠습니다. - 매도 취소")  # 매도 취소
                            self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [
                                                    "sell_order_" + self.gen_uuid() + "_%d" % self.order_index, "0101", self.accounts_listbox.currentText(), 4, element[1], element[7], 0, "00", element[6]])
                            self.order_index += 1
                            self._list.remove(element)

    def on_receive_msg(self, screen_num, rqname, trcode, msg):  # 기타 메시지 수신
        print("OnReceiveMsg sScrNo: " + screen_num + ", sRQName: " +
              rqname + ", sTrCode: " + trcode + ", sMsg: " + msg)

    def bt_apiconn_clicked(self):  # API Connect
        self.kiwoom.dynamicCall("CommConnect()")

    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):  # 종목 조회
        if "che_req10046" in rqname:
            current_time = QTime.currentTime().toString("hh:mm:ss")

            trcode2 = rqname[:6]  # 종목코드
            trname = self.kiwoom.dynamicCall(
                "GetMasterCodeName(QString)", [trcode2]).strip()  # 종목명
            trname = trname.replace(' ', '-')
            pr_price = self.kiwoom.dynamicCall(
                "CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "현재가")
            if pr_price.strip()[0] != "+" and pr_price.strip()[0] != "-":
                pr_price = "+" + pr_price.strip()
            pr_id_rate = self.kiwoom.dynamicCall(
                "CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "등락율")
            pr_id_rate = float(pr_id_rate.strip())

            case = int(pr_price.strip()[1:]) < 2000 and pr_id_rate > 10

            if case:
                target_rate = 1.01
                order_q = int(self.base_bet // float(pr_price.strip()[1:]))

                self.trade_stat.append("---------------------------------")
                self.trade_stat.append("편입시각: " + current_time)
                self.trade_stat.append("종목명: " + trname)
                self.trade_stat.append("현재가: " + pr_price.strip()[1:])
                self.trade_stat.append("주문수량: " + str(order_q))
                self.trade_stat.append("---------------------------------")

                self._element.extend([time.time(), trcode2, pr_price.strip(
                )[1:], order_q, target_rate])  # [편입시각(f), 종목코드, 현재가, 주문수량, 목표수익률]
                self._list.append(self._element)

                self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", ["buy_order_" + self.gen_uuid(
                ) + "_%d" % self.order_index, "0101", self.accounts_listbox.currentText(), 1, self._element[1], self._element[3], 0, "03", ""])
                self.order_index += 1

                self._element = []

    def receive_realdata(self, code, real_type, data):
        if len(self._list) != 0 and real_type == "주식체결":
            # element : [편입시각(f), 종목코드, 현재가, 주문수량, 매도가, 손절가, 매도접수번호, 매도잔량, 최우선매수호가]
            for element in self._list:
                if len(element) >= 8 and element[1] == code:
                    if len(element) == 8:
                        element.append(self.kiwoom.dynamicCall(
                            "GetCommRealData(QString, int)", code, 28))  # element[8]
                    else:
                        element[8] = self.kiwoom.dynamicCall(
                            "GetCommRealData(QString, int)", code, 28)  # element[8]
                        if int(self.r_per_hoga(element[8][1:], 0.98)) > element[5]:
                            element[5] = int(self.r_per_hoga(
                                element[8][1:], 0.98))

    def receive_chejan(self, sGubun, nItemCnt, sFidList):
        if sGubun == "0":
            order_type = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 905)  # 주문구분
            order_type = order_type[1:]
            order_stat = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 913)  # 주문상태
            tr_name = self.kiwoom.dynamicCall("GetChejanData(int)", 302)  # 종목명
            tr_code = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 9001)  # 종목코드
            tr_code = tr_code[1:]
            order_num = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 9203)  # 주문번호
            order_time = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 908)  # 주문/체결시각
            order_q = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 900)  # 주문수량
            che_price = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 910)  # 체결가
            che_q = self.kiwoom.dynamicCall("GetChejanData(int)", 911)  # 체결량
            remain_q = self.kiwoom.dynamicCall(
                "GetChejanData(int)", 902)  # 미체결수량

            self.order_stat.append("---------------------------------")
            self.order_stat.append("주문구분: " + order_type)
            self.order_stat.append("주문상태: " + order_stat)
            self.order_stat.append("종목명: " + tr_name)
            self.order_stat.append("종목코드: " + tr_code)
            self.order_stat.append("주문번호: " + order_num)
            self.order_stat.append("주문/체결시각: " + order_time)
            self.order_stat.append("주문수량: " + order_q)
            self.order_stat.append("체결가: " + che_price)
            self.order_stat.append("체결량: " + che_q)
            self.order_stat.append("미체결수량: " + remain_q)
            self.order_stat.append("---------------------------------")

            # 손절
            if order_stat == "확인" and order_type == "매도취소":
                if int(remain_q) == 0:
                    self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [
                                            "sell_order_" + self.gen_uuid() + "_%d" % self.order_index, "0101", self.accounts_listbox.currentText(), 2, tr_code, int(order_q), 0, "03", ""])
                    self.order_index += 1

            if len(self._list) != 0:
                if order_stat == "체결" and order_type == "매수":
                    # element : [편입시각(f), 종목코드, 현재가, 주문수량, 목표수익률 -> 매도가, 손절가]
                    for element in self._list:
                        if element[1] == tr_code and element[3] == int(che_q):
                            self.trade_stat.append(
                                ": '" + tr_name.strip() + "' 주문수량을 모두 매수하였습니다.")
                            target_rate = element[4]
                            element[4] = self.r_per_hoga(
                                element[2], target_rate)  # element[4]
                            # 2.00% 손절가 element[5]
                            element.append(
                                int(self.r_per_hoga(element[2], 0.98)))
                            self.trade_stat.append(str(
                                target_rate) + "%(" + element[4] + "원) 보유물량(" + str(element[3]) + ") 매도 주문을 넣겠습니다.")  # 지정가 매도
                            self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", ["sell_order_" + self.gen_uuid(
                            ) + "_%d" % self.order_index, "0101", self.accounts_listbox.currentText(), 2, element[1], element[3], int(element[4]), "00", ""])
                            self.order_index += 1

                if order_stat == "접수" and order_type == "매도":
                    # element : [편입시각(f), 종목코드, 현재가, 주문수량, 매도가, 손절가, 매도접수번호, 매도잔량]
                    for element in self._list:
                        if element[1] == tr_code and element[3] == int(order_q):
                            element.append(order_num)  # element[6]
                            element.append(int(order_q))  # element[7]

                if order_stat == "체결" and order_type == "매도":
                    # element : [편입시각(f), 종목코드, 현재가, 주문수량, 매도가, 손절가, 매도접수번호, 매도잔량] + [최우선매수호가]
                    for element in self._list:
                        if len(element) >= 8:
                            if element[1] == tr_code and element[6] == order_num:
                                element[7] = int(order_q) - \
                                    int(che_q)  # element[7]
                                if element[3] == int(che_q):
                                    self.trade_stat.append(
                                        ": '" + tr_name.strip() + "' 스탑로스 전에 주문수량을 모두 매도하였습니다.")
                                    self.trade_stat.append(
                                        "매도 감시 리스트에서 제거하였습니다.")
                                    self._list.remove(element)

    def bt2_clicked(self):  # 계좌 얻기
        GetLoginInfo.GetAccounts(self)

    def bt2_1_clicked(self):  # 비밀번호 등록
        self.kiwoom.dynamicCall(
            "KOA_Functions(QString, QString)", "ShowAccountWindow", "")

    @staticmethod
    def r_per_hoga(data, r):
        data = round(int(data) * r)
        if 1000 <= data < 5000:
            data = str(data)[:-1] + "5"
        elif 5000 <= data < 10000:
            data = str(data)[:-1] + "0"
        elif 10000 <= data < 50000:
            data = str(data)[:-2] + "50"
        elif 50000 <= data < 100000:
            data = str(data)[:-2] + "00"
        elif 100000 <= data < 500000:
            data = str(data)[:-3] + "500"
        elif 500000 <= data:
            data = str(data)[:-3] + "000"
        return str(data)

    @staticmethod
    def gen_uuid():
        random_string = ''
        random_str_seq = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for i in range(0, 7):
            random_string += str(
                random_str_seq[random.randint(0, len(random_str_seq) - 1)])
        return random_string

    def closeEvent(self, event):
        self.deleteLater()


# Run Main Window #

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
