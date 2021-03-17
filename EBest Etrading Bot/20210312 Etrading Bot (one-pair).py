#-*-coding: utf-8 -*-
import win32com.client
import pythoncom
import time
import pandas as pd
import sqlite3
from datetime import datetime
import numpy as np


# 로그인 역할을 수행하는 클래스 생성
class XSession:
    def __init__(self):
        self.login_state = 0

    def OnLogin(self, code, msg):  # event handler
        if code == "0000":
            print("※ {0} 로그인 완료했습니다.\n".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.login_state = 1
        else:
            self.login_state = 2
            print("※ 로그인 fail.. \n code={0}, message={1}\n".format(code, msg))

    def api_login(self, id="", pwd="", cert_pwd="", url=""):
        self.ConnectServer(url, 20001)
        is_connected = self.Login(id, pwd, cert_pwd, 0, False)

        if not is_connected:  # 서버에 연결 안되거나, 전송 에러시
            print("※ 로그인 서버 접속 실패... ")
            return

        while self.login_state == 0:   # 로그인이 될때까지 대기
            pythoncom.PumpWaitingMessages()

    def account_info(self):
        """
        계좌 정보 조회
        """
        if self.login_state != 1:  # 로그인 성공 아니면, 종료
            return

        account_no = self.GetAccountListCount()
        print("계좌 갯수 = {0}".format(account_no))

        for i in range(account_no):
            account = self.GetAccountList(i)
            print("계좌번호 = {0}".format(account))

    @classmethod
    def get_instance(cls):
        # DispatchWithEvents로 instance 생성하기
        XSession = win32com.client.DispatchWithEvents("XA_Session.XASession", cls)  # 서버에 클래스를 요청해서 받아오는 부분 cls 고정 ?????. 결과값은 클래스고, 서버에 3개중에 하나를 요청하는데 쓰이는 메써드
        return XSession            ## XSession 이라는 인스턴스로 ..


class XQuery_t2105:
    def __init__(self):
        super().__init__()
        self.is_data_received = False

    def OnReceiveData(self, tr_code):
        self.is_data_received = True
        # 사용 항목
        price = self.GetFieldData("t2105OutBlock", "price", 0)
        offerho1 = self.GetFieldData("t2105OutBlock", "offerho1", 0)
        bidho1 = self.GetFieldData("t2105OutBlock", "bidho1", 0)
        hotime = self.GetFieldData("t2105OutBlock", "time", 0)

        # 비사용 항목
        name = self.GetFieldData("t2105OutBlock", "hname", 0)
        volume = self.GetFieldData("t2105OutBlock","volume",0)
        offerrem1 = self.GetFieldData("t2105OutBlock", "offerrem1", 0)
        bidrem1 = self.GetFieldData("t2105OutBlock", "bidrem1", 0)
        dcnt1 = self.GetFieldData("t2105OutBlock", "dcnt1", 0)
        scnt1 = self.GetFieldData("t2105OutBlock", "scnt1", 0)

        print("테스트 입니다.", price, offerho1, bidho1)
        price = float(price)
        offerho1 = float(offerho1)
        bidho1 = float(bidho1)

        print("TR code는 't2105'며, 옵션 코드는 {}입니다.".format(self.optionStock))

        """
        이쪽 수정해야 한다.
        """
        # stockOpts_realtimeLog.append([self.optionStock, hotime, price, offerho1, bidho1])

        stockOpts_statusInfo[self.optionStock]['curBidho'] = bidho1
        stockOpts_statusInfo[self.optionStock]['preBidho'] = bidho1
        stockOpts_statusInfo[self.optionStock]['curOfferho'] = offerho1
        stockOpts_statusInfo[self.optionStock]['preOfferho'] = offerho1
        stockOpts_statusInfo[self.optionStock]['price'] = price
        print(hotime, type(hotime))
        stockOpts_statusInfo[self.optionStock]['midHo'] = (bidho1 + offerho1) / 2

        print("bidho는",bidho1,"offerho는",offerho1)
        print(self.optionStock, hotime, price, offerho1, bidho1)

    def request(self,optionStock):
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t2105.res"  # RES 파일 등록
        self.SetFieldData("t2105InBlock", "shcode", 0, optionStock)
        self.optionStock = optionStock

        err_code = self.Request(False)  # data 요청하기 --  연속조회인경우만 True

        if err_code < 0:
            print("error... {0}".format(err_code)) # data 요청하기 --  연속조회인경우만 True

    @classmethod
    def get_instance(cls):
        xq_t2105 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", cls)
        return xq_t2105


class XReal_OC0_:
    def __init__(self):
        super().__init__()
        self.count = 0

    def set_data(self, optionStock):
        self.optionStock = optionStock

    def OnReceiveRealData(self, tr_code):
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """

        self.count += 1

        # 사용 항목
        optcode = self.GetFieldData("OutBlock","optcode") #옵션코드, String type
        chetime = self.GetFieldData("OutBlock", "chetime") # 체결시간, String type
        price = self.GetFieldData("OutBlock", "price") # 현가, Float type
        offerho1 = self.GetFieldData("OutBlock", "offerho1") # 매도 1호가(판매), Float type
        bidho1 = self.GetFieldData("OutBlock", "bidho1") # 매수 1호가(구입), Float type

        # 비사용 항목
        # mdvolume = self.GetFieldData("OutBlock", "mdvolume") # 매도누적체결량
        # mdchecnt = self.GetFieldData("OutBlock", "mdchecnt") # 매도누적체결건수
        # msvolume = self.GetFieldData("OutBlock", "msvolume")  # 매수누적체결량
        # mschecnt = self.GetFieldData("OutBlock", "mschecnt") #매수누적체결건수
        # open = self.GetFieldData("OutBlock", "open") #시가
        # high = self.GetFieldData("OutBlock", "high") #고가
        # low = self.GetFieldData("OutBlock", "low") #저가

        price = float(price)
        offerho1 = float(offerho1)
        bidho1 = float(bidho1)

        print("\n\n")
        print("=======================================================================")
        print("=========================="+ str(self.count)+"번째 수신데이터============================")
        print("=======================================================================")
        print("아래 데이터를 수신했습니다.")
        print(tr_code, optcode,chetime,price,offerho1,bidho1,"\n")

        """
        #1 기존, 매도1호가&매수1호가를 이전 매도1호가&매수1호가로 저장한다.
        #2 신규로 받아온, 매도1호가&매수1호가를 현재 매도1호가&매수1호가로 저장한다.
        #3 매도1호가&매수1호가를 각각 list에 추가해 기록한다.
        """
        """
        저장하는 과정에서 1초 당 데이터가 저장되는 게 아니고 변환하는 시점의 데이터가 변환되기 떄문에.. 평균을 내기가 쉽지 않다.
        """

        #1
        stockOpts_statusInfo[optcode]["preOfferho"] = stockOpts_statusInfo[optcode]["curOfferho"]
        stockOpts_statusInfo[optcode]["preBidho"] = stockOpts_statusInfo[optcode]["curBidho"]

        #2
        stockOpts_statusInfo[optcode]["curOfferho"] = offerho1
        stockOpts_statusInfo[optcode]["curBidho"] = bidho1
        stockOpts_statusInfo[optcode]["price"] = price

        stockOpts_statusInfo[optcode]['midHo'] = (bidho1 + offerho1) / 2

        # 확인용 print문
        print("\n","#1 input data phase")
        print("현재매수가:", stockOpts_statusInfo[optcode]["curOfferho"], "이전매수가", stockOpts_statusInfo[optcode]["preOfferho"])
        print("현재매도가:", stockOpts_statusInfo[optcode]["curBidho"], "이전매도가", stockOpts_statusInfo[optcode]["preBidho"])
        print("중간 값:", stockOpts_statusInfo[optcode]["midHo"], "가격", stockOpts_statusInfo[optcode]["price"])
        print()


        #3
        for stockOpt in list(stockOpts.values()):
            stockOpts_statusLog[stockOpt]['Bidho'].append(stockOpts_statusInfo[stockOpt]['curBidho'])
            stockOpts_statusLog[stockOpt]["Offerho"].append(stockOpts_statusInfo[stockOpt]['curOfferho'])
            stockOpts_statusLog[stockOpt]['midHo'].append(stockOpts_statusInfo[stockOpt]['midHo'])
            # 조율이 필요

        #4
        deviationValue = (stockOpts_statusInfo[stockOpts['lowStock']]['midHo'] * 2 ) - stockOpts_statusInfo[stockOpts['highStock']]['midHo']

        print(stockOpts_indicatorInfo['deviationLogic'], "deviation logic dictionary 입니다.")

        if stockOpts_indicatorInfo['deviationLogic']['curValue'] == 0.0:
            stockOpts_indicatorInfo['deviationLogic']['curValue'] = deviationValue
            stockOpts_indicatorLog['deviationLogic']['log'].append(deviationValue)
            print("deviation은 ", deviationValue, "입니다.")
            print("현재 값이 0.0 입니다.")
        else:
            stockOpts_indicatorInfo['deviationLogic']['preValue'] = stockOpts_indicatorInfo['deviationLogic']['curValue']
            stockOpts_indicatorInfo['deviationLogic']['curValue'] = deviationValue
            stockOpts_indicatorLog['deviationLogic']['log'].append(deviationValue)
            print("deviation은 ", deviationValue,"입니다.")
            print("현재 값이 0.0이 아닙니다.")


        # stockOpts_indicatorLog['deviationLogic']
        global isOver150
        # global option_A_offer_state
        # global option_B_bid_state
        # global option_A_bid_state
        # global option_B_offer_state

        #5
        tempList = []
        tempList.extend([-9999,chetime,optcode])
        for stockOpt_ in list(stockOpts.values()):
            tempList.append(stockOpts_statusInfo[stockOpt_]["price"])
            tempList.append(stockOpts_statusInfo[stockOpt_]["curOfferho"])
            tempList.append(stockOpts_statusInfo[stockOpt_]["curBidho"])
            tempList.append(stockOpts_statusInfo[stockOpt_]["midHo"])

        tempList.append(deviationValue)
        stockOpts_realtimeLog.append(tempList)

        """
        #분기 150개 이상의 데이터가 쌓이지 않았다면, 아래를 실행한다.
        #1 매도1호가&매수1호가의 기록이 150개 이상이 되면, 'isOver150'을 True로 변환한다.
        #2 임시 List를 만든 뒤, '-9999, -9999, 체결시간, 옵션코드, ( 가격, 매도가, 매수가 ) * 옵션 종목 당' 을 추가한다.
        #3 Deviation Logic에 따라 값을 구한 뒤, log에 추가하고 현재값에 추가한다.
        """
        if isOver150 == False:
            """
            이쪽에서는 조율이 필요하다..
            """

            #1
            if len(stockOpts_statusLog[stockOpt]['Bidho']) > 149:
                isOver150 = True

            tempList.append(-9999)
            stockOpts_realtimeLog.append(tempList)


            print()
            print("★★ Realdata receive - 개수 파악 ★★")
            print("HighStock")
            print("ongBidState",transaction_statusInfo[stockOpts['highStock']]['ongBidState'])
            print("ongOfferState", transaction_statusInfo[stockOpts['highStock']]['ongOfferState'])
            print("LowStock")
            print("ongBidState", transaction_statusInfo[stockOpts['lowStock']]['ongBidState'])
            print("ongOfferState", transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'])
            print()

            # 테스트 코드
            if len(transaction_statusInfo[stockOpts['highStock']]['ongBidState']) + len(transaction_statusInfo[stockOpts['lowStock']]['ongOfferState']) < 1 :
                print(stockOpts['highStock'],"1 개 매도(판매) 주문완료")
                CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['highStock'],매매구분="1",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curOfferho'],주문수량='1') #
                print(stockOpts['lowStock'],"2 개 매수(구매) 주문완료")
                CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #


            current_time_int = int(datetime.now().strftime("%H%M")) # 추후에 900까지 고려할 것.
            if len(transaction_statusInfo[stockOpts['highStock']]['ongBidState']) > 0:
                print("높은 가격 주식이 매도중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongBidState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongBidState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")
                    for ordNum in transaction_statusInfo[stockOpts['highStock']]['ongOfferState']:
                        CFOAT00200(계좌번호=userInfo['account_num'], \
                        비밀번호=userInfo['cert_password'], \
                        선물옵션종목번호=stockOpts['highStock'], \
                        원주문번호=ordNum, \
                        선물옵션호가유형코드=transaction_detailedInfo[ordNum]['hoType'], \
                        주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curOfferho'], \
                        정정수량=transaction_detailedInfo[ordNum]['ordVolume'])
                    # 주문 시간이나 값의 변환은 CF0AT00200에 추가할 것

            if len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']) > 0:
                print("높은 가격 주식이 매수중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongOfferState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongOfferState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")

            if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                print("낮은 가격 주식이 매도중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongBidState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongBidState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")

            if len(transaction_statusInfo[stockOpts['lowStock']]['ongOfferState']) > 0:
                print("낮은 가격 주식이 매수중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")
                    for ordNum in transaction_statusInfo[stockOpts['lowStock']]['ongOfferState']:
                        CFOAT00200(계좌번호=userInfo['account_num'], \
                        비밀번호=userInfo['cert_password'], \
                        선물옵션종목번호=stockOpts['lowStock'], \
                        원주문번호=ordNum, \
                        선물옵션호가유형코드=transaction_detailedInfo[ordNum]['hoType'], \
                        주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curOfferho'], \
                        정정수량=transaction_detailedInfo[ordNum]['ordVolume'])


        #분기 150개 이상의 데이터가 쌓였다면, 아래를 실행한다.
        #1
        elif isOver150 == True:
            print("150 이상 돌입")

            stockOpts_indicatorLog['deviationLogic']['log'].pop(0)

            #1
            for stockOpt in list(stockOpts.values()):
                stockOpts_statusLog[stockOpt]['Bidho'].pop(0)
                stockOpts_statusLog[stockOpt]['Offerho'].pop(0)
                stockOpts_statusLog[stockOpt]['midHo'].pop(0)


                #확인용 print문1
                print(len(stockOpts_statusLog[stockOpt]['Offerho']), "150이여야 한다.")
                """
                만약에 평균 가격을 log에 추가한다면, 이쪽에서 추가해야한다.
                """



            stockOpts_indicatorInfo['deviationLogic']['avgValue'] = sum(stockOpts_indicatorLog['deviationLogic']['log']) / len(stockOpts_indicatorLog['deviationLogic']['log'])

            tempList.append(stockOpts_indicatorInfo['deviationLogic']['avgValue'])
            stockOpts_realtimeLog.append(tempList)


            print()
            print("★★ Realdata receive - 개수 파악 ★★")
            print("HighStock")
            print("ongBidState",transaction_statusInfo[stockOpts['highStock']]['ongBidState'])
            print("ongOfferState", transaction_statusInfo[stockOpts['highStock']]['ongOfferState'])
            print("LowStock")
            print("ongBidState", transaction_statusInfo[stockOpts['lowStock']]['ongBidState'])
            print("ongOfferState", transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'])
            print()


            current_time_int = int(datetime.now().strftime("%H%M")) # 추후에 900까지 고려할 것.




            if stockOpts_indicatorLog['deviationLogic']['log'][-1] >= 0.03 and stockOpts_indicatorLog['deviationLogic']['log'][-2] < 0.03:
                print("Deviation이 주문 시점이 됐습니다.")

                if len(transaction_statusInfo[stockOpts['highStock']]['ongBidState']) + len(transaction_statusInfo[stockOpts['lowStock']]['ongOfferState']) < 1 :
                    print(stockOpts['highStock'],"1 개 매도(판매) 주문완료")
                    CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['highStock'],매매구분="1",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curOfferho'],주문수량='1') #

                    print(stockOpts['lowStock'],"2 개 매수(구매) 주문완료")
                    CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #
                    # #@1 매도(판매) @2 매수(구입)

                    global upperCaseCount
                    upperCaseCount += 1

            elif stockOpts_indicatorLog['deviationLogic']['log'][-1] <= -0.03 and stockOpts_indicatorLog['deviationLogic']['log'][-2] > -0.03:
                print("Deviation이 주문 시점이 됐습니다.")
                if len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']) + len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) < 1:
                    print(stockOpts['lowStock'],"2 개 매도(판매) 주문완료")
                    CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['lowStock'],매매구분="1",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curOfferho'],주문수량='2') #

                    print(stockOpts['highStock'],"1 개 매수(구매) 주문완료")
                    CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'],선물옵션종목번호=stockOpts['highStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curBidho'],주문수량='1') #

                    global lowerCaseCount
                    lowerCaseCount += 1

            else:
                print("주문없이 종료시켰습니다.")
            print("\n")





            if len(transaction_statusInfo[stockOpts['highStock']]['ongBidState']) > 0:
                print("높은 가격 주식이 매도중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongBidState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongBidState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")

                    # CFOAT00200(계좌번호='',비밀번호='',선물옵션종목번호='0',원주문번호='',선물옵션호가유형코드='',주문가격='',정정수량='')
                    if len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']) > 0:
                        for ordNum in transaction_statusInfo[stockOpts['highStock']]['ongOfferState']:
                            CFOAT00200(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], 선물옵션종목번호=stockOpts['highStock'], \
                            원주문번호=ordNum,선물옵션호가유형코드=transaction_detailedInfo[ordNum]['hoType'],주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curOfferho'],정정수량=transaction_detailedInfo[ordNum]['hoType'])
                            #주문 수량이 다를 가능성이 있다.

                    # 재주문
                    #
                    #         time.sleep(5)
                    #         CFOAT00200(계좌번호='',비밀번호='',선물옵션종목번호='0',원주문번호='',선물옵션호가유형코드='',주문가격='',정정수량='')
                    #         if len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']) > 0:
                    #             for ordNum in transaction_statusInfo[stockOpts['highStock']]['ongOfferState']:
                    #                 CFOAT00200(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], 선물옵션종목번호=stockOpts['highStock'], \
                    #                 원주문번호=ordNum,선물옵션호가유형코드=transaction_detailedInfo[ordNum]['hoType'],주문가격=stockOpts_statusInfo[stockOpts['highStock']]['curOfferho'],정정수량=transaction_detailedInfo[ordNum]['hoType'])
                    #                 #주문 수량이 다를 가능성이 있다.
                    #         if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                    #             CFOAT00100(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], \
                    #             선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #



                    # 주문 시간이나 값의 변환은 CF0AT00200에 추가할 것

            if len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']) > 0:
                print("높은 가격 주식이 매수중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongOfferState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['highStock']]['ongOfferState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")
                    if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                        CFOAT00200(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], \
                        선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #


            if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                print("높은 가격 주식이 매도중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongBidState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongBidState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")
                    if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                        CFOAT00200(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], \
                        선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #


            if len(transaction_statusInfo[stockOpts['lowStock']]['ongOfferState']) > 0:
                print("높은 가격 주식이 매수중")
                print(int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'][0]]["OrdTime"])+0)
                print(current_time_int)
                if int(transaction_detailedInfo[transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'][0]]["OrdTime"]) + 0 < current_time_int:
                    print("재주문 합니다.")
                    if len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']) > 0:
                        CFOAT00200(계좌번호=userInfo['account_num'],비밀번호=userInfo['cert_password'], \
                        선물옵션종목번호=stockOpts['lowStock'],매매구분="2",선물옵션호가유형코드="00",주문가격=stockOpts_statusInfo[stockOpts['lowStock']]['curBidho'],주문수량='2') #






        print("길이 확인",len(stockOpts_statusLog[stockOpt]['Bidho']),len(stockOpts_statusLog[stockOpt]['Offerho']),len(stockOpts_statusLog[stockOpt]['midHo']),len(stockOpts_indicatorLog['deviationLogic']['log']))
        print()

        print("=======================================================================")
        print("=======================================================================")

    def start(self):
        """
        이베스트 서버에 실시간 data 요청함.
        """
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\OC0.res"  # RES 파일 등록
        self.SetFieldData("InBlock", "optcode",  self.optionStock)
        self.AdviseRealData()   # 실시간데이터 요청

    def add_item(self, stockcode):
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "optcode", stockcode)
        self.AdviseRealData()

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

    def end(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소
#
    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal

class XAQueryEvents:
    상태 = False

    def OnReceiveData(self, szTrCode):
        print("OnReceiveData : %s" % szTrCode)
        XAQueryEvents.상태 = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("OnReceiveMessage : ", systemError, messageCode, message)

class XReal_C01:
    def __init__(self):
        super().__init__()
        self.count = 0

    def set_data(self,option_code):
        self.option_code = option_code

    def OnReceiveRealData(self, tr_code):
        # 사용 항목
        ordno = self.GetFieldData("OutBlock","ordno") #  주문번호
        trcode = self.GetFieldData("OutBlock","trcode")
        orgordno = self.GetFieldData("OutBlock","orgordno") # 원주문번호
        chetime = self.GetFieldData("OutBlock", "chetime") # 체결시간
        chedate = self.GetFieldData("OutBlock", "chedate") # 체결일자
        chevol = self.GetFieldData("OutBlock", "chevol") # 체결량
        cheprice = self.GetFieldData("OutBlock", "cheprice") # 체결가격
        expcode = self.GetFieldData("OutBlock", "expcode") # 옵션종목
        dosugb = self.GetFieldData("OutBlock", "dosugb") #매도수 구분
        lineseq = self.GetFieldData("OutBlock", "lineseq") #라인 일련번호
        seq = self.GetFieldData("OutBlock", "seq") #일련번호
        megrpno = self.GetFieldData("OutBlock", "megrpno") #매칭그룹번호
        boardid = self.GetFieldData("OutBlock", "boardid") #보드ID
        sessionid = self.GetFieldData("OutBlock", "sessionid") #세션ID
        yakseq = self.GetFieldData("OutBlock", "yakseq") #약정번호

        # 비사용 항목
        # mdvolume = self.GetFieldData("OutBlock", "mdvolume") # 매도누적체결량
        # mdchecnt = self.GetFieldData("OutBlock", "mdchecnt") # 매도누적체결건수
        # msvolume = self.GetFieldData("OutBlock", "msvolume")  # 매수누적체결량
        # mschecnt = self.GetFieldData("OutBlock", "mschecnt") #매수누적체결건수
        # open = self.GetFieldData("OutBlock", "open") #시가
        # high = self.GetFieldData("OutBlock", "high") #고가
        # low = self.GetFieldData("OutBlock", "low") #저가
        print("원주문",ordno)
        ordno = ordno[-5:]
        expcode = expcode[3:-1]



        print()
        print("☆★☆★☆★☆★거래 완료 데이터 도착☆★☆★☆★☆★")
        print(ordno,trcode,chetime,chedate,chevol,cheprice,expcode,orgordno,dosugb)
        transaction_realtimeLog.append([ordno,trcode,chetime,chedate,chevol,cheprice,expcode,orgordno,dosugb])
        print("주문번호", "ordno",ordno)
        print("체결시간", "chetime",chetime)
        print("trcode", "trcode",trcode)
        print("체결일자", "chedate",chedate)
        print("체결량", "chevol",chevol)
        print("체결가격", "cheprice",cheprice)
        print("원주문번호", "expcode",expcode)
        print("옵션종목", "orgordno",orgordno)
        print("매도수 구분", "dosugb",dosugb)
        print("라인 일련번호", "lineseq",lineseq)
        print("일련번호", "seq",seq)
        print("매칭그룹번호", "megrpno",megrpno)
        print("보드ID", "boardid",boardid)
        print("세션ID", "sessionid",sessionid)
        print("약정번호", "yakseq",yakseq)

        print("☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★")
        print()

        if dosugb == "1": # 매도
            transaction_statusInfo[expcode]['finBidState'].append(ordno)
            transaction_statusInfo[expcode]['ongBidState'].remove(ordno)
        elif dosugb == "2": # 매수
            transaction_statusInfo[expcode]['finOfferState'].append(ordno)
            transaction_statusInfo[expcode]['ongOfferState'].remove(ordno)





        print("현재 거래중인 옵션은:",expcode)

        print()
        print("★★ Transaction - 개수 파악 ★★")
        print("HighStock", stockOpts['highStock'])
        print("ongBidState",transaction_statusInfo[stockOpts['highStock']]['ongBidState'] )
        print("ongOfferState", transaction_statusInfo[stockOpts['highStock']]['ongOfferState'])
        print("LowStock",stockOpts['lowStock'])
        print("ongBidState", transaction_statusInfo[stockOpts['lowStock']]['ongBidState'])
        print("ongOfferState", transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'])
        print()




    def start(self):
        """
        이베스트 서버에 실시간 data 요청함.
        """
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\C01.res"  # RES 파일 등록
        self.AdviseRealData()   # 실시간데이터 요청

    def end(self):
        self.UnadviseRealData()  # 실시간데이터 요청 모두 취소

    @classmethod
    def get_instance(cls):
        xreal = win32com.client.DispatchWithEvents("XA_DataSet.XAReal", cls)
        return xreal

class XAQueryEvents:
    상태 = False

    def OnReceiveData(self, szTrCode):
        print("OnReceiveData : %s" % szTrCode)
        XAQueryEvents.상태 = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("OnReceiveMessage : ", systemError, messageCode, message)

def CFOAT00100(계좌번호='',비밀번호='',선물옵션종목번호='0',매매구분='',선물옵션호가유형코드='',주문가격='',주문수량=''):
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    RESFILE = "C:\\eBEST\\xingAPI\\Res\\CFOAT00100.res"

    query.LoadFromResFile(RESFILE)
    query.SetFieldData("CFOAT00100InBlock1", "AcntNo", 0, 계좌번호) #계좌번호
    query.SetFieldData("CFOAT00100InBlock1", "Pwd", 0, 비밀번호) #비밀번호
    query.SetFieldData("CFOAT00100InBlock1", "FnoIsuNo", 0, 선물옵션종목번호) #선물옵션종목번호
    query.SetFieldData("CFOAT00100InBlock1", "BnsTpCode", 0, 매매구분) #매매구분
    query.SetFieldData("CFOAT00100InBlock1", "FnoOrdprcPtnCode", 0, 선물옵션호가유형코드) #선물옵션호가유형코드
    query.SetFieldData("CFOAT00100InBlock1", "OrdPrc", 0, 주문가격) #주문가격
    query.SetFieldData("CFOAT00100InBlock1", "OrdQty", 0, 주문수량) #주문수
    query.Request(0)

    print("갸수", len(transaction_statusInfo[stockOpts['highStock']]['ongOfferState']))
    print("갸수", len(transaction_statusInfo[stockOpts['lowStock']]['ongBidState']))

    # @1 매도(판매) @2 매수(구입)
    if 매매구분 == "1":
        transaction_statusInfo[선물옵션종목번호]['ongBidState'].append("TempOrdering")

    if 매매구분 == "2":
        transaction_statusInfo[선물옵션종목번호]['ongOfferState'].append("TempOrdering")

    print("Test prtin, 임시 거래")
    print(transaction_statusInfo[선물옵션종목번호]['ongBidState'])
    print(transaction_statusInfo[선물옵션종목번호]['ongOfferState'])


    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()

    result = []
    nCount = query.GetBlockCount("CFOAT00100InBlock1")
    transaction_statusInfo[stockOpts['highStock']]['ongOfferState']

    print()
    print("★★ Transaction - 거래 대기 중 ★★")
    print("HighStock", stockOpts['highStock'])
    print("ongBidState",transaction_statusInfo[stockOpts['highStock']]['ongBidState'] )
    print("ongOfferState", transaction_statusInfo[stockOpts['highStock']]['ongOfferState'])
    print("LowStock",stockOpts['lowStock'])
    print("ongBidState", transaction_statusInfo[stockOpts['lowStock']]['ongBidState'])
    print("ongOfferState", transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'])
    print()

    for i in range(nCount):
        레코드갯수 = query.GetFieldData("CFOAT00100OutBlock1", "RecCnt", i).strip() #레코드 갯수
        계좌번호 = query.GetFieldData("CFOAT00100OutBlock1", "AcntNo", i).strip() #계좌번호
        비밀번호 = query.GetFieldData("CFOAT00100OutBlock1", "Pwd", i).strip() #비밀번호
        매매구분 = query.GetFieldData("CFOAT00100OutBlock1", "BnsTpCode", i).strip() #매매구분
        주문번호 = query.GetFieldData("CFOAT00100OutBlock2", "OrdNo", i).strip() #주문번호
        OrdSeqno = query.GetFieldData("CFOAT00100OutBlock1", "OrdSeqno", i).strip() #매매구분
        Grpid = query.GetFieldData("CFOAT00100OutBlock1", "Grpid", i).strip() #주문번호
        PtflNo = query.GetFieldData("CFOAT00100OutBlock1", "PtflNo", i).strip() #매매구분
        BskNo = query.GetFieldData("CFOAT00100OutBlock1", "BskNo", i).strip() #주문번호
        TrchNo = query.GetFieldData("CFOAT00100OutBlock1", "TrchNo", i).strip() #매매구분
        ItemNo = query.GetFieldData("CFOAT00100OutBlock1", "ItemNo", i).strip() #주문번호

        FundId = query.GetFieldData("CFOAT00100OutBlock1", "FundId", i).strip() #매매구분
        FundOrdNo = query.GetFieldData("CFOAT00100OutBlock1", "FundOrdNo", i).strip() #주문번호

        lst = [레코드갯수,계좌번호,비밀번호, 매매구분,주문번호,OrdSeqno,Grpid,PtflNo,BskNo,TrchNo,ItemNo,FundId,FundOrdNo]

        result.append(lst)



        print(lst)
        if 매매구분 == "1": # 매도
            transaction_statusInfo[선물옵션종목번호]['ongBidState'].append(주문번호)
        elif 매매구분 == "2": # 매수
            transaction_statusInfo[선물옵션종목번호]['ongOfferState'].append(주문번호)




        print("Test print, 임시 갱신")
        print(transaction_statusInfo[선물옵션종목번호]['ongBidState'])
        print(transaction_statusInfo[선물옵션종목번호]['ongOfferState'])
        tempDict = {}
        tempDict["stoctCode"] = 선물옵션종목번호
        tempDict["transactionType"] = 매매구분
        tempDict["hoType"] = 선물옵션호가유형코드
        tempDict["ordPrice"] = 주문가격
        tempDict["ordVolume"] = 주문수량
        tempDict["OrgOrdNo"] = "origin"
        tempDict["OrdTime"] =  datetime.now().strftime("%H%M")

        transaction_detailedInfo[주문번호] = tempDict
        print("★ 데이터 추가했습니다.")
        print(transaction_detailedInfo[주문번호])



    if 매매구분 == "1":
        transaction_statusInfo[선물옵션종목번호]['ongBidState'].remove("TempOrdering")

    elif 매매구분 == "2":
        transaction_statusInfo[선물옵션종목번호]['ongOfferState'].remove("TempOrdering")

    print("현재 거래중인 옵션은:",선물옵션종목번호)

    print()
    print("★★ Transaction - 개수 파악 ★★")
    print("HighStock", stockOpts['highStock'])
    print("ongBidState",transaction_statusInfo[stockOpts['highStock']]['ongBidState'] )
    print("ongOfferState", transaction_statusInfo[stockOpts['highStock']]['ongOfferState'])
    print("LowStock",stockOpts['lowStock'])
    print("ongBidState", transaction_statusInfo[stockOpts['lowStock']]['ongBidState'])
    print("ongOfferState", transaction_statusInfo[stockOpts['lowStock']]['ongOfferState'])
    print()

    transaction_resultLog.append(lst)
    XAQueryEvents.상태 = False


def CFOAT00200(계좌번호='',비밀번호='',선물옵션종목번호='0',원주문번호='',선물옵션호가유형코드='',주문가격='',정정수량=''):
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    RESFILE = "C:\\eBEST\\xingAPI\\Res\\CFOAT00200.res"

    query.LoadFromResFile(RESFILE)
    query.SetFieldData("CFOAT00200InBlock1", "AcntNo", 0, 계좌번호) #계좌번호
    query.SetFieldData("CFOAT00200InBlock1", "Pwd", 0, 비밀번호) #비밀번호
    query.SetFieldData("CFOAT00200InBlock1", "FnoIsuNo", 0, 선물옵션종목번호) #선물옵션종목번호
    query.SetFieldData("CFOAT00200InBlock1", "OrgOrdNo", 0, 원주문번호) #원주문번호
    query.SetFieldData("CFOAT00200InBlock1", "FnoOrdprcPtnCode", 0, 선물옵션호가유형코드) #선물옵션호가유형코드
    query.SetFieldData("CFOAT00200InBlock1", "OrdPrc", 0, 주문가격) #주문가격
    query.SetFieldData("CFOAT00200InBlock1", "MdfyQty", 0, 정정수량) #정정수량
    # 매매 구분은 기존의 값을 유지해야함.
    query.Request(0)

    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()

    매매구분 = transaction_detailedInfo[원주문번호]['transactionType']

    if 매매구분 == "1": # 매도
        transaction_statusInfo[선물옵션종목번호]['finBidState'].append(원주문번호)
        transaction_statusInfo[선물옵션종목번호]['ongBidState'].remove(원주문번호)
    elif 매매구분 == "2": # 매수
        transaction_statusInfo[선물옵션종목번호]['finOfferState'].append(원주문번호)
        transaction_statusInfo[선물옵션종목번호]['ongOfferState'].remove(원주문번호)


    result = []
    nCount = query.GetBlockCount("CFOAT00100InBlock1")

    for i in range(nCount):
        레코드갯수 = query.GetFieldData("CFOAT00200OutBlock1", "RecCnt", i).strip()
        계좌번호 = query.GetFieldData("CFOAT00200OutBlock1", "AcntNo", i).strip()
        비밀번호 = query.GetFieldData("CFOAT00200OutBlock1", "Pwd", i).strip()
        선물옵션종목번호 = query.GetFieldData("CFOAT00200OutBlock1", "FnolsuNo", i).strip()
        선물옵션호가유형코드 = query.GetFieldData("CFOAT00200OutBlock1","FnoOrdprcPtnCode",i).strip()
        원주문번호 = query.GetFieldData("CFOAT00200OutBlock1","OrgOrdNo",i).strip()
        주문번호 = query.GetFieldData("CFOAT00200OutBlock2", "OrdNo", i).strip()
        주문가격 = query.GetFieldData("CFOAT00200OutBlock1","OrdPrc",i).strip()
        주문수량 = query.GetFieldData("CFOAT00200OutBlock2", "MdfyQty", i).strip()


        lst = [레코드갯수,계좌번호,비밀번호, 매매구분,주문번호,원주문번호]
        result.append(lst)

        if 매매구분 == "1": # 매도
            transaction_statusInfo[선물옵션종목번호]['ongBidState'].append(주문번호)
        elif 매매구분 == "2": # 매수
            transaction_statusInfo[선물옵션종목번호]['ongOfferState'].append(주문번호)



        tempDict = {}
        tempDict["stoctCode"] = 선물옵션종목번호
        tempDict["transactionType"] = 매매구분
        tempDict["hoType"] = 선물옵션호가유형코드
        tempDict["ordPrice"] = 주문가격
        tempDict["ordVolume"] = 주문수량
        tempDict["OrgOrdNo"] = 원주문번호

        transaction_detailedInfo[주문번호] = tempDict

        transaction_resultLog.append(lst)
    XAQueryEvents.상태 = False


"""
01. 유저 정보를 선택한다. (계좌, ID, 비밀번호 등)
"""

def read_and_choose_userInfo(filePath):
    print("===========================================================================")
    print("==============================프로그램 변수입력============================")
    print("===========================================================================")
    userInfo_df = pd.read_csv(filePath, encoding='cp949')
    print("\n",userInfo_df)
    userInfo_index = input("\n※ 사용하실 계좌를 선택해주십시오:")
    userInfo_index = int(userInfo_index)
    selectedUserInfo_list = userInfo_df.loc[userInfo_index].tolist()
    print(selectedUserInfo_list,"\n")  # ['모의', 555510.0, 'id', 'PWD', 'PWD', 'PWD', 'demo.ebestsec.co.kr']
    selectedUserInfo_dict = {"type":selectedUserInfo_list[0],"account_num":str(int(selectedUserInfo_list[1])),"user_id":selectedUserInfo_list[2],"password":selectedUserInfo_list[3],"cert_password":selectedUserInfo_list[4],"URL":selectedUserInfo_list[6]}
    return selectedUserInfo_dict # dictionary type

"""
02. 거래할 두 개의 옵션 종목을 선택한다.
"""

def read_and_choose_stockOpts(filePath):
    stockOpts_df = pd.read_csv(stockOpts_filepath, encoding='cp949')
    print(stockOpts_df,"\n")
    stockOpts_index = input("\n※ 사용하실 코드를 선택해주십시오:")
    stockOpts_list = [i for i in stockOpts_df.loc[int(stockOpts_index)].tolist() if i == i] # 리스트로 변경된 부분
    print("※ 다운로드 받을 옵션 코드는 다음과 같습니다.")
    print(stockOpts_list,"\n")  # ['201QA327', '201QA332']
    selectedUserInfo_dict = {"highStock":stockOpts_list[0], "lowStock":stockOpts_list[1]}
    return selectedUserInfo_dict  # dictionary type

    ####################################################
    # 여기서 어떤 주식이 가격이 높은 지 설정해주게 하고 싶#

"""
03. 프로그램을 돌릴 시간을 입력한다.
"""
def input_time_limit():
    print("※ 언제까지 데이터를 수신할까요? \n 입력 예제) 3시 15분 - 1515")
    due_time = input()
    print("※",due_time,"까지 데이터를 수신합니다.")
    print("\n")

    print("===========================================================================")
    print("===========================프로그램 변수입력완료===========================")
    print("===========================================================================")
    print("\n")
    return due_time # string type


"""
04. 시스템 변수를 출력한다.
"""
def print_system_variables(systemVariables_dict):
        print("현재 시각: ",systemVariables_dict["currentTime"])
        print("종료예약 시각: ",systemVariables_dict["dueTime"])
        print("oldCount: ",systemVariables_dict["oldCount"])
        print("upperCaseCount: ",systemVariables_dict["upperCaseCount"])
        print("lowerCaseCount: ",systemVariables_dict["lowerCaseCount"])

"""
05. 데이터프레임을 저장한다.
"""
def save_dataframe_to_file(input_list,file_name,columnsName_list=[]):
    input_df = pd.DataFrame(input_list)
    current_time_stamp = datetime.now().strftime("%Y%m%d%H-%M-%S")
    if (len(columnsName_list) > 1) and (input_df.shape[1] == len(columnsName_list)):
        input_df.to_excel(current_time_stamp+"_"+file_name+".xlsx",columns=columnsName_list,sheet_name="output")

    else:
        input_df.to_excel(current_time_stamp+"_"+file_name+".xlsx",sheet_name="output")

    print(current_time_stamp+"_"+file_name+".xlsx","  이름으로 저장 완료했습니다.")
        # ## 이쪽에 column names 길이랑 실제 df 길이랑 비교하는 거 넣을 것

if __name__ == "__main__":
        """
        01. 필요한 변수를 미리 생성한다.
        """
        userInfo_filePath = "./secret/passwords.csv"
        stockOpts_filepath = "./secret/code_list.csv"

        systemVariables = {}
        systemVariables['currentTime'] = ""
        systemVariables['dueTime'] = ""
        systemVariables['oldCount'] = 0
        systemVariables['lowerCaseCount'] = 0
        systemVariables['upperCaseCount'] = 0
        # Real 데이터를 몇 개 받았는 지 count하기 위해서 변수(Integer) 생성
        # 현재 시간을 기록하기 위해 변수(String) 생성
        # "0.03 이상 혹은 -0.03 이하가 되는 횟수"를 저장할 변수를 생성 (Integer)

        """
        01. 유저 정보를 선택한다. (계좌, ID, 비밀번호 등)
        """
        userInfo = read_and_choose_userInfo(userInfo_filePath)

        """
        02. 거래할 두 개의 옵션 종목을 선택한다.
        """
        stockOpts = read_and_choose_stockOpts(stockOpts_filepath) # 2가지, {"highStock":, "lowStock":}

        """
        03. 프로그램을 돌릴 시간을 입력한다.
        """
        systemVariables["dueTime"] = input_time_limit() #0120

        """
        04. 실거래를 선택한 경우, 실제로 주문을 넣을 지 선택한다. (추가 예정)
        """
        stockOpts_statusInfo = {} # 각각 옵션코드의 5개의 정보(curBidho, curOfferho, preBidgo,...)
        stockOpts_statusLog = {} #{optCode:[]}

        stockOpts_indicatorInfo = {} # 각각 옵션코드의 2개의 정보(difference, deviation)
        stockOpts_indicatorLog = {} # 생성중
        stockOpts_transactionStatus = {} #거래 원주문번호
        # record_of_each_hoprice = {} # 각각 옵션에 대해서 7개의 정보(bidho, offerho, count, bid_status, offerstatus, bidho_average, offerho_average)
        indicator_names = ['deviationLogic']

        transaction_statusInfo = {}
        transaction_detailedInfo = {}

        for idx,stockOpt in enumerate(stockOpts.values()):
            # dictionary_of_code[option_code] = ordinal_numbers[idx]
            stockOpts_statusInfo[stockOpt] =  {"curBidho":-9999.0,"curOfferho":-9999.0,"preBidho":-9999.0,"preOfferho":-9999.0,"price":-9999.0,'avgBidho':0.0,'avgOfferho':0.0,'midHo':0.0}
            stockOpts_statusLog[stockOpt] = {"Bidho":[],"Offerho":[],"price":[],'midHo':[]}
            transaction_statusInfo[stockOpt] = {"ongBidState":[],"finBidState":[],"ongOfferState":[],"finOfferState":[]}

            # record_of_each_hoprice[option_code] = {'bidho': [] , 'offerho': [], 'count': 0, 'bid_status' : False, 'offer_status' : False, 'bidho_average' : 0, 'offerho_average' : 0}

        for indicator_name in indicator_names:
            stockOpts_indicatorInfo[indicator_name] = {'curValue':0.0,'preValue':0.0,'avgValue':0.0}
            stockOpts_indicatorLog[indicator_name] = {'log':[]}

        stockOpts_realtimeLog = []
        transaction_resultLog = []
        transaction_realtimeLog = []
        current_time_int = 7

        isOver150 = False

        """
        06. 서버에 로그인 한다.
        """
        xsession = XSession.get_instance() # 로그인 세션 활성화
        xsession.api_login(id = userInfo["user_id"], pwd = userInfo["password"], cert_pwd = userInfo["cert_password"], url = userInfo["URL"]) # 로그인 정보 입력

        """
        07. 입력받은 옵션에 대한 TR 데이터 요청 후 수신한다.
        """
        print("=======================================================================")
        print("===========================옵션 초기값 수신============================")
        print("=======================================================================")
        print("\n")

        for stockOpt in stockOpts.values(): ## 이 부분은 upper와 down을 각각 하던지 판단해야한다.
            t2105_TRquery = XQuery_t2105.get_instance()
            # t2105_TRquery.set_data("t2105",stockOpt)
            t2105_TRquery.request(stockOpt)

            while t2105_TRquery.is_data_received == False:
                pythoncom.PumpWaitingMessages()

        print("\n")
        print("=======================================================================")
        print("=========================옵션 초기값 갱신완료==========================")
        print("=======================================================================")

        """
        08. 옵션 거래 결과를 실시간 수신하기 위해 real data를 연결한다.
        """
        C01_Realquery = XReal_C01.get_instance()
        C01_Realquery.start()

        """
        09. 옵션의 실시간 호가 변화를 수신하기 위해 real data를 연결한다.
        """
        print(list(stockOpts.values())[0])
        OC0_Realquery = XReal_OC0_.get_instance()
        print(type(OC0_Realquery))
        OC0_Realquery.set_data(list(stockOpts.values())[0])
        OC0_Realquery.start()
        print("리얼쿼리 실행")

        print(list(stockOpts.values())[1:])
        for stockOpt in list(stockOpts.values())[1:]:
            OC0_Realquery.add_item(stockOpt)
            print(stockOpt)


        """
        10. 옵션의 실시간 호가 변화를 수신하기 위해 real data를 연결한다.
        """

        while systemVariables["currentTime"] != systemVariables["dueTime"]: # 입력한 시간에 Real 데이터 수신을 종료한다.
            systemVariables["currentTime"] = datetime.now().strftime("%H%M")
            pythoncom.PumpWaitingMessages()

            if systemVariables["currentTime"] == systemVariables["dueTime"]:
                OC0_Realquery.end()  # 실시간 조회 중단.
                print("---- 시스템 변수를 출력합니다. -----")
                print_system_variables(systemVariables)
                time.sleep(10)
                print("-----  프로그램을 종료합니다. ------")

                # save_dataframe_to_file(리얼타임 df 추가할 것)
                #
                # order_df = pd.DataFrame(data=transaction_resultLog, columns=['레코드갯수', '계좌번호', '비밀번호', '매매구분', '주문번호'])
                # order_df.to_excel(datetime.now().strftime("%Y%m%d%H-%M-%S")+"_ordered_ReceivedData.xlsx",sheet_name="test")



                # complete_df = pd.DataFrame(data=transaction_realtimeLog)
                # complete_df.to_excel(datetime.now().strftime("%Y%m%d%H-%M-%S")+"_complete_ReceivedData.xlsx",sheet_name="test")
                #
                # differ_df = pd.DataFrame([differList,differList2])
                # differ_df.to_excel(datetime.now().strftime("%Y%m%d%H-%M-%S")+"_"+code_info[0]+"-"+code_info[1]+"_differ_value.xlsx",sheet_name="test")

                # stockOpts_realtimeLog
                save_dataframe_to_file(stockOpts_realtimeLog,"RealtimeLog",columnsName_list=[])


                # save_dataframe_to_file(input_list,file_name,columnsName_list=[])
                #
                #
                # save_dataframe_to_file(input_list,file_name,columnsName_list=[])



                break
