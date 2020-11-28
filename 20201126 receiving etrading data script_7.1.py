#-*-coding: utf-8 -*-

import win32com.client
import pythoncom
import time
import pandas as pd
import sqlite3
from datetime import datetime

# 로그인 역할을 수행하는 클래스 생성777
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
                                                    # ??????????????????????????????????????

        self.ConnectServer(url, 20001)                      # 커넥트서버에 접속하고
        is_connected = self.Login(id, pwd, cert_pwd, 0, False)  # 로그인 하기 ..로그인 상태를 is_connected 에 저장

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
    category_code = ""  #왜 빈칸으로 변수를 만드는 가?  왜 이 클래스에 만들어서 다른 클래스에서도 쓰는가???????????????????????????????????????????????????????????????????
    option_code = ""  ## __init__ 에 들어 가야 할 수도

    def __init__(self):
        super().__init__()
        self.is_data_received = False

    def set_data(self,category_code,option_code): # 이 함수의 역할이 머지??   ## 강의 50min 근처  ?
        self.category_code = category_code ## 삭제 예정??????? 카테고리 코드는 t2105일 가능성   확인 /필요없는 것인가?  325에서 이리 옴
        self.option_code = option_code  # 62라인에 있는 것을 self를 붙여서 넣는 이유?

    def OnReceiveData(self, tr_code):
        self.is_data_received = True
        # 사용 항목
        price = self.GetFieldData("t2105OutBlock", "price", 0)
        offerho1 = self.GetFieldData("t2105OutBlock", "offerho1", 0)
        bidho1 = self.GetFieldData("t2105OutBlock", "bidho1", 0)
        hotime = self.GetFieldData("t2105OutBlock", "time", 0)

        # 비사용 항목
        # name = self.GetFieldData("t2105OutBlock", "hname", 0)
        # volume = self.GetFieldData("t2105OutBlock","volume",0)
        # offerrem1 = self.GetFieldData("t2105OutBlock", "offerrem1", 0)
        # bidrem1 = self.GetFieldData("t2105OutBlock", "bidrem1", 0)
        # dcnt1 = self.GetFieldData("t2105OutBlock", "dcnt1", 0)
        # scnt1 = self.GetFieldData("t2105OutBlock", "scnt1", 0)

        print("TR code는 {0}이며, 옵션 코드는 {1}입니다.".format(tr_code,self.option_code))
        print("종목은 {0},\n현재가는 {1},\n 거래가, {2},\n 매도호가1은 {3},\n 매수호가1은  {4}\n".format(name, hotime, price, offerho1, bidho1))
        return_real_items.append([self.option_code, hotime, price, offerho1, bidho1])

        status_bundle[self.option_code] =  {"curBidho":bidho1,"curOfferho":offerho1,"preBidho":bidho1,"preOfferho":offerho1,"price":price}



    def request(self):
        option_code = self.option_code
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t2105.res"  # RES 파일 등록
        self.SetFieldData("t2105InBlock", "shcode", 0, self.option_code)

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

    def set_data(self,category_code,option_code): ## from 384
        self.category_code = category_code        ##
        self.option_code = option_code
                                            # 메써드는 반드시 불러야 실행이 된다,,안불러도 실행 되는 것이 아니고 서버에서 메써드 실행 시킨것
    def OnReceiveRealData(self, tr_code):  # event handler  ## 이베스트에서 신호를 줘서 실행 된다/ 그쪽에서 코드를 실행했다  겟인스턴스는 그쪽 것을 받아 온것인데...
        """
        이베스트 서버에서 ReceiveRealData 이벤트 받으면 실행되는 event handler
        """

        self.count += 1

        # 사용 항목
        optcode = self.GetFieldData("OutBlock","optcode")
        chetime = self.GetFieldData("OutBlock", "chetime") # 체결시간
        price = self.GetFieldData("OutBlock", "price") # 현가
        offerho1 = self.GetFieldData("OutBlock", "offerho1") # 매도 1호가(판매)
        bidho1 = self.GetFieldData("OutBlock", "bidho1") # 매수 1호가(구입)

        # 비사용 항목
        # mdvolume = self.GetFieldData("OutBlock", "mdvolume") # 매도누적체결량
        # mdchecnt = self.GetFieldData("OutBlock", "mdchecnt") # 매도누적체결건수
        # msvolume = self.GetFieldData("OutBlock", "msvolume")  # 매수누적체결량
        # mschecnt = self.GetFieldData("OutBlock", "mschecnt") #매수누적체결건수
        # open = self.GetFieldData("OutBlock", "open") #시가
        # high = self.GetFieldData("OutBlock", "high") #고가
        # low = self.GetFieldData("OutBlock", "low") #저가


        print("아래 데이터를 수신했습니다.")
        print(optcode,chetime,price,offerho1,bidho1)
        print("\n")

        if status_bundle
            status_bundle[optcode]["preOfferho"] = status_bundle[optcode]["curOfferho"]
            status_bundle[optcode]["preBidho"] = status_bundle[optcode]["curBidho"]

            status_bundle[optcode]["curOfferho"] = float(offerho1)
            status_bundle[optcode]["curBidho"] = float(bidho1)
            status_bundle[optcode]["price"] = float(price)

        else:
            print("일치하는 항목이 없습니다.")
            print("\n")


        option_A_bidho_items.append(statesOfOptionA["curBidho"])
        option_A_offerho_items.append(statesOfOptionA["curOfferho"])
        option_B_bidho_items.append(statesOfOptionB["curBidho"])
        option_B_offerho_items.append(statesOfOptionB["curOfferho"])
        global isOver150
        global option_A_offer_state
        global option_B_bid_state
        global option_A_bid_state
        global option_B_offer_state

        if isOver150 == False:
            if len(option_A_bidho_items) > 150:
                isOver150 = True
            return_real_items.append([-9999,-9999,chetime,statesOfOptionA["price"],statesOfOptionA["curOfferho"],statesOfOptionA["curBidho"],statesOfOptionB["price"],statesOfOptionB["curOfferho"],statesOfOptionB["curBidho"]])

        elif isOver150 == True:
            print("150 이상 돌입")
            option_A_bidho_items.pop(0)
            option_A_offerho_items.pop(0)
            option_B_bidho_items.pop(0)
            option_B_offerho_items.pop(0)

            optionA_bidhoAverage = sum(option_A_bidho_items) / len(option_A_bidho_items) # a 매수호가 평균
            optionA_offerhoAverage = sum(option_A_offerho_items) / len(option_A_offerho_items) # a 매도호가 평균
            optionB_bidhoAverage = sum(option_B_bidho_items) / len(option_B_bidho_items) # b 매수호가 평균
            optionB_offerhoAverage = sum(option_B_offerho_items) / len(option_B_offerho_items) # b 매도호가 평균


        # print(statesOfOptionA[0],statesOfOptionB[3]) # A옵션 매도 B 옵션 매수 호가  1.00   0.49
        # print(statesOfOptionB[0],statesOfOptionA[3]) # B옵션 매도호 A옵션 매수호    0.5   0.99
            Bid_A_Offer_B =  optionA_bidhoAverage - (optionB_offerhoAverage * 2) #  B to A 상태  : a매수호가 b 매도호가 b스테이지에서 a 스테이지 갈때 참고가격 0.99-0.5*2
            Bid_B_Offer_A = (optionB_bidhoAverage*2) - optionA_offerhoAverage #  A to B 상태  : A매도호가 B 매수호가 b스테이지에서 a 스테이지 갈때 참고가격    1.0-0.49*2
            # deviation = Bid_A_Offer_B + Bid_B_Offer_A # ??????????????? 오류 이 값은 항상 0.01
            deviation =  ((statesOfOptionB["curBidho"]*2)-statesOfOptionA["curOfferho"] )- Bid_B_Offer_A  # 현재가의 차에서 평균의 차를 뺀값
            differList.append(deviation)  ## 이부분을 처음부터 기록으로 남겨서 엑셀에 시간과 함께 기록
            difference =(statesOfOptionB["curBidho"]*2)-statesOfOptionA["curOfferho"]
            differList2.append(difference)

            # A가 더 비싸다고 가정
            print("H1_L2",Bid_A_Offer_B) # 비싼것 매도호에 사고  싼것 두개 매수호 산다  비용으로 계산한것  1.00-0.49*2 =0.02 비용 발생
            print("L2_H1",Bid_B_Offer_A) # 싼것 두개 사고 비싼것 하나 판다                             0.5*2 -0.99 = 0.01 비용 발생
            print("\n Deviation 값입니다.",deviation) # 대부분 0.04가 정상인데,,,
            print("\n Difference 값입니다.", difference)

            return_real_items.append([deviation,difference,chetime,statesOfOptionA["price"],statesOfOptionA["curOfferho"],statesOfOptionA["curBidho"],statesOfOptionB["price"],statesOfOptionB["curOfferho"],statesOfOptionB["curBidho"]])

            if differList[-1] >= 0.03 and differList[-2] < 0.03:   #
                if (option_A_offer_state == False) and (option_B_bid_state == False):
                    print(first_option_code[0],"1 개 매도(판매) 주문완료")
                    CFOAT00100(계좌번호=str(int(id_info[1])),비밀번호=id_info[4],선물옵션종목번호=option_codes[0],매매구분="1",선물옵션호가유형코드="00",주문가격=statesOfOptionA['curOfferho'],주문수량='1') #
                    print(first_option_code[1],"2 개 매수(구매) 주문완료")
                    CFOAT00100(계좌번호=str(int(id_info[1])),비밀번호=id_info[4],선물옵션종목번호=option_codes[1],매매구분="2",선물옵션호가유형코드="00",주문가격=statesOfOptionB['curBidho'],주문수량='2') #
                    #@1 매도(판매) @2 매수(구입)
                    global upperCaseCount
                    upperCaseCount += 1


                    option_A_offer_state = True
                    option_B_bid_state = True

            elif differList[-1] <= -0.03 and differList[-2] > -0.03:
                if (option_A_bid_state == False) and (option_B_offer_state == False):
                    print(first_option_code[0],"2 개 매도(판매) 주문완료")
                    CFOAT00100(계좌번호=str(int(id_info[1])),비밀번호=id_info[4],선물옵션종목번호=option_codes[1],매매구분="1",선물옵션호가유형코드="00",주문가격=statesOfOptionB['curOfferho'],주문수량='2')
                    print(first_option_code[1],"1 개 매수(구매) 주문완료")
                    CFOAT00100(계좌번호=str(int(id_info[1])),비밀번호=id_info[4],선물옵션종목번호=option_codes[0],매매구분="2",선물옵션호가유형코드="00",주문가격=statesOfOptionA['curBidho'],주문수량='1')
                    global lowerCaseCount
                    lowerCaseCount += 1

                    option_A_bid_state = True
                    option_B_offer_state = True

            else:
                print("주문없이 종료시켰습니다.")
            print("\n")

    def start(self):
        """
        이베스트 서버에 실시간 data 요청함.
        """
        self.ResFileName = "C:\\eBEST\\xingAPI\\Res\\OC0.res"  # RES 파일 등록
        self.SetFieldData("InBlock", "optcode",  self.option_code)
        self.AdviseRealData()   # 실시간데이터 요청



    def add_item(self, stockcode):                          # 이것이 왜 필요하지??????????????????????????????????
        # 실시간데이터 요청 종목 추가
        self.SetFieldData("InBlock", "optcode", stockcode)  ## stock 코드 맞는가? 옵션코드 아닌가 확인    365에서 옴  ???????????????????????????????????
        self.AdviseRealData()                               ## 실시간 리얼데이타를 등록한다

    def remove_item(self, stockcode):
        # stockcode 종목만 실시간데이터 요청 취소
        self.UnadviseRealDataWithKey(stockcode)

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






def CFOAT00100(계좌번호='',비밀번호='',선물옵션종목번호='0',매매구분='',선물옵션호가유형코드='',주문가격='',주문수량=''):   ############여기는 왜 함수가 클래스급인가????????????????????????????????????????????????????????????????????

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    RESFILE = "C:\\eBEST\\xingAPI\\Res\\CFOAT00100.res"

    query.LoadFromResFile(RESFILE)
    query.SetFieldData("CFOAT00100InBlock1", "AcntNo", 0, 계좌번호)
    query.SetFieldData("CFOAT00100InBlock1", "Pwd", 0, 비밀번호)
    query.SetFieldData("CFOAT00100InBlock1", "FnoIsuNo", 0, 선물옵션종목번호)
    query.SetFieldData("CFOAT00100InBlock1", "BnsTpCode", 0, 매매구분)
    query.SetFieldData("CFOAT00100InBlock1", "FnoOrdprcPtnCode", 0, 선물옵션호가유형코드)
    query.SetFieldData("CFOAT00100InBlock1", "OrdPrc", 0, 주문가격)
    query.SetFieldData("CFOAT00100InBlock1", "OrdQty", 0, 주문수량)
    query.Request(0)

    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()

    result = []
    nCount = query.GetBlockCount("CFOAT00100InBlock1")
    for i in range(nCount):
        레코드갯수 = query.GetFieldData("CFOAT00100OutBlock1", "RecCnt", i).strip()
        계좌번호 = query.GetFieldData("CFOAT00100OutBlock1", "AcntNo", i).strip()
        비밀번호 = query.GetFieldData("CFOAT00100OutBlock1", "Pwd", i).strip()
        매매구분 = query.GetFieldData("CFOAT00100OutBlock1", "BnsTpCode", i).strip()

        lst = [레코드갯수,계좌번호,비밀번호,매매구분]
        result.append(lst)
        print(lst)
    df = pd.DataFrame(data=result, columns=['레코드갯수', '계좌번호', '비밀번호', '매매구분'])

    XAQueryEvents.상태 = False

    # return (df, df1)
    return (df)







if __name__ == "__main__":

        # Password 파일에서 원하는 계좌정보 선택
        info_df = pd.read_csv("./secret/passwords.csv", encoding='cp949')
        print("\n",info_df)
        info_Num = input("\n※ 사용하실 계좌를 선택해주십시오:")
        id_info = info_df.loc[int(info_Num)].tolist()
        print(id_info,"\n")

        # code_list 파일에서 원하는 옵션 코드 불러오기
        code_list_df = pd.read_csv("./secret/code_list.csv", encoding='cp949')
        print(code_list_df,"\n") # 이부분 주석처리 가능---code_list_df는 데이타프레임
        code_Num = input("\n※ 사용하실 코드를 선택해주십시오:")
        option_codes = code_list_df.loc[int(code_Num)].tolist() # 리스트로 변경된 부분
        print("※ 다운로드 받을 옵션 코드는 다음과 같습니다.(벌크 다운로드용 기능입니다.)")
        print(option_codes,"\n")                          # ['201QA327', '201QA332']

        # 몇시 몇분까지 데이터를 받아올 지 입력
        print("※ 언제까지 데이터를 수신할까요? \n 입력 예제) 3시 15분 - 1515")
        due_time = input()
        print("※",due_time,"까지 데이터를 수신합니다.\n")


        ## 데이터를 받을 코드 단일(String) 생성
        first_option_code = option_codes[0] # 삭제예정
        print(option_codes,"\n")

        # "0.03 이상 혹은 -0.03 이하가 되는 횟수"를 저장할 변수를 생성 (Integer)
        lowerCaseCount = 0
        upperCaseCount = 0

        status_bundle = {}
        statistics_bundle = {}
        dictionary_of_code = {}
        ordinal_numbers = ['first','second','third','fourth','fifth','sixth','seventh','eighth','ninth','tenth','eleventh','twelfth','thirteenth','fourteenth','fifteenth']

        for idx,option_code in enumerate(option_codes):
            dictionary_of_code[option_code] = ordinal_numbers[idx]
            status_bundle[option_code] =  {"curBidho":-9999.0,"curOfferho":-9999.0,"preBidho":-9999.0,"preOfferho":-9999.0,"price":-9999.0}

        for log_title in ['difference_log','deviation_log']:
            statistics_bundle[log_title] = {option_code:[-9999,-9999] for option_code in option_codes}



        # 두 옵션의 차이를 저장할 묶음(List) 생성
        differList = [-9999,-9999]
        differList2 = [-9999,-9999]
        # 리얼 데이터를 저장할 묶음(List)
        return_real_items = []



        option_A_bidho_items = []
        option_A_offerho_items = []
        option_B_bidho_items = []
        option_B_offerho_items = []

        isOver150 = False

        option_A_bid_state = False
        option_A_offer_state = False
        option_B_bid_state = False
        option_B_offer_state = False


        optionA_bidhoAverage = 0
        optionA_offerhoAverage = 0
        optionB_bidhoAverage = 0
        optionB_offerhoAverage = 0


        # 로그인 세션
        xsession = XSession.get_instance()

        # 로그인 정보 입력
        xsession.api_login(id=id_info[2], pwd=id_info[3], cert_pwd=id_info[4], url=id_info[6])

        # 입력받은 옵션에 대한 TR 데이터 요청 후 수신
        for option_code in option_codes:
            query = XQuery_t2105.get_instance()
            query.set_data("t2105",option_code)
            query.request()

            while query.is_data_received == False:
                pythoncom.PumpWaitingMessages()


        # 입력받은 옵션에 대한 Real 데이터 요청 후 수신
        xreal = XReal_OC0_.get_instance()
        xreal.set_data("t2105",option_codes[0])
        xreal.start()

        for option_code in option_codes[1:]:
            xreal.add_item(option_code)


        # Real 데이터를 몇 개 받았는 지 count하기 위해서 변수(Integer) 생성
        old_count = 0

        # 현재 시간을 기록하기 위해 변수(String) 생성
        current_time = ""

        # 입력한 시간에 Real 데이터 수신을 종료한다.
        while current_time != due_time:
            current_time = datetime.now().strftime("%H%M")
            pythoncom.PumpWaitingMessages()

            if current_time == due_time:
                xreal.end()  # 실시간 조회 중단.
                print("upperCaseCount: ",upperCaseCount)
                print("lowerCaseCount: ",lowerCaseCount)
                time.sleep(10)
                print("---- end -----")


                # if old_count < xreal.count:
                #     old_count = xreal.count
                realtime_df = pd.DataFrame(return_real_items)
                realtime_df.columns = ['optcode','chetime','price','offerho1','bidho1']
                realtime_df.to_excel(datetime.now().strftime("%Y%m%d%H-%M-%S")+"_"+option_codes[0]+"-"+option_codes[1]+"_ReceivedData.xlsx",sheet_name="test")


                differ_df = pd.DataFrame([differList,differList2])
                differ_df.to_excel(datetime.now().strftime("%Y%m%d%H-%M-%S")+"_"+option_codes[0]+"-"+option_codes[1]+"_differ_value.xlsx",sheet_name="test")
                break
