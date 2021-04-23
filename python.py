import requests
import pandas 
import re
from io import BytesIO

headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.128 Safari/537.36"}

def download_excel(rcp_no, dcm_no, period, company):
    #
    url = "http://dart.fss.or.kr/pdf/download/excel.do?rcp_no={}&dcm_no={}&lang=ko".format(rcp_no, dcm_no)
    res = requests.get(url , headers=headers)
    res.raise_for_status()

    table = BytesIO(res.content)
    sheetName = ["연결 재무상태표","연결 손익계산서","연결 포괄손익계산서"]
    #data = pandas.read_excel(table, sheet_name="연결 재무상태표" , skiprows=5)
    #data.to_csv("test.csv" , encoding = "euc-kr")
    for sheet in sheetName:
        data = pandas.read_excel(table, sheet_name=sheet , skiprows=5)
        data.to_csv("{}_{}_{}.csv".format(company,period,sheet) , encoding = "euc-kr")

#API 이용하여 rcp_no, dcm_no 얻기
#API = 요청을 받고 응답을 받을때 규정되어있는 rule

api_key = "4e5fddad078b4a16dba62bfe1193d4194ec97847"
corp_code = "005930"
url = "https://opendart.fss.or.kr/api/list.xml?crtfc_key={}&corp_code={}&bgn_de=20150101&pblntf_ty=A&pblntf_detail_ty=A001&pblntf_detail_ty=A002&pblntf_detail_ty=A003&page_count=100".format(api_key, corp_code)
res = requests.get(url , headers=headers)
res.raise_for_status()
apiResult = res.content.decode("UTF-8")

corpNameArr = re.findall(r'<corp_name>(.*?)</corp_name>',apiResult )
rptNampeArr = re.findall(r'<report_nm>(.*?)</report_nm>',apiResult )
rcpNoArr = re.findall(r'<rcept_no>(.*?)</rcept_no>',apiResult )
dcmNoArr = []
# dict(zip(rptNampeArr ,rcpNoArr))
for rcpNo in rcpNoArr:
    url = "http://dart.fss.or.kr/dsaf001/main.do?rcpNo={}".format(rcpNo)
    res = requests.get(url , headers=headers)
    res.raise_for_status()
    apiResult = res.content.decode("UTF-8")
    dcmNoArr.append(re.findall(r"'{}', '(.*?)'".format(rcpNo),apiResult )[0])

for corpName, rptName, rcpNO, dcmNo in zip(corpNameArr,rptNampeArr,rcpNoArr,dcmNoArr):
    download_excel(rcpNO, dcmNo, rptName, corpName)