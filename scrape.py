from lxml import html
from openpyxl import Workbook
from openpyxl import load_workbook
import os.path
import requests

#Posts the data to Coh Metrix website
def get_data():
    s = requests.Session()
    starter = s.get('http://141.225.42.101/cohmetrix3/login.aspx')
    tree = html.fromstring(starter.text)

    cookies = {
        'ASP.NET_SessionId': starter.cookies['ASP.NET_SessionId']
    }

    headers = {
        'Host': '141.225.42.101',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.10; rv:41.0) Gecko/20100101 Firefox/41.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': 'http://141.225.42.101/cohmetrix3/login.aspx',
        'Connection': 'keep-alive',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    viewstate = tree.xpath('//input[@id="__VIEWSTATE"]/@value')[0]
    validation = tree.xpath('//input[@id="__VIEWSTATEGENERATOR"]/@value')[0]
    view = tree.xpath('//input[@id="__EVENTVALIDATION"]/@value')[0]

    data = {
        '__VIEWSTATE': viewstate,
        '__VIEWSTATEGENERATOR': validation,
        '__EVENTVALIDATION': view,
        'TextBox1': 'glenn.saqui@gmail.com',
        'TextBox2': "bebZjcis",
        'Button1': 'Login',
    }
    l = s.post('http://141.225.42.101/cohmetrix3/login.aspx', headers=headers, data=data, cookies=cookies)

    headers = {
        'Host': '141.225.42.101',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.10; rv:41.0) Gecko/20100101 Firefox/41.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': 'http://141.225.42.101/cohmetrix3/input.aspx',
        'Connection': 'keep-alive',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    viewstate = "/wEPDwUJNzU4MDYzOTUyZGQKNmk5nIq1XYBgXUM1EZ1Hs0bJWQ=="
    validation = '/wEWCgLrrJCoDQK506zcAwLy07LaDALjz837AwLBx6n3AwLptdyVCAL0geaDCwLCoveuCALpouS8DQLCi9reA0rdDT2rbbn4VrlhjGW0OKnV7hBW'

    tempData = {
        '__VIEWSTATE': viewstate,
        '__VIEWSTATEGENERATOR': 'AD710423',
        '__EVENTVALIDATION': validation,
        'tbTitle': 'test',
        'ddlGenre': 'Science',
        'tbSource': 'source',
        'tbUserCode': 'job+code',
        'ddlLSASpace': 'CollegeLevel',
        'tbInput': 'this+is+some+really+nice+code',
        'btnSubmit': 'Submit',
    }

    t = s.post('http://141.225.42.101/cohmetrix3/input.aspx', headers=headers, cookies=cookies, data=tempData)
    return t.text


# parses the data from the html and sticks it into the excel spreadsheet
def parse_data(text):
    tree = html.fromstring(text)
    rows = tree.xpath('//tr')
    workbook = get_result_file('results.xlsx')

    worksheet = workbook.active
    worksheet.title = "Coh Results"

    add_excel_header(rows, worksheet)



    for row in rows:
        cols = row.getchildren()

        # if len(cols) == 5 :
            # print ""+cols[0].text+", "+ cols[1].text +", "+cols[2].text +", "+ cols[3].text+", "+ cols[4].text+" "
                # columns.append(col.text)
            # data.append(columns)

            # print data.dict


    workbook.save('results.xlsx')

#Adds the excel column header
def add_excel_header(rows, worksheet):
    columnNumber = 1
    if worksheet['A1'].value != 'ID':
        worksheet.cell(row=1, column=columnNumber).value = 'ID'

        for row in rows:
            cols = row.getchildren()
            if len(cols) == 5:
                columnNumber += 1
                worksheet.cell(row=1, column=columnNumber).value = cols[1].text


# Returns a workbook either by creating a new one or by reading the previous results
def get_result_file(fileName) :
    if os.path.isfile(fileName) :
        return load_workbook(fileName)

    return Workbook()

# resultsFromSearch = get_data()
file = open('temp2.html', 'r')
parse_data(file.read())
