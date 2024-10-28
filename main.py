from __future__ import print_function
import os.path
import pickle

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


import time

from selenium import webdriver

from selenium.webdriver.common.by import By

# опции
options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                     "(KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
                     )

url = "https://intra.lantaservice.com/slareport/index"

driver = webdriver.Chrome(executable_path="C:/Users/klial/PycharmProjects/pythonProject/chromedriver.exe",
                         options=options)

ch = 0
while ch<10:
    driver.get(url=url)
    login_input = driver.find_element(By.ID, "login")
    login_input.send_keys("*****")

    pss_input = driver.find_element(By.ID, "password")
    pss_input.send_keys("*****")
    cl = driver.find_element(By.XPATH, "//*[@id='loginholder']/form/div[4]/button")
    cl.click()
    time.sleep(2)

    try:
        class GoogleSheet:
            SPREADSHEET_ID = '19Iug2QxGRmgt36SU0rLPmhYjOyFOn15B4t50JtXR9Ts'
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            service = None

            def __init__(self):
                creds = None
                if os.path.exists('token.pickle'):
                    with open('token.pickle', 'rb') as token:
                        creds = pickle.load(token)

                if not creds or not creds.valid:
                    if creds and creds.expired and creds.refresh_token:
                        creds.refresh(Request())
                    else:
                        print('flow')
                        flow = InstalledAppFlow.from_client_secrets_file(
                            'credentials.json', self.SCOPES)
                        creds = flow.run_local_server(port=0)
                    with open('token.pickle', 'wb') as token:
                        pickle.dump(creds, token)

                self.service = build('sheets', 'v4', credentials=creds)

            def updateRangeValues(self, range, values):
                data = [{
                    'range': range,
                    'values': values
                }]
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': data
                }
                result = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                                          body=body).execute()
                print('{0} cells updated.'.format(result.get('totalUpdatedCells')))


        gs = GoogleSheet()
        time.sleep(10)





        alf1 = 'C'
        alf2 = 'J'
        alf3 = 'Q'
        BD1 = ["//*[@id='savedfilter9723']/a", "//*[@id='savedfilter9715']/a",
               "//*[@id='savedfilter9797']/a",
               "//*[@id='savedfilter9766']/a", "//*[@id='savedfilter9726']/a", "//*[@id='savedfilter9763']/a",
               "//*[@id='savedfilter9724']/a",
               "//*[@id='savedfilter9716']/a", "//*[@id='savedfilter9759']/a", "//*[@id='savedfilter9760']/a",
               "//*[@id='savedfilter9722']/a",
               "//*[@id='savedfilter9767']/a", "//*[@id='savedfilter9762']/a", "//*[@id='savedfilter9720']/a",
               "//*[@id='savedfilter9761']/a",
               "//*[@id='savedfilter9794']/a", "//*[@id='savedfilter9796']/a", "//*[@id='savedfilter9795']/a",
               "//*[@id='savedfilter9718']/a",
               "//*[@id='savedfilter9798']/a", "//*[@id='savedfilter9717']/a", "//*[@id='savedfilter9757']/a",
               "//*[@id='savedfilter9764']/a",
               "//*[@id='savedfilter9727']/a", "//*[@id='savedfilter9758']/a", "//*[@id='savedfilter9765']/a"]

        BD2 = ["//*[@id='savedfilter9774']/a", "//*[@id='savedfilter9768']/a", "//*[@id='savedfilter9792']/a",
               "//*[@id='savedfilter9787']/a",
               "//*[@id='savedfilter9776']/a", "//*[@id='savedfilter9784']/a", "//*[@id='savedfilter9775']/a",
               "//*[@id='savedfilter9769']/a",
               "//*[@id='savedfilter9780']/a", "//*[@id='savedfilter9781']/a", "//*[@id='savedfilter9773']/a",
               "//*[@id='savedfilter9788']/a",
               "//*[@id='savedfilter9783']/a", "//*[@id='savedfilter9772']/a", "//*[@id='savedfilter9782']/a",
               "//*[@id='savedfilter9789']/a",
               "//*[@id='savedfilter9791']/a", "//*[@id='savedfilter9790']/a", "//*[@id='savedfilter9771']/a",
               "//*[@id='savedfilter9793']/a",
               "//*[@id='savedfilter9770']/a", "//*[@id='savedfilter9778']/a", "//*[@id='savedfilter9785']/a",
               "//*[@id='savedfilter9777']/a",
               "//*[@id='savedfilter9779']/a", "//*[@id='savedfilter9786']/a"]

        BD3 = ["//*[@id='savedfilter9735']/a", "//*[@id='savedfilter9728']/a", "//*[@id='savedfilter9755']/a",
               "//*[@id='savedfilter9750']/a",
               "//*[@id='savedfilter9739']/a", "//*[@id='savedfilter9747']/a", "//*[@id='savedfilter9737']/a",
               "//*[@id='savedfilter9729']/a",
               "//*[@id='savedfilter9743']/a", "//*[@id='savedfilter9744']/a", "//*[@id='savedfilter9734']/a",
               "//*[@id='savedfilter9751']/a",
               "//*[@id='savedfilter9746']/a", "//*[@id='savedfilter9733']/a", "//*[@id='savedfilter9745']/a",
               "//*[@id='savedfilter9752']/a",
               "//*[@id='savedfilter9754']/a", "//*[@id='savedfilter9753']/a", "//*[@id='savedfilter9731']/a",
               "//*[@id='savedfilter9756']/a",
               "//*[@id='savedfilter9730']/a", "//*[@id='savedfilter9741']/a", "//*[@id='savedfilter9748']/a",
               "//*[@id='savedfilter9740']/a",
               "//*[@id='savedfilter9742']/a", "//*[@id='savedfilter9749']/a"]

        alf = ["C", "D", "E", "F", "G", "J", "K", "L", "M", "N", "R", "S", "T", "U"]
        stl = 4
        for i in range(0, 26):
                if i == 11 or i == 12 or i == 13:

                    driver.find_element(By.XPATH, BD1[i]).click()
                    time.sleep(70)
                    text1 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text
                    test_range = str('Подгрузка-отчет!' + alf1 + str(2 + i * 9) + ':' + alf1 + str(6 + i * 9))
                    test_values = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]

                    gs.updateRangeValues(test_range, test_values)





                else:
                    driver.find_element(By.XPATH, BD1[i]).click()
                    time.sleep(10)
                    text1 = driver.find_element(By.XPATH,"//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text
                    test_range = str('Подгрузка-отчет!' + alf1 + str(2 + i * 9) + ':' + alf1 + str(6 + i * 9))

                    test_values = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]

                    gs.updateRangeValues(test_range, test_values)



        for i in range(0, 26):
                if i == 11 or i == 12 or i == 13:

                    driver.find_element(By.XPATH, BD2[i]).click()
                    time.sleep(70)
                    text1 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text

                    test_range1 = str('Подгрузка-отчет!' + alf2 + str(2 + i * 9) + ':' + alf2 + str(6 + i * 9))
                    test_values1 = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]

                    gs.updateRangeValues(test_range1, test_values1)





                else:
                    driver.find_element(By.XPATH, BD2[i]).click()
                    time.sleep(10)
                    text1 = driver.find_element(By.XPATH,"//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text

                    test_range1 = str('Подгрузка-отчет!' + alf2 + str(2 + i * 9) + ':' + alf2 + str(6 + i * 9))

                    test_values1 = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]


                    gs.updateRangeValues(test_range1, test_values1)


        for i in range(0, 26):
                if i == 11 or i == 12 or i == 13:

                    driver.find_element(By.XPATH, BD3[i]).click()
                    time.sleep(70)
                    text1 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text

                    test_range2 = str('Подгрузка-отчет!' + alf3 + str(2 + i * 9) + ':' + alf3 + str(6 + i * 9))

                    test_values2 = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]

                    gs.updateRangeValues(test_range2, test_values2)




                else:
                    driver.find_element(By.XPATH, BD3[i]).click()
                    time.sleep(10)
                    text1 = driver.find_element(By.XPATH,"//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[1]/td[" + str(stl) + "]").text
                    text2 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[2]/td[" + str(
                                                    stl) + "]").text
                    text3 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[3]/td[" + str(
                                                    stl) + "]").text
                    text4 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[4]/td[" + str(
                                                    stl) + "]").text

                    text6 = driver.find_element(By.XPATH,
                                                "//*[@id='tableholder']/div[2]/div/div[2]/table/tbody/tr[6]/td[" + str(
                                                    stl) + "]").text

                    test_range2 = str('Подгрузка-отчет!' + alf3 + str(2 + i * 9) + ':' + alf3 + str(6 + i * 9))

                    time.sleep(2)
                    test_values2 = [
                        [text1],
                        [text2],
                        [text3],
                        [text4],
                        [text6],

                    ]

                    gs.updateRangeValues(test_range2, test_values2)
    finally:
        ch += 1
        print('Обновленно')
        time.sleep(10)

driver.close()
driver.quit()