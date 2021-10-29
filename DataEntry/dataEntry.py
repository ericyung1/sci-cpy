from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd

df = pd.read_excel('Excel Spreadsheet.xlsx',
converters = {'First Name':str, 'Last Name':str, 'Birthdate':str, 'Email':str, 'Phone':str, 'Zip Code':int})

for index, row in df.head(n = 100).iterrows():
    web = webdriver.Chrome()
    web.get('https://www.goarmy.com/info.html?iom=ARJA')
    FirstName = row[0]
    LastName = row[1]
    Birthday = row[2]
    BirthMonth = str(int(Birthday[5:7]) + 1)
    BirthDay = str(int(Birthday[8:10]) + 1)
    BirthYear = str(2010 - int(Birthday[0:4]))
    EmailAddress = row[3]
    PhoneNumber = row[4]
    Phone = '(' + PhoneNumber[:3] + ') ' + PhoneNumber[3:6] + ' ' + PhoneNumber[6:]
    ZipCode = row[5]

    xMonth = '//*[@id="emailquestion"]/form/div[7]/div/ul/li[' + BirthMonth + ']/label'
    xDay = '//*[@id="emailquestion"]/form/div[8]/div/ul/li[' + BirthDay + ']/label'
    xYear = '//*[@id="emailquestion"]/form/div[9]/div/ul/li[' + BirthYear + ']/label'

    month = web.find_element_by_xpath('//*[@id="emailquestion"]/form/div[7]/p/span').click()
    dropdownElement = web.find_element_by_xpath(xMonth)
    action1 = ActionChains(web)
    action1.click(on_element=dropdownElement)
    action1.perform()

    day = web.find_element_by_xpath('//*[@id="emailquestion"]/form/div[8]/p/span').click()
    dropdownElement = web.find_element_by_xpath(xDay)
    action2 = ActionChains(web)
    action2.click(on_element = dropdownElement)
    action2.perform()

    year = web.find_element_by_xpath('//*[@id="emailquestion"]/form/div[9]/p/span').click()
    dropdownElement = web.find_element_by_xpath(xYear)
    action3 = ActionChains(web)
    action3.click(on_element = dropdownElement)
    action3.perform()

    first = web.find_element_by_xpath('//*[@id="firstName"]')
    first.send_keys(FirstName)

    last = web.find_element_by_xpath('//*[@id="lastName"]')
    last.send_keys(LastName)

    email = web.find_element_by_xpath('//*[@id="emailAddress"]')
    email.send_keys(EmailAddress)

    zip = web.find_element_by_xpath('//*[@id="addressZip"]')
    zip.send_keys(ZipCode)

    phone = web.find_element_by_xpath('//*[@id="phoneNumber"]')
    phone.send_keys(Phone)

    confirm = web.find_element_by_xpath('//*[@id="agreeWaiver"]')
    confirm.click()

    getmoreInfo = web.find_element_by_xpath('//*[@id="g_send_request_brc"]')
    getmoreInfo.click()

    time.sleep(2)

    web.close()