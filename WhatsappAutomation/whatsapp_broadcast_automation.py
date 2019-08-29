from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl as excel

sheetFileName = "phone_numbers.xlsx"
broadcastMsg = "Hey"


def readPhoneNumbers(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol) - 1):
        contact = str(firstCol[cell].value)
        contact = contact.replace('+', '')
        contact = contact.replace('-', '')
        contact = contact.replace('(', '')
        contact = contact.replace(')', '')

        if contact:
            lst.append(contact)

    return lst


contactList = readPhoneNumbers(sheetFileName)

driver = webdriver.Chrome(executable_path=r'C:\PythonWorkspace\WhatsApp-bot-selenium-master\chromedriver.exe')

driver.get("https://web.whatsapp.com/")

# note this time is being used below also
wait = WebDriverWait(driver, 5)
input("Scan the QR code and then press Enter")

success = 0
sNo = 1
failList = []

for target in contactList:
    print(sNo, ". Target is: " + target)
    sNo += 1
    url = "https://api.whatsapp.com/send?phone={}&text={}&source=&data=".format(target, broadcastMsg)
    print("fetching", url)
    driver.get(url)

    try:
        # alert dialog to accept leaving the site
        driver.switch_to.alert.accept()
    except:
        print('No alert dialog')

    time.sleep(2)

    driver.find_element_by_id("action-button").click()

    try:
        # Select the Input Box
        inp_xpath = "//div[@contenteditable='true']"
        input_box = wait.until(EC.presence_of_element_located((
            By.XPATH, inp_xpath)))
        time.sleep(1)

        # Send message
        input_box.send_keys(broadcastMsg + Keys.SHIFT + Keys.ENTER + Keys.SPACE)
        # Link Preview Time, Reduce this time, if internet connection is Good
        time.sleep(5)
        input_box.send_keys(Keys.ENTER)
        print("Successfully sent to : " + target + '\n')
        success += 1
        time.sleep(0.5)

    except:
        # If target Not found Add it to the failed List
        print("Cannot find Target: " + target)
        failList.append(target)
        pass

print("\nSuccessfully Sent to: ", success)
print("Failed to Sent to: ", len(failList))
print(failList)
print('\n\n')
driver.quit()
