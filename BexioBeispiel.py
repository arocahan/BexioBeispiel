from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
import sys

#Variables to filter excel table #Variable to look for the Pharmacy in bexio
zeitDau = "2021-11"

#Variable to look for the Pharmacy in bexio
apoName = 'Apotheke Zur Rose Schoenbuehl Shoppyland'


#Method to store name of the pharmacy in Excel and Id of the Pharmacy in Bexio
def switch(value):
    return{
        'Schwanen Apotheke Baden AG': 23,
        'Benu Apotheke Langnau am Albis': 18,
        'Galexis Apotheke Niederbipp': 8,
        'Galenicare Management AG'	: 27,
        'BENU Apotheke Goldach'	: 22, 
        'Apotheke Zur Rose Schoenbuehl Shoppyland': 31,
        'Apotheke Zur Rose Buchs Wynecenter': 36,
        'MBZR Apotheken AG (Buchs)': 21,
        'MBZR Apotheken AG (Limmatplatz)': 24,
    }.get(value, 0)
    

#Open Bexio window in Chrome 
driver = webdriver.Chrome(r"chromedriver")
driver.get("https://idp.bexio.com/login")
driver.maximize_window()

#Login in Bexio
inputEmail = driver.find_element_by_id('j_username')
inputEmail.send_keys('info@pharmycare.ch')

inputPassword = driver.find_element_by_id('j_password')
inputPassword.send_keys('_________') #to be filled out

button = driver.find_element_by_class_name('button')
button.click()


#Open the window of the pharmacy in bexio
kontakID = switch(apoName)
driver.get("https://office.bexio.com/index.php/kontakt/show/id/" + str(kontakID) + "#invoices")


#Create a new bill
time.sleep(2)
button1 = driver.find_element_by_link_text('Neue Rechnung')
button1.click()
time.sleep(2)
button2 = driver.find_element_by_xpath('//*[@id="editKbItemForm"]/div/div/div[6]/button')
button2.click()
time.sleep(2)
button3 = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
button3.click()


#find the excel document on desktop with all our data to fill out the fields in Bexio
df = openpyxl.load_workbook(r'2021_Pharmy_plan.xlsx', data_only=True)
worksheet = df['MJJA']


#---------------------------------- Start of example to fill out the working hours in a month -------------------------------------------

text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
time.sleep(2)
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Stundenhonorar')
text.send_keys(Keys.ENTER)

totalStunde = 0.0

#Filter data from excel
for i in range(1, 100):
    datum = worksheet['A' + str(i)].value
    datumStr = str(datum)
    if zeitDau in datumStr:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        von = worksheet['D' + str(i)].value
        bis = worksheet['E' + str(i)].value
        stunden = worksheet['F' + str(i)].value
        if apo == apoName: #create a message in bexio
            text.send_keys(str(datum)[0:10] + ': ' + str(von)[0:5] + ' - ' + str(bis)[0:5] + ' Uhr (' + str(round(stunden, 2)) + 'h)')
            text.send_keys(Keys.ENTER)
            #calculate the total of hours 
            totalStunde += stunden

#copy results to bexio 
time.sleep(2)
button4 = driver.find_element_by_id('kb_position_custom_amount')
button4.send_keys(str(round(totalStunde, 2)))
time.sleep(2)
button5 = driver.find_element_by_id('kb_position_custom_unit_price')
#preis of the hours (120CHF)
button5.send_keys(120)

#Save the module in bexio
time.sleep(2)
buttonx = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttonx.click()

#---------------------------------- End of example to fill out the working hours in a month -------------------------------------------

"""
At this point of the code I just repeat the same process to calculate the rest of the costs of a working day at the Pharmacy.
Sorry about the mix of languages and the naming of the variables
"""


if (driver.close()):
    sys.exit()

