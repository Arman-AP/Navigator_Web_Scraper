from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from selenium import webdriver
from bs4 import BeautifulSoup

import time
import sys
import requests
# Config:
path_to_driver = r"C:\chromedriver" #Path to Selenium Driver
navigator_username = ""
navigator_password = ""
template_path = r"C:\cv.docx" #Path to word template

# Driver Authentication and Job Finding
driver = webdriver.Chrome(executable_path=path_to_driver)

driver.get("https://navigator.wlu.ca/notLoggedIn.htm")

driver.implicitly_wait(3)
driver.find_element_by_xpath("/html/body/div[2]/div/div/div[3]/div/div/table/tbody/tr/td[1]/div/div[2]/strong/a").click()

driver.implicitly_wait(3)
driver.find_element_by_name("j_username").send_keys(navigator_username)
driver.find_element_by_name("j_password").send_keys(navigator_password)
driver.find_element_by_name("_eventId_proceed").click()
driver.get("https://navigator.wlu.ca/myAccount/co-op/postings.htm")

driver.implicitly_wait(6)
driver.find_element_by_xpath("//*[@id='quickSearchCountsContainer']/table/tbody/tr[2]/td[2]/a").click()

driver.implicitly_wait(5)

# Find by Job Title, Custom input poisiton and industry
job_title = input("job title :")
position = input("job title to post :")
industry = input("Industry :")

driver.find_element_by_link_text(job_title).click()

driver.implicitly_wait(5)
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)

driver.implicitly_wait(5)

page = driver.page_source
soup = BeautifulSoup(page, 'html.parser')
rows = soup.find_all("table", {"class": "table table-bordered"})[3]

table_data_body = rows.tbody.find_all("tr")

data = []
company = ""
salutation =""
firstname=""
lastname = ""
title = ""
address1 = ""
address2 = ""
city = ""
postalcode = ""

for tr in table_data_body:
    new_value = tr.find_all("td")
    new_value = [ele.text.strip() for ele in new_value]
    data.append(new_value)
for counter in range(len(data)):
    if("Organization" in data[counter][0]):
        company = data[counter][1]
    elif("Salutation" in data[counter][0]):
        salutation = data[counter][1]
    elif("Job Contact First Name" in data[counter][0]):
        firstname = data[counter][1]
    elif("Job Contact Last Name" in data[counter][0]):
        lastname = data[counter][1]
    elif("Contact Title" in data[counter][0]):
        title = data[counter][1]
    elif("Address Line One" in data[counter][0]):
        address1 = data[counter][1]
    elif("Address Line Two" in data[counter][0]):
        address2 = data[counter][1]
    elif("City" in data[counter][0]):
        city = data[counter][1]
    elif("Postal Code" in data[counter][0]):
        postalcode = data[counter][1]
# Error Reporting
print("company: " + company)
print("salutation " + salutation)
print("first name " + firstname)
print("last name " + lastname)
print("title " + title)
print("address" + address1)
print("address2" + address2)
print("city" + city)
print("postal" + postalcode)
print(data)
template_1 = template_path

# Merge data into a word doc
document_1 = MailMerge(template_1)

document_1.merge(
    today= date.today().strftime("%B %d, %Y"),
    company=company,    
    address=address1,
    address2 = address2,
    city=city, 
    postal=postalcode,
    salutation=salutation,
    firstname=firstname, 
    lastname = lastname,
    position=position,
    industry = industry
    )

# Save the document with the appropriate title
document_1.write(company + 'Coverletter.docx')

print("Cover Letter for:" + company)