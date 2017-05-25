'''
    CSHP PSN Extraction Tool

    A tool designed to archive the CSHP PSN before sunsetting QID
'''

import sys
from unipath import Path
import configparser
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

config = configparser.ConfigParser()
config.read(root.parent.child("config").child("cshp_config.cfg").absolute())
username = config.get("cshp", "user")
password = config.get("cshp", "password")

# Connect to the chrome webdriver
driver = webdriver.Chrome()

# Go to login page
driver.get("http://www.cshp.ca/login_e.asp")

# Login to CSHP
tUsername = driver.find_element_by_id("username")
tUsername.clear()
tUsername.send_keys(username)

tPassword = driver.find_element_by_id("password")
tPassword.clear()
tPassword.send_keys(password)
tPassword.submit()

# Go to the eForum
driver.get("http://psn.cshp.ca/")

# Pass the session over to Requests Session
headers = {"User-Agent": "Chrome/44.0.2403.157"}
s = requests.session()
s.headers.update(headers)

for cookie in driver.get_cookies():
    c = {cookie["name"]: cookie["value"]}
    s.cookies.update(c)

driver.close()

page = s.get("http://psn.cshp.ca/")
print (page.text)

http://psn.cshp.ca/list.cfm?ListID=004804|4177B4AA4E0DF7EA6404FADBCA4CCA96