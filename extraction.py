'''
    CSHP PSN Extraction Tool

    A tool designed to archive the CSHP PSN before sunsetting QID
'''
class Forum:
    def __init__(self, title, url):
        self.title = title
        self.url = url

class Thread:
    def __init__(self, title, date, url):
        self.title = title
        self.date = date
        self.url = url

class Attachment:
    def __init__(self, title, url):
        self.title = title
        self.url = url

def get_rows(url):
    page = s.get(url)
    soup = BeautifulSoup(page.text, "lxml")
    table = soup.find("table", class_="efMainTable")
    rows = table.find_all("tr")

    return rows

def extract_threads(rows, list):
    for row in rows:
        link = row.find("a", class_="ftlevel2", href=True)

        try:
            if "thread.cfm" in link["href"]:
                title = link.string
                dateString = row.find_all("td")[-1].find_all(text=True, recursive=False)[1].strip()
                date = datetime.strptime(dateString, "%b %d, %Y %H:%M")
                url = "http://psn.cshp.ca/%s" % link["href"]
                list.append(Thread(title, date, url))
        except Exception as e:
            None
            # print (e)

    return list
    
def sanitize_names(name):
    name = name.replace("/", " ")
    name = name.replace("<", " ")
    name = name.replace(">", " ")
    name = name.replace(":", "")
    name = name.replace('"', "'")
    name = name.replace("\\", " ")
    name = name.replace("|", " ")
    name = name.replace("?", "")
    name = name.replace("*", "")

    return name


import sys
from unipath import Path
import configparser
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time
from datetime import datetime
import re

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

# Find all the proper link addresses to each forum
rows = get_rows("http://psn.cshp.ca/")

forumList = []

for row in rows:
    link = row.find("a", class_="ftlevel2", href=True)
    
    try:
        if "list.cfm" in link["href"]:
            title = link.string
            url = "http://psn.cshp.ca/%s" % link["href"]
            forumList.append(Forum(title, url))
    except:
        None

# Cycle through each forum and save each thread
for forum in forumList:
    print ("\nAccessing %s Forum" % forum.title)
    print ("----------------------------------------------------------------------")

    # Get the table rows for this page
    rows = get_rows(forum.url)

    # Extract the first page threads
    threadList = []
    threadList = extract_threads(rows, threadList)
    
    # Loop through any subsequent pages and collect the threads
    nextLink = rows[-1].find("a", string="Next >>")

    if nextLink:
        moreThreads = True
        nextURL = "http://psn.cshp.ca/%s" % nextLink["href"]
    else:
        moreThreads = False

    while moreThreads:
        time.sleep(1)

        rows = get_rows(nextURL)
        
        threadList = extract_threads(rows, threadList)

        # Check for an active "Next >>" link to see if this continues
        nextLink = rows[-1].find("a", string="Next >>")

        if nextLink:
            moreThreads = True
            nextURL = "http://psn.cshp.ca/%s" % nextLink["href"]
        else:
            moreThreads = False
    
    # Create a folder to hold thread content
    fForum = root.child("PSN").child(forum.title)
    Path.mkdir(fForum)

    # Access each thread and download html + any relevant attachemnts
    for thread in threadList:
        time.sleep(1)
        print ("Accessing thread: %s" % thread.title)

        # Create folder to hold thread (and escape unsafe path characters)
        fName = "%s - %s" % (thread.date.strftime("%Y-%m-%d"), thread.title)

        fName = sanitize_names(fName)
        
        # Truncate file name to be a max of 100 charcters
        fName = fName[:100].strip()

        # Create the final file path and folder
        fThread = fForum.child(fName)
        Path.mkdir(fThread)

        # Access the thread URL
        page = s.get(thread.url)
        soup = BeautifulSoup(page.text, "lxml")
        table = soup.find("table", class_="efMainTable")

        # Save the table contents as an HTML file
        # Create a html file name to save the thread
        fHTML = fThread.child("thread.html")

        with open(fHTML.absolute(), "w", encoding="utf-8", errors="replace") as file:
            file.write(table.prettify())

        # Collect a list of all the attachments on the page
        attachments = table.find_all("a", href=True)

        attachmentList = []

        for attachment in attachments:
            try:
                if "getAttachment.cfm" in attachment["href"]:
                    title = attachment.string
                    url = "http://psn.cshp.ca/%s" % attachment["href"]
                    attachmentList.append(Attachment(title, url))
            except Exception as e:
                None
                #print (e)

        # Cycle through each attachment and download it
        attachmentMatching = []
        i = 1

        for attachment in attachmentList:
            onlineFile = s.get(attachment.url)

            if len(attachment.title) > 75:
                name = sanitize_names(Path(attachment.title).stem)
                name = name[:75].strip()

                extension = Path(attachment.title).ext
                fileName = "%s%s" % (name, extension)
            else:
                fileName = attachment.title

            # Number the attachments to prevent duplicates
            fileName = "%02d - %s" % (i, fileName)
            i = i + 1

            # Collect the data on the original and update file names for matching later
            attachmentMatching.append([attachment.title, fileName])

            # Create the final file name
            fileName = fThread.child(fileName)

            # Save the file to disk
            with open(fileName, "wb") as saveFile:
                saveFile.write(onlineFile.content)

        if attachmentMatching:
            with open(fThread.child("attachments.txt"), "w") as file:
                for item in attachmentMatching:
                    file.write("%s    |    %s" % (item[0], item[1]))

