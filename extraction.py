'''
    CSHP PSN Extraction Tool

    A tool designed to archive the CSHP PSN before sunsetting eForums
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
        self.loc = "http://psn.cshp.ca/%s" % url
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
    name = name.replace(".", " ")
    name = name.replace("\t", " ")

    return name

def sanitize_extension(ext):
    # Remove any characters that are not letters
    # (no known extensions using numbers)
    regex = re.compile("[^a-zA-Z]")
    ext = regex.sub("", ext)
    
    # Add back leading period
    ext = ".%s" % ext
    
    return ext

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
driver.get("http://www.cshp.ca/user/login")

# Login to CSHP
tUsername = driver.find_element_by_id("edit-name")
tUsername.clear()
tUsername.send_keys(username)

tPassword = driver.find_element_by_id("edit-pass")
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
        fName = "%s - %s" % (thread.date.strftime("%Y-%m-%d %H:%M"), thread.title)

        fName = sanitize_names(fName)
        
        # Truncate file name to be a max of 90 charcters
        fName = fName[:70].strip()

        # Create the final file path and folder
        fThread = fForum.child(fName)
        Path.mkdir(fThread)

        try:
            # Access the thread URL
            page = s.get(thread.url)
            soup = BeautifulSoup(page.text, "lxml")
            table = soup.find("table", class_="efMainTable")
        except Exception as e:
            print ("Error isolating data table: %s" % e)

        # Save the table contents as an HTML file
        # Create a html file name to save the thread
        fHTML = fThread.child("thread.html")

        try:
            with open(fHTML.absolute(), "w", encoding="utf-8", errors="replace") as file:
                file.write(table.prettify())
        except Exception as e:
            print ("Error saving HTML: %s" % e)

        # Collect a list of all the attachments on the page
        try:
            attachments = table.find_all("a", href=True)
        except Exception as e:
            print ("Error finding attachments: %s" % e)

        attachmentList = []

        for attachment in attachments:
            try:
                if "getAttachment.cfm" in attachment["href"]:
                    # Get file name
                    title = attachment.string

                    # Record the url
                    url = attachment["href"]

                    # Save attachment data to list
                    attachmentList.append(Attachment(title, url))
            except Exception as e:
                None
                #print (e)

        # Cycle through each attachment and download it
        attachmentMatching = []

        # Attachment number starts at 2 (1 is reserved for the thread)
        i = 2

        for attachment in attachmentList:
            try:
                onlineFile = s.get(attachment.loc)

                # Truncate and sanitize file names
                name = Path(attachment.title).stem
                name = sanitize_names(name)
                name = name[:70].strip()

                # Sanitize the extension
                extension = Path(attachment.title).ext
                extension = sanitize_extension(extension)
                fileName = "%s%s" % (name, extension)
                
                # Number the attachments to prevent duplicates
                fileName = "%02d - %s" % (i, fileName)
                i = i + 1

                # Collect the data on the original and update file names for matching later
                attachmentMatching.append([attachment.url, attachment.title, fileName])

                # Create the final file name
                fileName = fThread.child(fileName)

                # Save the file to disk
                with open(fileName, "wb") as saveFile:
                    saveFile.write(onlineFile.content)

            except Exception as e:
                print ("Error saving attachment: %s" % e)

        if attachmentMatching:
            try:
                with open(fThread.child("attachments.txt"), "w") as file:
                    for item in attachmentMatching:
                        file.write("%s    |    %s    |    %s\n" % (item[0], item[1], item[2]))
            except Exception as e:
                print ("Error saving attachment list: %s" % e)

print ("\n----------------------------------------------------------------------")
print ("EXTRACTION COMPLETE")
print ("----------------------------------------------------------------------\n")