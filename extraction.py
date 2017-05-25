'''
    CSHP PSN Extraction Tool

    A tool designed to archive the CSHP PSN before sunsetting QID
'''

import sys
from unipath import Path
import configparser
import requests

# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

config = configparser.ConfigParser()
config.read(root.parent.child("config").child("cshp_config.cfg").absolute())
username = config.get("cshp", "user")
password = config.get("cshp", "password")

s = requests.Session()
p = s.post("http://www.cshp.ca/login_e.asp", data={"u": "", "formAction": "login", "username": username, "password": password, "Submit": "Login"})