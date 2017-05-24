'''
    CSHP PSN Extraction Tool

    A tool designed to archive the CSHP PSN before sunsetting QID
'''

import sys
from unipath import Path
import mechanize
import configparser

# APPLICATION SETUP
# Set up root path to generate absolute paths to files
root = Path(sys.argv[1])

config = configparser.ConfigParser()
config.read(root.parent.child("config").child("cshp_config.cfg").absolute())
username = config.get("cshp", "user")
password = config.get("cshp", "password")

br = mechanize.Browser()
br.open("http://www.cshp.ca/login_e.asp?targetURL=http%3A%2F%2Fpsn%2Ecshp%2Eca")
print (br.title())