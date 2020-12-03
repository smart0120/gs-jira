#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Template for python3 terminal scripts.
This gist allows you to quickly create a functioning
python3 terminal script using argparse and subprocess.
"""

import os, gspread
from dotenv import load_dotenv
from jira import JIRA

__author__ = "bursno22"
__license__ = "MIT"
__version__ = "0.0.1"

load_dotenv()

def main():
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))
    print (sh.sheet1.get('G8'))

    auth_jira = JIRA(options={'server': os.getenv('JIRA_SERVER_URL')}, 
        basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN')))
    issue = auth_jira.issue('TIC-1')
    print (issue.fields.summary)

if __name__ == '__main__':
    main()