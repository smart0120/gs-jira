#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script for converting google sheet rows to Jira tickets
"""

import os, gspread
from dotenv import load_dotenv
from jira import JIRA
from jira.exceptions import JIRAError

__author__ = "bursno22"
__license__ = "MIT"
__version__ = "0.0.1"

load_dotenv()

def index_from_col(col_name):
    """
    Return index from column name
    For example, 0 for A, 1 for B, etc
    """
    return ord(col_name.upper()) - 65

def generate_comment(assignee, done_date, cid):
    """
    comment text for issue
    """
    return "Hi {}, Just a quick reminder that it's almost \
        {} to execute this IT Control Request.".format(assignee, done_date)

def main():
    # Open Google Sheet
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))

    # Open JIRA
    auth_jira = JIRA(options={'server': os.getenv('JIRA_SERVER_URL')}, 
        basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN')))

    # Loop each row of Google sheet
    row_range = [int(val) for val in os.getenv('DATA_RANGE').split(':')]

    for row in range(row_range[0], row_range[1]+1):
        # Read one row from google spread sheet
        record = sh.get_worksheet(0).row_values(row)
        cid = record[index_from_col(os.getenv('CID'))]
        done_date = record[index_from_col(os.getenv('DONE_DATE'))]
        jira_issue_key = record[index_from_col(os.getenv('JIRA_ISSUE_KEY'))]
        ticket_status = record[index_from_col(os.getenv('TICKET_STATUS'))]
        assignee = record[index_from_col(os.getenv('ASSIGNEE'))]

        # check if ticket status is 'Open'
        if (ticket_status and ticket_status.lower() == 'open'):
            try:
                issue = auth_jira.issue(jira_issue_key)
                comment = auth_jira.add_comment(jira_issue_key, generate_comment(assignee, done_date, cid))
                print (''.join(["#", jira_issue_key, " - Added Comment"]))
            except JIRAError:
                print (''.join(["#", jira_issue_key, " - ", \
                    "Issue doesn't exist or you don't have permission to use it"]))

if __name__ == '__main__':
    main()
