#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script for converting google sheet rows to Jira tickets
"""

import os, gspread
from dotenv import load_dotenv
from jira import JIRA
from jira.exceptions import JIRAError
from dateutil.parser import *
from dateutil.relativedelta import *
from datetime import *

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

def workdays(start, end, excluded=(6, 7)):
    """
    calcuate count of week days. (Mon to Fri)
    """
    days = 0
    while start <= end:
        if start.isoweekday() not in excluded:
            days += 1
        start += timedelta(days=1)
    return days

def generate_comment(assignee, assignee_id, due_date, cid):
    """
    comment text for issue
    """
    delta = relativedelta(parse(due_date).date(), date.today())
    reminder_text = ""
    if delta.months == 0 and delta.days > 0:
        reminder_text = ", this is a reminder that this IT Control will be due in {} working days"
    elif delta.months == 0 and delta.days == 0:
        reminder_text = ", this is a reminder that this IT Control is currently due"
    elif delta.months >= 0 and delta.days <= -18:
        reminder_text = ", this is a reminder that this IT Control is overdue and will be marked as a nonconformity in one week"
    elif delta.months < 0 or delta.days <= -25:
        reminder_text = ", this is a reminder that this IT Control is overdue and will now be marked as a nonconformity"
    if reminder_text:
        template = {
            "type": "doc",
            "version": 1,
            "content": [{
                "type": "paragraph",
                "content": [
                    {
                        "text": "Hello ",
                        "type": "text"
                    },
                    {
                        "type": "mention",
                        "attrs": {
                            "id": "",
                            "text": "",
                            "userType": "DEFAULT"
                        }
                    },
                    {
                        "text": "",
                        "type": "text"
                    }
                ]
            }]
        }
        template["content"][0]["content"][1]["attrs"]["id"] = assignee_id
        template["content"][0]["content"][1]["attrs"]["text"] = "@%s" % assignee
        template["content"][0]["content"][2]["text"] = reminder_text.format(workdays(date.today(), parse(due_date).date()))
        return template
    return ""

def main():
    # Open Google Sheet
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))

    # Open JIRA
    auth_jira = JIRA(options={'server': os.getenv('JIRA_SERVER_URL'), 'rest_api_version': 3}, 
        basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN')))

    # Loop each row of Google sheet
    row_range = [int(val) for val in os.getenv('DATA_RANGE').split(':')]

    for row in range(row_range[0], row_range[1]+1):
        # Read one row from google spread sheet
        record = sh.get_worksheet(0).row_values(row)
        cid = record[index_from_col(os.getenv('CID'))]
        due_date = record[index_from_col(os.getenv('DUE_DATE'))]
        done_date = record[index_from_col(os.getenv('DONE_DATE'))]
        jira_issue_key = record[index_from_col(os.getenv('JIRA_ISSUE_KEY'))]
        ticket_status = record[index_from_col(os.getenv('TICKET_STATUS'))]
        assignee = record[index_from_col(os.getenv('ASSIGNEE'))]

        # get JIRA User ID from 2nd sheet
        assignee_id = None
        try:
            assignee_detail = sh.get_worksheet(1).find(assignee)
            assignee_id = sh.get_worksheet(1).cell(assignee_detail.row, index_from_col(os.getenv('ASSIGNEE_ID'))+1).value
        except:
            pass

        # check if ticket status is 'Open'
        if (assignee_id and ticket_status and ticket_status.lower() == 'open'):
            try:
                issue = auth_jira.issue(jira_issue_key)
                comment = auth_jira.add_comment(jira_issue_key, generate_comment(assignee, assignee_id, due_date, cid))
                print (''.join(["#", jira_issue_key, " - Added Comment"]))
            except JIRAError as err:
                print (str(err))

if __name__ == '__main__':
    main()
