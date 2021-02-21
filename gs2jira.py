#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script for converting google sheet rows to Jira tickets
"""

import os, gspread
from dotenv import load_dotenv
from jira import JIRA
from jira.exceptions import JIRAError
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
from datetime import date, timedelta

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

def generate_comment(assignee, assignee_id, due_date, cid, risk_link, delta):
    """
    comment text for issue
    """
    reminder_text = ""
    risk_flag = False
    if delta.months == 0 and delta.days == 0:
        reminder_text = "I know you have a lot on your plate right now. This is just a gentle reminder that this IT Control Review Request is due today. Process change update: if it’s not completed in a week then it will be escalated to the C-level, and if it is not completed in two weeks a risk will need to be created for the ITSC Risk Review."
    elif delta.months == 0 and delta.days in range(-13, -6):
        reminder_text = "This is just a gentle reminder that this IT Control Review Request was due a week ago. Since it has not completed, it will now be escalated to the respective C-level. If it is not completed in one more week a risk will need to be created for the ITSC Risk Review."
    elif delta.months < 0 or delta.days <= -14:
        risk_flag = True
        reminder_text = "This is just a gentle reminder that this IT Control Review Request was due two weeks ago. Since it has not completed, it has already been escalated to the respective C-level. A risk must be created for the ITSC Risk Review. "
    else:
        return reminder_text
    template = {
        "type": "doc",
        "version": 1,
        "content": [{
            "type": "paragraph",
            "content": [
                {
                    "text": "Hi, ",
                    "type": "text"
                },
                {
                    "type": "mention",
                    "attrs": {
                        "id": "",
                        "text": "@Hi",
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
    template["content"][0]["content"][2]["text"] = reminder_text
    if risk_flag:
        template["content"][0]["content"].append({
            "type": "text",
            "text": "Link to Risk",
            "marks": [
                {
                    "type": "link",
                    "attrs": {
                        "href": risk_link,
                        "title": "Atlassian"
                    }
                }
            ]
        })
    return template

def main():
    # Open Google Sheet
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))

    # Open JIRA
    auth_jira = JIRA(
        options={'server': os.getenv('JIRA_SERVER_URL'), 'rest_api_version': 3}, 
        basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN'))
    )

    # Loop each row of Google sheet
    row_range = [int(val) for val in os.getenv('DATA_RANGE').split(':')]

    for row in range(row_range[0], row_range[1]+1):
        # Read one row from google spread sheet
        primary_sheet = int(os.getenv('PRIMARY_SHEET')) #performance report sheet
        record = sh.get_worksheet(primary_sheet).row_values(row)
        cid = record[index_from_col(os.getenv('CID'))]
        due_date = record[index_from_col(os.getenv('DUE_DATE'))]
        done_date = record[index_from_col(os.getenv('DONE_DATE'))]
        jira_issue_key = record[index_from_col(os.getenv('JIRA_ISSUE_KEY'))]
        ticket_status = record[index_from_col(os.getenv('TICKET_STATUS'))]
        assignee = record[index_from_col(os.getenv('ASSIGNEE'))]
        risk_link_id = record[index_from_col(os.getenv('ITSC_RISK'))]
        risk_link = ''.join([os.getenv('RISK_ATTR'), risk_link_id])

        # get JIRA User ID from 2nd sheet
        assignee_id = None
        try:
            secondary_sheet = int(os.getenv('SECONDARY_SHEET'))
            assignee_detail = sh.get_worksheet(secondary_sheet).find(assignee)
            assignee_id = sh.get_worksheet(secondary_sheet).cell(assignee_detail.row, index_from_col(os.getenv('ASSIGNEE_ID'))+1).value
        except:
            pass
        # check if ticket status is 'Open'
        if (assignee_id and ticket_status and ticket_status.lower() == 'open'):
            try:
                auth_jira.issue(jira_issue_key)
                delta = relativedelta(parse(due_date, dayfirst=True).date(), date.today())
                comment_body = generate_comment(assignee, assignee_id, due_date, cid, risk_link, delta)
                if comment_body:
                    auth_jira.add_comment(jira_issue_key, comment_body)
                    print(''.join(["#", jira_issue_key, " - Added Comment"]))
                else:
                    if delta.days in range(-6, 0) and delta.months == 0:
                        print('IT Control Review Request was few days ago')
                    else:
                        print('Have some time before IT Control Review Request')
            except JIRAError as err:
                print (str(err))

if __name__ == '__main__':
    main()
