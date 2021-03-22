#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script for converting google sheet rows to Jira tickets
"""

import os, gspread, time
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

def generate_comment(assignee, assignee_id, due_date, cid, delta):
    """
    comment text for issue
    """
    reminder_text = ""
    risk_flag = False
    if delta.months == 0 and delta.days > 0:
        days = delta.days
        reminder_text = f", this is a gentle reminder that this IT Control will be due in {days} days"
    elif delta.months == 0 and delta.days == 0:
        reminder_text = "I know you have a lot on your plate right now. This is just a gentle reminder that this IT Control Review Request is due today. Process change update: if it’s not completed in a week then it will be escalated to the C-level, and if it is not completed in two weeks a risk will need to be created for the ITSC Risk Review."
    elif delta.months == 0 and delta.days in range(-13, -6):
        reminder_text = "This is just a gentle reminder that this IT Control Review Request was due a week ago. Since it has not completed, it will now be escalated to the respective C-level. If it is not completed in one more week a risk will need to be created for the ITSC Risk Review."
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
    return template

def main():
    # Open Google Sheet
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))
    primary_worksheet = sh.get_worksheet(int(os.getenv('PRIMARY_SHEET')))
    secondary_worksheet = sh.get_worksheet(int(os.getenv('SECONDARY_SHEET')))
    in_review_worksheet = sh.get_worksheet(int(os.getenv('IN_REVIEW_SHEET')))
    jira_server_url = os.getenv('JIRA_SERVER_URL')

    reviews_template = [
        {
            "type": "text",
            "text": "Hi "
        }
    ]
    review_data_list = in_review_worksheet.get_all_values()
    review_assignee_name_idx = index_from_col(os.getenv('IN_REVIEW_ASSIGNEE_NAME'))
    review_assignee_id_idx = index_from_col(os.getenv('IN_REVIEW_ASSIGNEE_ID'))
    review_assignee_status_idx = index_from_col(os.getenv('IN_REVIEW_ASSIGNEE_STATUS'))
    review_list_len = len(review_data_list)
    for row in range(1, review_list_len):
        assignee_name = review_data_list[row][review_assignee_name_idx]
        assignee_id = review_data_list[row][review_assignee_id_idx]
        assignee_status = review_data_list[row][review_assignee_status_idx]

        if assignee_status != 'enabled':
            pass

        assignee = {
            "type": "mention",
            "attrs": {
                "id": assignee_id,
                "text": "@%s" % assignee_name,
                "userType": "DEFAULT"
            }
        }
        if row != 1 and row == review_list_len - 1:
            reviews_template.append({
                "type": "text",
                "text": " and "
            })
        elif row != 1:
            reviews_template.append({
                "type": "text",
                "text": ", "
            })
        reviews_template.append(assignee)
        if row == review_list_len - 1:
            reviews_template.append({
                "type": "text",
                "text": ","
            })
    reviews_template.append({
        "type": "text",
        "text": " When you have a moment please perform the quality assurance review. Thank you "
    })

    # Open JIRA
    auth_jira = JIRA(
        options={'server': jira_server_url, 'rest_api_version': 3}, 
        basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN'))
    )

    # Loop each row of Google sheet
    row_range = [int(val) for val in os.getenv('DATA_RANGE').split(':')]
    
    for row in range(row_range[0], row_range[1]+1):
        # Google spread has limit of 100 read request per 100 seconds, 
        # So we put some 2 seconds sleep before every read to avoid quot exceed exception
        time.sleep(2)

        # Read one row from google spread sheet
        record = primary_worksheet.row_values(row)
        cid = record[index_from_col(os.getenv('CID'))]
        due_date = record[index_from_col(os.getenv('DUE_DATE'))]
        done_date = record[index_from_col(os.getenv('DONE_DATE'))]
        jira_issue_key = record[index_from_col(os.getenv('JIRA_ISSUE_KEY'))]
        ticket_status = record[index_from_col(os.getenv('TICKET_STATUS'))]
        assignee = record[index_from_col(os.getenv('ASSIGNEE'))]

        # get JIRA User ID from 2nd sheet
        assignee_id = None
        try:
            assignee_detail = secondary_worksheet.find(assignee)
            assignee_id = secondary_worksheet.cell(assignee_detail.row, index_from_col(os.getenv('ASSIGNEE_ID'))+1).value
        except:
            pass

        # check if ticket status is 'In Review'
        if (assignee_id and ticket_status and ticket_status.lower() == 'in review'):
            content = reviews_template.copy()
            content.append({
                "type": "mention",
                "attrs": {
                    "id": assignee_id,
                    "text": "@%s" % assignee,
                    "userType": "DEFAULT"
                }
            })
            comment_body = {
                "type": "doc",
                "version": 1,
                "content": [{
                    "type": "paragraph",
                    "content": content
                }]
            }
            auth_jira.add_comment(jira_issue_key, comment_body)
            print(''.join(["#", jira_issue_key, " - Added Comment"]))
        # check if ticket status is 'Open', 'In Review', 'To do', 'Open nonconformity(s) and si' or 'Open nonconformity(s)'
        elif (assignee_id and ticket_status and ticket_status.lower() in ['open', 'to do', 'open nonconformity(s) and si', 'open nonconformity(s)']):
            try:
                delta = relativedelta(parse(due_date, dayfirst=True).date(), date.today())
                if delta.months < 0 or delta.days <= -14:
                    issue_url = jira_server_url + 'browse/' + jira_issue_key
                    risk_policy_url = os.getenv('JIRA_RISK_POLICY_URL')
                    template = {
                        "type": "doc",
                        "version": 1,
                        "content": [{
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Risk Description: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": jira_issue_key,
                                    "marks": [
                                        {
                                            "type": "link",
                                            "attrs": {
                                                "href": issue_url
                                            },
                                        }
                                    ]
                                },
                                {
                                    "text": " is not yet completed. What is the risk associated with not executing this IT Control?",
                                    "type": "text"
                                },
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Probability Rating: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "Unlikely through Frequent as described in the "
                                },
                                {
                                    "type": "text",
                                    "text": "Operational Risk Management Policy",
                                    "marks": [
                                        {
                                            "type": "link",
                                            "attrs": {
                                                "href": risk_policy_url
                                            }
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Probability Description: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "Explain why the Probability Rating is appropriate."
                                }
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Severity Rating: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "None through Critical as described in the "
                                },
                                {
                                    "type": "text",
                                    "text": "Operational Risk Management Policy",
                                    "marks": [
                                        {
                                            "type": "link",
                                            "attrs": {
                                                "href": risk_policy_url
                                            }
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Severity Description: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "Explain why the Severity Rating is appropriate."
                                }
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Overall Risk Rating: ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "Low through Critical after classifying the Risk Severity and Risk Probability, the overall risk can be determined by using the Risk Classification Matrix in the "
                                },
                                {
                                    "type": "text",
                                    "text": "Operational Risk Management Policy",
                                    "marks": [
                                        {
                                            "type": "link",
                                            "attrs": {
                                                "href": risk_policy_url
                                            }
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Action Plan (or a link to an Action Plan Jira ticket): ",
                                    "marks": [
                                        {
                                            "type": "strong"
                                        }
                                    ]
                                },
                                {
                                    "type": "text",
                                    "text": "Describing 3-5 steps on how the IT Control will be partially executed until it can be fully implemented, and the timeline to do so. This will be checked on at least a monthly basis moving forward."
                                }
                            ]
                        }]
                    }
                    issue_dict = {
                        'project': os.getenv('JIRA_PROJECT_KEY'),
                        'summary': f'IT Controls {date.today().year} - Risk Log',
                        'description': template,
                        'issuetype': {'name': os.getenv('JIRA_RISK_ISSUE_TYPE')},
                        'assignee': {'name': assignee}
                    }
                    new_issue_key = str(auth_jira.create_issue(fields=issue_dict))
                    new_issue_url = jira_server_url + 'browse/' + new_issue_key
                    primary_worksheet.update_cell(row, index_from_col(os.getenv('ITSC_RISK'))+1, f'=HYPERLINK("{new_issue_url}","{new_issue_key}")')
                    primary_worksheet.update_cell(row, index_from_col(os.getenv('ITSC_RISK_STATUS'))+1, os.getenv('ITSC_RISK_STATUS_TYPE'))
                    print ('Create a new risk issue in jira - ' + new_issue_key)
                else:
                    auth_jira.issue(jira_issue_key)
                    comment_body = generate_comment(assignee, assignee_id, due_date, cid, delta)
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
