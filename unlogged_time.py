#!/usr/bin/python

"""
Jira Worklog Monthly Report
===========================

This script:
+ Retrieve all the tickets in WPPOPENCOM
+ Look at their histories to see who all have worked on a ticket
+ Eliminate history entries that don't require time to be added (like adding a child to the ticket)
- Fetch the time logged by each user for each ticket. 
- If someone has worked on a ticket but not logged time against it, then list their name and ticket. 
- Generate a report for the same indicating when they worked on a ticket and which field was changed by them. 

Tech changes TBD:
- Make the configuration an external file that would avoid unnecessary entries. 

Potential known issue:
- Why am I not seeing any "ignored field" log for "Epic Child" field? 

Variation:
- Get the project name (can be "ALL) and username as inputs.
- Generate a report for that person alone.
- Host this somewhere so people can see their report. 

Requirements
------------
pip install requests pandas openpyxl

Jira Permissions
----------------
Your Jira account/token must have permission to:
- Browse projects
- View worklogs

Supports:
- Jira Cloud
- Jira Data Center / Server

For Jira Cloud:
    EMAIL = your Jira email
    API_TOKEN = API token from Atlassian

For Jira Server/Data Center:
    USERNAME = username
    API_TOKEN = password or personal access token
"""

import calendar
import json
import pprint
from datetime import datetime
from collections import defaultdict
import logging

import pandas as pd
import requests
import os
from requests.auth import HTTPBasicAuth
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type


# =========================================================
# CONFIGURATION
# =========================================================

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL",  "https://jira.uhub.biz")

EMAIL = os.getenv("JIRA_EMAIL", "karthick.sundararajan@example.com")
API_TOKEN = os.getenv("JIRA_API_TOKEN", "")

# Users to include in report
USERS = [
    "alejandro.godoy@wpp.com",
    "anna.mcgrath@wpp.com",
    "otemu.akpobaro@wpp.com",
    "noel.carcases@wpp.com",
    "alex.ladutska@wpp.com",
    "sejir.ali@wpp.com",
    "adrian.agundez@wpp.com", 
    "denis.cabral@wpp.com",
    "slava.yakovenko@wpp.com", 
    "oleh.babych@wpp.com", 
    "alex.karpukhin@wpp.com",
    "Mohammad.Danish@wpp.com",
    "dmytro.sharnin@wpp.com",
    "Prashanth.S@wpp.com",
    "Leeza.Chadha@wpp.com",
    "alex.khaliavinskyi@wpp.com",
    "alina.kapshyk@wpp.com",
    "mehmet.aydin@wpp.com",
    "Dmitrii.Rogozin@wpp.com",
    "Andrii.Lukanin@wpp.com",
    "Hande.Hubar@wpp.com",
    "Daniel.Christopher@wpp.com",
    "abderrahmane.adouielouadrhiri@wpp.com",
    "kunal.talyan@wpp.com",
    "sainikhil.abburi@wpp.com",
    "inigo.angulo@wpp.com",
    "ruta.remutyte@wpp.com",
    "Karthick.Sundararajan@wpp.com",
    "amanda.els@wpp.com",
    "daniele.mula@wpp.com",
    "gianpiero.pavesi@wpp.com",
    "lucy.goudie@wpp.com",
    "saradha.sethuraman@wpp.com",
    "pratyush.singh@wpp.com",
    "phil.robertson@wpp.com"
]

USERS_LOWER = {user.lower() for user in USERS}

# Month to analyze
# YEAR = 2026
# MONTH = 4

# Output files
CACHE_FILE = f"output/jira_history_cache.json"
EXCEL_OUTPUT = f"output/jira_worklog_report.xlsx"#_{YEAR}_{MONTH:02d}.xlsx"
JSON_OUTPUT = f"output/jira_worklog_report.json"#_{YEAR}_{MONTH:02d}.json"

# Number of issues retrieved per page
MAX_RESULTS = 100

# =========================================================
# AUTHENTICATION
# =========================================================

auth = HTTPBasicAuth(EMAIL, API_TOKEN)

headers = {
    "Accept": "application/json",
    "Authorization": f"Bearer {API_TOKEN}"
}

# =========================================================
# DATE HELPERS
# =========================================================

# start_date = datetime(YEAR, MONTH, 1)
# last_day = calendar.monthrange(YEAR, MONTH)[1]
# end_date = datetime(YEAR, MONTH, last_day, 23, 59, 59)
# day_columns = [f"{day:02d}" for day in range(1, last_day + 1)]


def save_history_cache(history_info):
    """Saves the flat history_info list to a JSON file."""
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(history_info, f, indent=4)
        logging.info(f"Successfully cached {len(history_info)} history items to disk.")
    except Exception as e:
        logging.error(f"Failed to write cache file: {e}")

def load_history_cache():
    """Loads the history_info list from disk if it exists."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            logging.info(f"Loaded {len(data)} history items from local cache file.")
            return data
        except Exception as e:
            logging.error(f"Cache file found but corrupted. Deleting. Error: {e}")
            return None
    return None

# =========================================================
# SEARCH ISSUES
# =========================================================

def search_issues():
    """
    Retrieves all issues that contain worklogs
    in the target month.
    """
    start_at = 0
    all_issues = []

    # PROJECTS = ["WPPOPENPMO", "WPPOPENCOM"]
    PROJECTS = ["WPPOPENCOM"]

    project_filter = ",".join(PROJECTS)

    jql = (
        f'project IN ({project_filter}) '
        # f'AND worklogDate >= "{YEAR}-{MONTH:02d}-01" '
        # f'AND worklogDate <= "{YEAR}-{MONTH:02d}-{last_day:02d}"'
    )
    
    while True:
        url = f"{JIRA_BASE_URL}/rest/api/2/search"

        params = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": MAX_RESULTS,
            "fields": "summary"
        }

        response = requests.get(
            url,
            # auth=auth,
            headers=headers,
            params=params
        )

        response.raise_for_status()
        data = response.json()
        issues = data.get("issues", [])
        if not issues:
            print("No issues found")
            break

        all_issues.extend(issues)
        start_at += MAX_RESULTS
        if start_at >= data.get("total", 0):
            break
    logging.debug(f"Found {len(all_issues)}")
    return all_issues

def get_histories(issue_key):
    """
    Gets the details of all the actors on a ticket based on the change log
    """
    logging.debug(f"Starting get_histories for {issue_key}")
    url = f"{JIRA_BASE_URL}/rest/api/2/issue/{issue_key}?expand=changelog"

    response = requests.get(
        url,
        headers=headers
    )

    response.raise_for_status()
    histories = response.json().get("changelog").get("histories", []) or []
    logging.debug(f"Ending get_histories for {issue_key}. Found {len(histories)} history entries.")
    return histories


# =========================================================
# GET WORKLOGS
# =========================================================
@retry(
    stop=stop_after_attempt(4),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((requests.exceptions.Timeout, requests.exceptions.RequestException)),
    reraise=True  # Reraises the error if all 4 attempts fail so the caller function can catch it
)
def get_issue_worklogs(issue_key):
    """
    Retrieves all worklogs for an issue.
    """
    url = f"{JIRA_BASE_URL}/rest/api/2/issue/{issue_key}/worklog"

    response = requests.get(
        url,
        # auth=auth,
        headers=headers, 
        timeout = 15
    )
    response.raise_for_status()

    return response.json().get("worklogs", [])

def get_history_info(history_info, issue_id):
    histories = get_histories(issue_id)
    logging.debug(f"Starting processing of ticket history for {issue_id} which has {len(histories)} entries.") #, end="")

    for history in histories:
        user = (
                history.get("author", {}).get("emailAddress")
            )
        logging.info(f"Issue: {issue_id}\tAuthor: {user}")

        if user.lower() not in USERS_LOWER:
            logging.debug(f"{user} not in USERS")
            continue

        ignorable_fields = {"Epic Child", "timespent"}
        items = history.get("items")
        for item in items:
            action = item.get("field")
            if action in ignorable_fields:
                continue
            history_info.append({
                "user": user.lower(),
                "issue_id": issue_id,
                "timestamp": history.get("created"),
                "action": str(action)
            })
    logging.debug(f"Ending processing for ticket history for {issue_id}")

def get_time_logged(time_logged, issue_id):
    logging.debug(f"Starting get_time_logged")

    try:
        worklogs = get_issue_worklogs(issue_id)
    except Exception as e:
        logging.error(f"Failed to fetch worklogs for {issue_id} after retries. Skipping. Error: {e}")
        return

    for worklog in worklogs:
        author = (
                worklog.get("author", {})
                .get("emailAddress")
            )
        if author.lower() not in USERS_LOWER:
            continue
        seconds = worklog.get("timeSpentSeconds", 0)
        hours = round(seconds / 3600, 2)
        time_logged[(author, issue_id)] += hours

    logging.debug(f"Ending get_time_logged")

def get_unlogged_actions(unlogged_actions, history_info, time_logged):
    logging.debug(f"Starting get_unlogged_actions")
    for record in history_info:
        user = record["user"]
        issue = record["issue_id"]

        if time_logged.get((user, issue), 0) == 0:
            unlogged_actions.append(record)
    logging.debug(f"Ending get_unlogged_actions")

def generate_report_by_user(report_by_user, unlogged_actions):
    logging.debug(f"Starting generate_report_by_user")
    for record in unlogged_actions:
        report_by_user[record["user"]].append(record)
    logging.debug(f"Ending generate_report_by_user")

def export_unlogged_to_excel(unlogged_actions, filename="unlogged_time_report.xlsx"):
    logging.info("Starting Excel export process...")
    
    if not unlogged_actions:
        logging.warning("No unlogged actions found. Excel file will not be created.")
        return

    # 1. Create a DataFrame directly from your list of dictionaries
    df = pd.DataFrame(unlogged_actions)
    
    # 2. Keep ONLY the 'user' and 'issue_id' columns, dropping everything else
    df = df[['user', 'issue_id']]
    
    # 3. Optional: Remove identical duplicate rows 
    # (e.g., if a user updated 3 fields on 1 ticket, they only need 1 line in this report)
    df = df.drop_duplicates()
    
    # 4. Rename columns so they look clean and professional in Excel
    df = df.rename(columns={
        "user": "User Email",
        "issue_id": "Jira Issue ID"
    })
    
    # 5. Save the file without the index column
    df.to_excel(filename, index=False)
    logging.info(f"Successfully created Excel report: {filename}")

# =========================================================
# CONVERT TO DATAFRAME
# =========================================================
def report_to_dataframe(report):
    rows = []

    for (user, issue_key), values in sorted(report.items()):
        row = {
            "User": user,
            "Issue": issue_key
        }

        total = 0
        for day in day_columns:
            hours = round(values.get(day, 0), 2)
            row[day] = hours
            total += hours

        row["Total"] = round(total, 2)
        rows.append(row)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df[
            ["User", "Issue"] +
            day_columns +
            ["Total"]
        ]

    return df


# =========================================================
# EXPORT FUNCTIONS
# =========================================================

def export_excel(df):

    with pd.ExcelWriter(
        EXCEL_OUTPUT,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Worklog Report",
            index=False
        )

    print(f"\nExcel exported: {EXCEL_OUTPUT}")


def export_json(df):

    records = df.to_dict(orient="records")

    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=4)

    print(f"JSON exported: {JSON_OUTPUT}")


def export_to_excel(unlogged_actions):
    # 1. Convert your flat list of dictionaries directly into a DataFrame
    df = pd.DataFrame(unlogged_actions)
    
    # 2. Rename columns to look professional in Excel (Optional)
    df = df.rename(columns={
        "user": "User ID",
        "issue_id": "Issue Key",
        "timestamp": "Date/Time Worked",
        "action": "Field Changed"
    })
    
    # 3. Save directly to an Excel file
    df.to_excel(EXCEL_OUTPUT, index=False)
    print(f"[SUCCESS] Report saved to {EXCEL_OUTPUT}")

# =========================================================
# BUILD REPORT
# =========================================================

def build_report():
    """
    Creates:
    {
        (user, issue_key): {
            "01": hours,
            "02": hours,
            ...
        }
    }
    """
    logging.debug("Starting build_report")
    history_info = []
    time_logged = defaultdict(int)
    unlogged_actions = []

    history_info = load_history_cache()
    if history_info is None:
        logging.error("No cache found. Please run ticket_history_details.py to get the data.")
        print(f"No cache found. Please run ticket_history_details.py to get the data.")
        exit(-1)

    unique_issues = {record["issue_id"] for record in history_info if "issue_id" in record}
    sorted_issues = sorted(unique_issues)
    issue_cnt = len(sorted_issues)

    logging.info(f"\nFound {issue_cnt} unique issues with history information.\n")

    for index, issue_id in enumerate(sorted_issues, 1):
        logging.debug(f"[{index}/{issue_cnt}] Checking worklogs for {issue_id}...")
        get_time_logged(time_logged, issue_id)
    
    get_unlogged_actions(unlogged_actions, history_info, time_logged)

    export_unlogged_to_excel(unlogged_actions)

# =========================================================
# MAIN
# =========================================================
def main():
    logging.basicConfig(level=logging.DEBUG, format="%(levelname)s: %(message)s")
    logging.debug("Starting run...")

    build_report()

    # df = report_to_dataframe(report)

    # print("\n================ REPORT ================\n")

    # if df.empty:
    #     print("No worklogs found.")
    # else:
    #     print(df.to_string(index=False))

    #     export_excel(df)
    #     export_json(df)


if __name__ == "__main__":
    main()
