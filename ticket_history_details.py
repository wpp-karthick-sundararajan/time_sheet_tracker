#!/usr/bin/env python

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
API_TOKEN = os.getenv("JIRA_API_TOKEN", "") # Fill before execution

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
            # history_info[user][issue_key][field] = 0
            history_info.append({
                "user": user.lower(),
                "issue_id": issue_id,
                "timestamp": history.get("created"),
                "action": str(action)
            })
    logging.debug(f"Ending processing for ticket history for {issue_id}")

# =========================================================
# BUILD REPORT
# =========================================================

def get_history_of_all_project_issues():
    logging.debug("Starting build_report")
    history_info = []
    issues = search_issues()

    for issue in issues:
        issue_id = issue["key"]
        get_history_info(history_info, issue_id)
    save_history_cache(history_info)

# =========================================================
# MAIN
# =========================================================
def main():
    logging.basicConfig(level=logging.DEBUG, format="%(levelname)s: %(message)s")
    logging.debug("Starting run...")

    get_history_of_all_project_issues()

if __name__ == "__main__":
    main()
