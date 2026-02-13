"""Monday.com API client for fetching Vans department projects."""

import os
import json
import urllib.request
import urllib.error
from typing import List, Dict
from datetime import datetime
from .models import Project, FilmingDate, EditingRange
from ..config import *
from .. import token_manager


class MondayClient:
    """Client for interacting with Monday.com GraphQL API."""

    def __init__(self):
        """Initialise the Monday.com client."""
        # Get API token using token manager (prompts user if needed)
        self.api_token = token_manager.get_token()
        if not self.api_token:
            raise ValueError(
                "Monday.com API token is required.\n\n"
                "The app needs your API token to fetch project data.\n"
                "Get your token from: https://monday.com → Profile → Developers → My Access Tokens"
            )

        self.api_url = "https://api.monday.com/v2"

    def _make_request(self, query: str, variables: Dict = None) -> Dict:
        """
        Make a GraphQL request to Monday.com API.

        Args:
            query: GraphQL query string
            variables: Query variables

        Returns:
            Response data dictionary
        """
        headers = {
            "Content-Type": "application/json",
            "Authorization": self.api_token
        }

        data = {"query": query}
        if variables:
            data["variables"] = variables

        request = urllib.request.Request(
            self.api_url,
            data=json.dumps(data).encode('utf-8'),
            headers=headers,
            method='POST'
        )

        try:
            with urllib.request.urlopen(request) as response:
                result = json.loads(response.read().decode('utf-8'))

                if "errors" in result:
                    error_msgs = [e.get("message", str(e)) for e in result["errors"]]
                    raise Exception(f"Monday.com API errors: {', '.join(error_msgs)}")

                return result.get("data", {})
        except urllib.error.HTTPError as e:
            error_body = e.read().decode('utf-8') if e.fp else str(e)
            raise Exception(f"HTTP {e.code} error from Monday.com: {error_body}")
        except urllib.error.URLError as e:
            raise Exception(f"Network error connecting to Monday.com: {e.reason}")

    def get_vans_projects(self) -> List[Project]:
        """
        Fetch all Vans department projects from Monday.com.

        Returns:
            List of Project objects with filming dates and editing ranges.
        """
        # GraphQL query to fetch board items with all needed columns
        query = """
        query ($boardId: [ID!]) {
          boards(ids: $boardId) {
            items_page(limit: 500) {
              items {
                id
                name
                column_values {
                  id
                  text
                  value
                }
              }
            }
          }
        }
        """

        variables = {"boardId": [str(BOARD_ID)]}

        try:
            data = self._make_request(query, variables)

            if not data.get("boards") or not data["boards"]:
                return []

            items = data["boards"][0].get("items_page", {}).get("items", [])

            # Parse items and filter for Vans department
            projects = []
            for item in items:
                # Check if item belongs to Vans department
                column_values = item.get('column_values', [])
                department_col = self._find_column_value(column_values, COLUMN_DEPARTMENT)

                # Only process Vans department items (ID 17)
                if department_col and department_col.get('value'):
                    try:
                        dept_value = json.loads(department_col['value'])
                        if dept_value.get('index') == VANS_DEPARTMENT_ID:
                            project = self._parse_project(item)
                            if project:
                                projects.append(project)
                    except (json.JSONDecodeError, KeyError):
                        continue

            return projects

        except Exception as e:
            raise Exception(f"Failed to fetch projects from Monday.com: {str(e)}")

    def _find_column_value(self, column_values: List[Dict], column_id: str) -> Dict:
        """Find a column value by column ID."""
        for col in column_values:
            if col.get('id') == column_id:
                return col
        return None

    def _parse_project(self, item: Dict) -> Project:
        """Parse a Monday.com item into a Project object."""
        column_values = item.get('column_values', [])

        # Get basic info
        name = item.get('name', 'Unnamed Project')
        item_id = item.get('id', '')

        # Get status
        status_col = self._find_column_value(column_values, COLUMN_STATUS)
        status = 'Unknown'
        if status_col and status_col.get('text'):
            status = status_col['text']

        # Get assigned people
        people_col = self._find_column_value(column_values, COLUMN_PEOPLE)
        assigned_people = []
        if people_col and people_col.get('text'):
            # The text field contains comma-separated names
            text = people_col['text'].strip()
            if text:
                # Split by comma and clean up
                names = [name.strip() for name in text.split(',')]
                assigned_people = [name for name in names if name]

        # Get filming dates
        filming_dates = []
        for col_id in [COLUMN_FILMING_DATE_1, COLUMN_FILMING_DATE_2,
                       COLUMN_FILMING_DATE_3, COLUMN_FILMING_DATE_4]:
            date_col = self._find_column_value(column_values, col_id)
            if date_col and date_col.get('value'):
                try:
                    date_data = json.loads(date_col['value'])
                    if date_data.get('date'):
                        filming_date = datetime.strptime(date_data['date'], '%Y-%m-%d').date()
                        time_slot = date_data.get('time', '')
                        filming_dates.append(FilmingDate(date=filming_date, time_slot=time_slot))
                except (json.JSONDecodeError, ValueError, KeyError):
                    continue

        # Get editing ranges
        editing_ranges = []
        for col_id in [COLUMN_EDITING_TIME_1, COLUMN_EDITING_TIME_2, COLUMN_EDITING_TIME_3]:
            range_col = self._find_column_value(column_values, col_id)
            if range_col and range_col.get('value'):
                try:
                    range_data = json.loads(range_col['value'])
                    if range_data.get('from') and range_data.get('to'):
                        start = datetime.strptime(range_data['from'], '%Y-%m-%d').date()
                        end = datetime.strptime(range_data['to'], '%Y-%m-%d').date()
                        editing_ranges.append(EditingRange(start_date=start, end_date=end))
                except (json.JSONDecodeError, ValueError, KeyError):
                    continue

        return Project(
            id=item_id,
            name=name,
            status=status,
            assigned_people=assigned_people,
            filming_dates=filming_dates,
            editing_ranges=editing_ranges
        )
