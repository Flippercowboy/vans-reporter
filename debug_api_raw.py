#!/usr/bin/env python3
"""Debug script to see raw API response for people column."""

import sys
from pathlib import Path
import json

parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

from vans_reporting_app.data.monday_client import MondayClient

client = MondayClient()

# Make a simple query to get one project with people info
query = """
query {
  boards(ids: [1215254769]) {
    items_page(limit: 5) {
      items {
        id
        name
        column_values(ids: ["people4"]) {
          id
          text
          value
        }
      }
    }
  }
}
"""

data = client._make_request(query)
items = data["boards"][0]["items_page"]["items"]

print("Raw People Column Data:")
print("=" * 60)

for item in items:
    print(f"\nProject: {item['name']}")
    for col in item['column_values']:
        if col['id'] == 'people4':
            print(f"  Text: {col['text']}")
            print(f"  Value: {col['value']}")
            if col['value']:
                try:
                    parsed = json.loads(col['value'])
                    print(f"  Parsed: {json.dumps(parsed, indent=4)}")
                except:
                    pass
