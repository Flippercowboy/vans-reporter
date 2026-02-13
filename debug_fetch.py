#!/usr/bin/env python3
"""Debug script to see what data is being fetched from Monday.com."""

import sys
from pathlib import Path
from datetime import date

# Add parent directory to path
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

from vans_reporting_app.data.monday_client import MondayClient
from vans_reporting_app.data.calculator import TimeCalculator

print("=" * 60)
print("Debug: Fetching Vans Projects from Monday.com")
print("=" * 60)

# Fetch projects
client = MondayClient()
projects = client.get_vans_projects()

print(f"\nFound {len(projects)} Vans projects:")
print()

for i, project in enumerate(projects, 1):
    print(f"{i}. {project.name}")
    print(f"   Status: {project.status}")
    print(f"   Assigned People: {project.assigned_people}")
    print(f"   Filming Dates: {len(project.filming_dates)} dates")
    for fd in project.filming_dates:
        print(f"      - {fd.date} ({fd.time_slot})")
    print(f"   Editing Ranges: {len(project.editing_ranges)} ranges")
    for er in project.editing_ranges:
        print(f"      - {er.start_date} to {er.end_date}")
    print()

# Calculate hours
print("=" * 60)
print("Calculating hours for February 2026...")
print("=" * 60)

calculator = TimeCalculator()
as_of_date = date(2026, 2, 12)  # Today
summary = calculator.calculate_project_hours(projects, as_of_date)

print(f"\nTotal Hours: {int(summary.total_hours)}h")
print(f"Completed: {int(summary.complete_hours)}h")
print(f"Remaining: {int(summary.remaining_hours)}h")
print()

print("Projects:")
for project_name, hours in summary.projects.items():
    print(f"  - {project_name}: {int(hours)}h")
print()

print("People:")
for person_name, person_hours in summary.people.items():
    print(f"  - {person_name}: {int(person_hours.total_hours)}h")
    print(f"      Complete: {int(person_hours.complete_hours)}h")
    print(f"      Remaining: {int(person_hours.remaining_hours)}h")
print()
