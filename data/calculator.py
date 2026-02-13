"""Time calculation logic for project hours."""

from datetime import date, timedelta
from typing import List, Dict, Tuple
from collections import defaultdict
from .models import Project, ProjectSummary, PersonHours
from ..config import *


class TimeCalculator:
    """Calculates project hours with overlap detection and weekday filtering."""

    def __init__(self):
        """Initialise the calculator."""
        pass

    def calculate_project_hours(self, projects: List[Project], as_of_date: date) -> ProjectSummary:
        """
        Calculate hours for all projects split by completed and remaining.
        Only counts dates within the same month as as_of_date.

        Args:
            projects: List of Project objects
            as_of_date: Date to split completed vs remaining hours

        Returns:
            ProjectSummary with all calculated data
        """
        # Determine month boundaries
        month_start = date(as_of_date.year, as_of_date.month, 1)

        # Get last day of month
        if as_of_date.month == 12:
            next_month = date(as_of_date.year + 1, 1, 1)
        else:
            next_month = date(as_of_date.year, as_of_date.month + 1, 1)
        month_end = next_month - timedelta(days=1)

        # Build daily schedule for all people (filtered to this month only)
        daily_schedule = self._build_daily_schedule(projects, month_start, month_end)

        # Resolve overlapping work
        resolved_hours = self._resolve_conflicts(daily_schedule)

        # Aggregate by project and person
        summary = self._aggregate_hours(resolved_hours, projects, as_of_date)

        return summary

    def _get_weekdays_in_range(self, start: date, end: date) -> List[date]:
        """
        Get all weekday dates (Mon-Fri) in a range.

        Args:
            start: Start date
            end: End date

        Returns:
            List of weekday dates
        """
        weekdays = []
        current = start

        while current <= end:
            # Monday = 0, Sunday = 6
            if current.weekday() < 5:  # Mon-Fri only
                weekdays.append(current)
            current += timedelta(days=1)

        return weekdays

    def _build_daily_schedule(self, projects: List[Project], month_start: date, month_end: date) -> Dict:
        """
        Build a schedule mapping (person, date) -> list of (project, activity_type, nominal_hours).
        Only includes dates within the specified month.

        Args:
            projects: List of Project objects
            month_start: First day of the month
            month_end: Last day of the month

        Returns:
            Dict mapping (person, date) to list of activities
        """
        schedule = defaultdict(list)

        for project in projects:
            # Process filming dates
            for filming_date in project.filming_dates:
                film_date = filming_date.date
                # Only include weekdays within the target month
                if (film_date.weekday() < 5 and
                    month_start <= film_date <= month_end):
                    for person in project.assigned_people:
                        schedule[(person, film_date)].append({
                            'project': project.name,
                            'type': 'filming',
                            'nominal_hours': FILMING_HOURS_PER_DAY
                        })

            # Process editing ranges
            for edit_range in project.editing_ranges:
                # Clip the editing range to the target month
                range_start = max(edit_range.start_date, month_start)
                range_end = min(edit_range.end_date, month_end)

                # Only process if range overlaps with target month
                if range_start <= range_end:
                    weekdays = self._get_weekdays_in_range(range_start, range_end)
                    for edit_date in weekdays:
                        for person in project.assigned_people:
                            schedule[(person, edit_date)].append({
                                'project': project.name,
                                'type': 'editing',
                                'nominal_hours': EDITING_HOURS_PER_DAY
                            })

        return schedule

    def _resolve_conflicts(self, daily_schedule: Dict) -> Dict:
        """
        Resolve overlapping work by splitting available hours equally.

        Args:
            daily_schedule: Dict from _build_daily_schedule

        Returns:
            Dict mapping (person, date, project) to actual allocated hours
        """
        resolved = {}

        for (person, day), activities in daily_schedule.items():
            # Calculate total nominal hours for this person-day
            total_nominal = sum(act['nominal_hours'] for act in activities)

            if total_nominal <= MAX_HOURS_PER_DAY:
                # No conflict, use nominal hours
                for act in activities:
                    key = (person, day, act['project'])
                    if key in resolved:
                        resolved[key] += act['nominal_hours']
                    else:
                        resolved[key] = act['nominal_hours']
            else:
                # Conflict: split MAX_HOURS_PER_DAY equally among all activities
                hours_per_activity = MAX_HOURS_PER_DAY / len(activities)
                for act in activities:
                    key = (person, day, act['project'])
                    if key in resolved:
                        resolved[key] += hours_per_activity
                    else:
                        resolved[key] = hours_per_activity

        return resolved

    def _aggregate_hours(self, resolved_hours: Dict, projects: List[Project],
                        as_of_date: date) -> ProjectSummary:
        """
        Aggregate resolved hours by project and person, split by completed/remaining.

        Args:
            resolved_hours: Dict from _resolve_conflicts
            projects: List of Project objects
            as_of_date: Date to split completed vs remaining

        Returns:
            ProjectSummary object
        """
        # Determine month boundaries
        # Assume we're working with the month containing as_of_date
        month_start = date(as_of_date.year, as_of_date.month, 1)

        # Get last day of month
        if as_of_date.month == 12:
            next_month = date(as_of_date.year + 1, 1, 1)
        else:
            next_month = date(as_of_date.year, as_of_date.month + 1, 1)
        month_end = next_month - timedelta(days=1)

        # Initialize aggregation structures
        project_totals = defaultdict(float)
        project_complete = defaultdict(float)
        project_remaining = defaultdict(float)
        people_hours = defaultdict(lambda: PersonHours(name=''))

        # Aggregate hours
        for (person, day, project_name), hours in resolved_hours.items():
            # Update project totals
            project_totals[project_name] += hours

            if day <= as_of_date:
                project_complete[project_name] += hours
            else:
                project_remaining[project_name] += hours

            # Update person totals
            if person not in people_hours:
                people_hours[person] = PersonHours(name=person)

            person_hours = people_hours[person]
            person_hours.total_hours += hours

            if project_name not in person_hours.project_hours:
                person_hours.project_hours[project_name] = 0
            person_hours.project_hours[project_name] += hours

            if day <= as_of_date:
                person_hours.complete_hours += hours
                if project_name not in person_hours.project_hours_complete:
                    person_hours.project_hours_complete[project_name] = 0
                person_hours.project_hours_complete[project_name] += hours
            else:
                person_hours.remaining_hours += hours
                if project_name not in person_hours.project_hours_remaining:
                    person_hours.project_hours_remaining[project_name] = 0
                person_hours.project_hours_remaining[project_name] += hours

        # Calculate grand totals
        total_hours = sum(project_totals.values())
        complete_hours = sum(project_complete.values())
        remaining_hours = sum(project_remaining.values())

        return ProjectSummary(
            projects=dict(project_totals),
            projects_complete=dict(project_complete),
            projects_remaining=dict(project_remaining),
            people=dict(people_hours),
            total_hours=total_hours,
            complete_hours=complete_hours,
            remaining_hours=remaining_hours,
            as_of_date=as_of_date,
            month_start=month_start,
            month_end=month_end
        )
