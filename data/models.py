"""Data models for Vans reporting application."""

from dataclasses import dataclass, field
from typing import List, Dict
from datetime import date


@dataclass
class FilmingDate:
    """Represents a single filming date."""
    date: date
    time_slot: str = ""  # "AM", "PM", "TBC", etc.


@dataclass
class EditingRange:
    """Represents an editing time range."""
    start_date: date
    end_date: date


@dataclass
class Project:
    """Represents a project with associated filming and editing dates."""
    id: str
    name: str
    status: str
    assigned_people: List[str]
    filming_dates: List[FilmingDate] = field(default_factory=list)
    editing_ranges: List[EditingRange] = field(default_factory=list)


@dataclass
class PersonHours:
    """Hours breakdown for a single person."""
    name: str
    project_hours: Dict[str, float] = field(default_factory=dict)  # project_name -> hours
    project_hours_complete: Dict[str, float] = field(default_factory=dict)
    project_hours_remaining: Dict[str, float] = field(default_factory=dict)
    total_hours: float = 0.0
    complete_hours: float = 0.0
    remaining_hours: float = 0.0


@dataclass
class ProjectSummary:
    """Complete summary of all calculated hours."""
    # By project
    projects: Dict[str, float] = field(default_factory=dict)  # project_name -> total_hours
    projects_complete: Dict[str, float] = field(default_factory=dict)
    projects_remaining: Dict[str, float] = field(default_factory=dict)

    # By person
    people: Dict[str, PersonHours] = field(default_factory=dict)

    # Totals
    total_hours: float = 0.0
    complete_hours: float = 0.0
    remaining_hours: float = 0.0

    # Metadata
    as_of_date: date = None
    month_start: date = None
    month_end: date = None
