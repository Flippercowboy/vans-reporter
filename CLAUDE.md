# Vans Department Monthly Reporting App

## Overview

This is a desktop GUI application that automates monthly resource allocation reports for the Vans department. It fetches project data from Monday.com, calculates time allocations using complex rules, and generates professional PowerPoint presentations.

## Quick Start

### Running the App

**Option 1: Double-click launcher** (Easiest)
```
Double-click: "Launch Vans Reporter.command"
Located in: ~/Desktop/Monday.com stuff/
```

**Option 2: From command line**
```bash
cd ~/Desktop/Monday.com\ stuff/vans_reporting_app
python3 run_app.py
```

**Option 3: Spotlight search**
```
⌘ + Space → Type "Vans Reporter" → Enter
```

## What It Does

1. **Fetches Data** - Retrieves all Vans department projects from Monday.com Project Tracker board (ID: 1215254769)
2. **Calculates Hours** - Applies complex time allocation rules (see below)
3. **Allows Editing** - Preview and adjust hours before generating reports
4. **Generates PowerPoint** - Creates 9-slide professional presentations matching existing style

## Key Features

- ✅ **Month Selector** - Generate reports for any month (past, present, or future)
- ✅ **Smart Calculations** - Weekdays only, overlap splitting, proper rounding
- ✅ **Dual Editing** - Edit hours at project level OR team member level
- ✅ **Auto-Recalculation** - All totals stay synchronized
- ✅ **Professional Output** - PowerPoint with charts, tables, and styling

## Time Calculation Rules

### Core Rules
- **Filming**: 4 hours per day
- **Editing**: 8 hours per day
- **Maximum**: 8 hours per person per day
- **Weekdays Only**: Monday-Friday (weekends excluded)
- **Overlap Handling**: When multiple projects occur on the same day, split the available 8 hours equally

### Example
If Rolf has 3 projects on Tuesday, Feb 12:
- Available: 8 hours
- Projects: Van Talent (editing), Vito Sport X (editing), Vans Call (filming)
- Allocation: 8h ÷ 3 = 2.67h per project

### Completed vs Remaining
- **Completed**: Hours up to and including the "as of" date
- **Remaining**: Hours after the "as of" date
- The "as of" date can be set to any day within the selected month

## Configuration

### Monday.com API Token
Location: `config.py` line 9
```python
MONDAY_API_TOKEN = "your_token_here"
```

To update your token:
1. Go to https://monday.com → Profile → Developers → My Access Tokens
2. Copy your token
3. Update `config.py` or set environment variable `MONDAY_API_TOKEN`

### Board Configuration
- **Board ID**: 1215254769 (Project Tracker)
- **Department Filter**: Vans (ID: 17)
- **Column IDs**: Defined in `config.py`

## File Structure

```
vans_reporting_app/
├── main.py                 # Alternative entry point
├── run_app.py              # Main entry point (use this)
├── config.py               # Configuration (API token, board IDs, rules)
├── requirements.txt        # Python dependencies
├── README.md              # User documentation
├── CLAUDE.md              # This file - developer documentation
│
├── gui/                   # User interface
│   ├── main_window.py     # Main 3-step workflow window
│   └── data_preview.py    # Editable data preview/edit dialog
│
├── data/                  # Data layer
│   ├── models.py          # Data classes (Project, PersonHours, ProjectSummary)
│   ├── monday_client.py   # Monday.com GraphQL API client
│   └── calculator.py      # Time allocation calculation engine
│
└── powerpoint/            # PowerPoint generation
    ├── generator.py       # Main presentation generator (9 slides)
    └── slides.py          # Reusable slide creation functions
```

## How It Works

### 1. Data Fetching (monday_client.py)
- Makes GraphQL requests to Monday.com API
- Filters for Vans department (single_select5 = 17)
- Extracts filming dates and editing time ranges
- Parses assigned people from text field

### 2. Time Calculation (calculator.py)
- Builds daily schedule for all people and projects
- Filters to weekdays only (Mon-Fri)
- Detects overlapping work
- Splits available hours equally when conflicts occur
- Sums hours by project and person
- Splits into completed/remaining based on "as of" date

### 3. PowerPoint Generation (generator.py)
- Creates 9-slide presentation:
  1. Title slide with date
  2. Month overview with pie chart
  3. Team member summary with table and bar chart
  4. Hours completed (stacked bar chart)
  5. Remaining hours (stacked bar chart)
  6. Project breakdown with progress
  7. Detailed schedule for top contributor
  8. Weekly progress tracking
  9. Key insights and recommendations

## Common Tasks

### Updating Time Calculation Rules
Edit `config.py`:
```python
FILMING_HOURS_PER_DAY = 4  # Change filming hours
EDITING_HOURS_PER_DAY = 8  # Change editing hours
MAX_HOURS_PER_DAY = 8      # Change max hours per day
```

### Changing PowerPoint Styling
Edit `config.py`:
```python
HEADER_COLOR = (0, 112, 192)      # Table header colour (RGB)
ALT_ROW_COLOR = (217, 225, 242)   # Alternating row colour (RGB)
WHITE_COLOR = (255, 255, 255)     # Text colour (RGB)
```

### Adding More Slides
1. Create new method in `powerpoint/generator.py`
2. Add call in `generate()` method
3. Use helper functions from `powerpoint/slides.py`

### Modifying Calculation Logic
Edit `data/calculator.py`:
- `_build_daily_schedule()` - How dates are processed
- `_resolve_conflicts()` - How overlaps are handled
- `_aggregate_hours()` - How totals are calculated

## Troubleshooting

### App Won't Launch
- Check Python 3 is installed: `python3 --version`
- Verify file permissions: `chmod +x run_app.py`
- Check for errors: `python3 run_app.py` (run from terminal to see errors)

### "Failed to fetch projects" Error
- Verify API token is set in `config.py`
- Check token is valid at https://monday.com/developers
- Ensure internet connection is active

### Calculations Don't Match
- Click "Recalculate Totals" button in preview window
- Verify "as of" date is set correctly
- Check that dates in Monday.com are within the selected month

### PowerPoint Has Formatting Issues
- Verify python-pptx is installed: `pip3 install python-pptx`
- Check colour values in `config.py` are valid RGB tuples
- Ensure no "\n" characters in text (should use spacing instead)

## Technical Notes

### Rounding
- All hours are rounded to 1 decimal place internally (0.1h precision)
- Display uses `round()` not `int()` to ensure complete + remaining = total
- Proportional adjustments preserve accuracy

### Data Sync
- Projects are recalculated from people's contributions
- People totals are recalculated from their project hours
- Source of truth: person-project hour assignments

### Month Boundaries
- Only dates within the selected month are counted
- Editing ranges spanning multiple months are clipped to the target month
- Weekend dates are filtered out during calculation

## Dependencies

- **Python 3.7+** - Core language
- **python-pptx** - PowerPoint generation
- **tkinter** - GUI (included with Python)
- **urllib** - HTTP requests (included with Python)

Install dependencies:
```bash
pip3 install -r requirements.txt
```

## Version History

**v1.0** (February 2026)
- Initial release
- Month selector
- Dual editing (projects and team members)
- Proper rounding and calculation sync
- 9-slide PowerPoint generation
- Direct Monday.com API integration

## Support

If you encounter issues or need modifications:
1. Check this CLAUDE.md file for common tasks
2. Review README.md for user-facing documentation
3. Check config.py for configuration options
4. Examine error messages in terminal output

## Future Enhancements (Not Implemented)

- Multi-department support
- Custom date range selection (beyond single month)
- Excel/CSV export option
- Historical comparison charts
- Email integration for automatic sending
- Web version for team access
- Automatic scheduling/recurring reports

## Notes for Claude

When continuing work on this project:
- User is Jason (use UK English spellings)
- Calculation rules are critical - maintain 4h filming, 8h editing, 8h max/day
- PowerPoint style must match existing presentations exactly
- Rounding must be precise to avoid complete + remaining ≠ total
- Month filtering is essential - don't include dates outside target month
- API token is already configured in config.py
- User prefers direct, specific feedback without hedging
