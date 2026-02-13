# Vans Department Monthly Reporting App

A desktop application for automating monthly resource allocation reports for the Vans department.

## Features

- Fetch project data from Monday.com automatically
- Calculate time allocations with custom rules:
  - Filming: 4 hours per day
  - Editing: 8 hours per day
  - Maximum 8 hours per person per day
  - Weekdays only (Monday-Friday)
  - Split overlapping work equally between projects
- Preview and edit calculated hours before generating reports
- Generate PowerPoint presentations matching your existing style

## Installation

1. Ensure Python 3.7+ is installed
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. **Get your Monday.com API token**:
   - Go to https://monday.com
   - Click your profile picture (bottom left)
   - Select "Developers" → "My Access Tokens"
   - Click "Show" next to your token and copy it

4. **Configure the API token** (choose one method):

   **Option A: Environment Variable (Recommended)**
   ```bash
   export MONDAY_API_TOKEN="your_token_here"
   python3 run_app.py
   ```

   **Option B: Edit config.py**
   Open `config.py` and set:
   ```python
   MONDAY_API_TOKEN = "your_token_here"
   ```

## Usage

Run the application using the run script:
```bash
python3 run_app.py
```

Or make it executable and run directly:
```bash
chmod +x run_app.py
./run_app.py
```

### Workflow

1. **Fetch Data**: Click "Fetch Vans Projects" to retrieve data from Monday.com
2. **Review & Edit**: Click "Preview Data" to review calculated hours and make adjustments
3. **Generate**: Click "Generate Presentation" to create the PowerPoint file

## Configuration

The app is pre-configured for the Vans department (Board ID: 1215254769).
To modify settings, edit `config.py`.

## Requirements

- Python 3.7 or higher
- python-pptx library
- Access to Monday.com via MCP tools (claude CLI with Monday.com MCP server configured)

**Note**: The Monday.com data fetching currently requires the Claude CLI with Monday.com MCP server properly configured. If you don't have this set up, the app will show an error when attempting to fetch projects. In that case, you may need to configure the Monday.com API access separately.

## Notes

- The app automatically detects the current month and splits hours into "completed so far" and "remaining"
- The "as of" date defaults to today and can be refreshed using the "Refresh Date" button
- All calculations follow the established rules: weekdays only, 8h max per day, equal splitting of overlaps
