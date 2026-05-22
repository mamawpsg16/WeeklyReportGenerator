# Weekly Report Generator

Generate weekly productivity reports from CSV logs. Compares planned vs actual hours by category.

## What it does

- Reads productivity log (CSV)
- Filters by month and week
- Groups activities by category
- Shows planned vs actual hours
- Calculates variance (over/under)

## Setup

1. Clone repo
2. Create virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Update variables in `main.py`:
   ```python
   month = '05_May'  # Change month
   week = 'Week 1'   # Change week
   ```

2. Place CSV in parent folder (or update path)

3. Run:
   ```bash
   python main.py
   ```

## Output

Shows report by category:
- Project name
- Planned hours
- Actual hours
- Remarks
- Total with variance

Example:
```
================================================================================
Category: Project
================================================================================
                              Project  Planned Hours  Actual Hours Consumed Remarks
[SFA-App] Sales Return & Replace            20.0                   13.0
[SFA-Web] Expenses Revamp                    3.0                    1.0  Support UAT

Total: Planned=23.0, Actual=14.0, Variance=-9.0
```

## CSV Format

Required columns:
- Month
- Week No
- Category
- Project
- Activity/Task
- Planned Hours
- Actual Hours Consumed
- Remarks

## Next Steps

- Export to PDF/PowerPoint
- Email reports automatically
- Add chart visualization
