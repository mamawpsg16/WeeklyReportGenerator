# Weekly Report Generator

Generates a branded PowerPoint weekly productivity report from a CSV log using AI analysis. Runs via Docker with a web UI.

## What it does

- Upload your prod-log CSV through a web UI
- Select month and week from auto-populated dropdowns
- Previews filtered data and hour metrics before generating
- Populates a PowerPoint template with task data across slides
- Generates a concise AI analysis paragraph using Groq (Llama 3.3)
- Download the finished PPTX directly from the browser

## Requirements

- Docker Desktop with WSL 2 backend
- A free [Groq API key](https://console.groq.com)

## Setup

1. Clone the repo
2. Add your Groq API key to a `.env` file in the project root:
   ```
   GROQ_API_KEY=your_key_here
   ```
3. Place your PowerPoint template as `weekly-report-template.pptx` in the project root

## Running

```bash
docker compose up -d --build
```

Open **http://localhost:8501** in your browser.

To stop:
```bash
docker compose down
```

## Usage

1. Upload your `prod-log.csv`
2. Select the month and week
3. Fill in presenter name and reporting date
4. Click **Generate Report**
5. Click **Download PPTX** when it's ready

## CSV Format

Required columns:

| Column | Description |
|---|---|
| Month | e.g. `05_May` |
| Week No | e.g. `Week 3` |
| Category | e.g. `Project`, `Major Enhancements` |
| Project | Application or domain name |
| Activity/Task | Task description |
| Planned Hours | Numeric |
| Actual Hours Consumed | Numeric |
| Remarks | Optional notes |

## Project Structure

```
├── app.py                        # Streamlit web UI
├── generate_report.py            # Report generation logic
├── main.py                       # CLI version (print to terminal)
├── weekly-report-template.pptx  # PowerPoint template (not committed)
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
└── .env                          # API keys (not committed)
```
