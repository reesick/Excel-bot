# Excel Bot

An AI-powered Telegram bot that transforms any business data file into professional Excel dashboards and PowerPoint presentations. Built with Claude's code execution capabilities for intelligent, fully automated data analysis.

## Overview

Excel Bot accepts CSV, XLS, and XLSX files of any structure or quality, from any industry. The user uploads a file via Telegram and receives back a multi-sheet Excel dashboard and a slide deck — no technical knowledge required, no questions asked.

The bot handles everything: detecting file structure, cleaning messy data, choosing appropriate chart types, generating visualizations, writing executive-level insights, and formatting publication-ready output files.

## Methodology

### Core Principle: Claude as the Analyst

Traditional dashboard tools require users to specify columns, chart types, and metrics manually. Excel Bot takes a fundamentally different approach — Claude acts as a senior data analyst who reads the raw data, understands its structure, and makes all analytical decisions autonomously.

The system is split into three layers:

1. **Analysis Layer** (`analyst.py`) — Claude receives the raw file via the Anthropic Files API, then uses the Code Execution tool to run Python code inside a sandboxed environment. Claude writes and executes its own pandas/matplotlib code to clean the data, detect patterns, generate charts, compute KPIs, and write insights. All outputs (cleaned CSV, chart PNGs, analysis JSON) are retrieved via the Files API.

2. **Formatting Layer** (`excel_skill.py`, `ppt_skill.py`) — Pure formatting engines with zero analysis logic. They receive a complete instruction package from the analysis layer and produce professionally styled output files. The Excel skill uses openpyxl to build a 4-sheet workbook. The PPT skill calls a Node.js script (PptxGenJS) via subprocess to render slides.

3. **Orchestration Layer** (`handlers.py`) — Routes Telegram messages, manages session state, and coordinates the pipeline. Sends typing indicators during long operations so the user knows the bot is working.

### Data Flow

```
1. User uploads file (CSV / XLS / XLSX) via Telegram
                        |
2. File downloaded from Telegram servers to local storage
                        |
3. File uploaded to Anthropic Files API
                        |
4. Claude receives file via container_upload content block
   with code_execution tool enabled
                        |
5. Claude autonomously:
   - Detects encoding, separators, header rows
   - Cleans data (removes blanks, fixes types, standardizes formats)
   - Identifies what kind of business data this is
   - Generates 4-6 matplotlib charts as PNG files
   - Computes 3-4 KPIs with formatted values
   - Writes 4-6 executive-level insights
   - Defines pivot table structures
   - Designs a full PowerPoint slide structure
   - Saves cleaned data as CSV
   - Returns a JSON instruction package
                        |
6. Output files (charts, CSV) downloaded from Files API
                        |
7. Instruction package passed to Excel Skill
   -> Produces 4-sheet dashboard (Raw Data, Dashboard, Pivots, Insights)
                        |
8. Same package passed to PPT Skill
   -> Produces 5-8 slide presentation via PptxGenJS
                        |
9. Both files sent back to user via Telegram
   along with a text summary of key findings
```

### Handling Long-Running Analysis

Claude's code execution can take 30-90 seconds for complex files. The system handles this through:

- **Typing indicators**: A background asyncio task sends Telegram's `sendChatAction` every 5 seconds so the user sees the bot is actively working.
- **pause_turn handling**: If Claude's execution exceeds API time limits, the API returns a `pause_turn` stop reason. The bot loops — re-sending the conversation to let Claude continue — accumulating file IDs and text from all turns until `end_turn` is received.

### Error Handling

Every skill call is wrapped in try/except. If Excel generation fails, PPT generation still proceeds (and vice versa). The user always receives a response, even on errors. Full tracebacks are logged with `exc_info=True`. Session state is reset on unrecoverable errors.

## Technical Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| Runtime | Python 3.10+ | Core application |
| Web Framework | FastAPI | Webhook server (production) |
| AI Engine | Claude API (claude-sonnet-4-6) | Data analysis and code execution |
| File Processing | Anthropic Files API | Upload/download files to Claude sandbox |
| Code Execution | code_execution_20250825 | Claude runs Python in sandboxed environment |
| Excel Output | openpyxl | Multi-sheet dashboard generation |
| PPT Output | PptxGenJS (Node.js) | Slide rendering via subprocess |
| Telegram | httpx (direct API) | No SDK dependency, raw HTTP calls |
| Deployment | Render | Web service with Python + Node.js |

## Project Structure

```
office-bot/
├── skills/
│   ├── analyst.py           # Claude brain - Files API + Code Execution
│   ├── analyst_SKILL.md     # Documentation
│   ├── excel_skill.py       # Excel dashboard formatter (openpyxl)
│   ├── excel_SKILL.md       # Documentation
│   ├── ppt_skill.py         # PPT formatter (calls build_ppt.js)
│   ├── ppt_SKILL.md         # Documentation
│   └── build_ppt.js         # PptxGenJS slide renderer
├── bot/
│   ├── __init__.py
│   ├── handlers.py          # Telegram routing and pipeline orchestration
│   └── sessions.py          # In-memory conversation state
├── main.py                  # Entry point (LOCAL_MODE=true: polling, else: webhook)
├── requirements.txt
├── render.yaml              # Render deployment config
├── .env.example             # Template for environment variables
└── .gitignore
```

## Setup

### Prerequisites

- Python 3.10+
- Node.js 18+ (for PPT generation)
- A Telegram bot token from [@BotFather](https://t.me/BotFather)
- An Anthropic API key from [console.anthropic.com](https://console.anthropic.com)

### Local Development

```bash
# Install Python dependencies
pip install -r requirements.txt

# Install Node.js dependency (global)
npm install -g pptxgenjs

# Configure environment
cp .env.example .env
# Fill in TELEGRAM_BOT_TOKEN and ANTHROPIC_API_KEY
# Set LOCAL_MODE=true for polling mode

# Run the bot
python main.py
```

### Deploy to Render

1. Push to GitHub
2. Create a Web Service on [render.com](https://render.com) and connect the repo
3. Render auto-detects `render.yaml` for build and start commands
4. Set environment variables in the Render dashboard:
   - `TELEGRAM_BOT_TOKEN`
   - `ANTHROPIC_API_KEY`
   - `WEBHOOK_SECRET` (generate with `openssl rand -hex 16`)
   - `APP_URL` (your Render public URL)
5. Deploy — the webhook registers automatically on startup

### Verify Deployment

```bash
# Health check
curl https://your-app.onrender.com/health

# Webhook status
curl https://api.telegram.org/bot<TOKEN>/getWebhookInfo
```

## Environment Variables

| Variable | Required For | Description |
|----------|-------------|-------------|
| `TELEGRAM_BOT_TOKEN` | All modes | Bot token from BotFather |
| `ANTHROPIC_API_KEY` | All modes | Claude API key |
| `LOCAL_MODE` | Local only | Set to `true` for polling mode |
| `WEBHOOK_SECRET` | Render only | Random string for webhook URL security |
| `APP_URL` | Render only | Public URL (e.g. `https://excel-bot.onrender.com`) |

## Bot Commands

| Input | Action |
|-------|--------|
| `/start` | Welcome message |
| `/cancel` | Reset session, delete temp files |
| Upload `.xlsx` / `.xls` / `.csv` | Full analysis pipeline |
| "make a ppt" | Regenerate presentation from last analysis |
| Any question about the data | Claude answers with specific numbers |

## License

Private project — not for public distribution.
