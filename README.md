# Excel Bot 🤖

AI-powered Telegram bot that turns **any** business file into professional dashboards and presentations — powered by Claude.

## What It Does

Upload any CSV, XLS, or XLSX file to Telegram → get back:
- 📊 **Excel Dashboard** — KPI cards, charts, pivot tables, insights (4 sheets)
- 🎨 **PowerPoint Presentation** — 5-8 slides, 3 color themes, ready for meetings
- 💡 **Key Insights** — plain English findings a CEO would care about

No questions asked. Claude reads the data like a senior analyst and figures everything out.

## Architecture

```
User uploads file → Claude API (Code Execution)
                      ↓
              Cleans, analyzes, generates charts
                      ↓
            ┌─────────┴─────────┐
      Excel Skill          PPT Skill
    (formatting only)    (PptxGenJS)
            ↓                  ↓
     Dashboard.xlsx    Presentation.pptx
            └─────────┬─────────┘
                      ↓
              Sent back via Telegram
```

## Setup

### Prerequisites
- Python 3.10+
- Node.js 18+
- Telegram bot token ([@BotFather](https://t.me/BotFather))
- Anthropic API key ([console.anthropic.com](https://console.anthropic.com))

### Local Development

```bash
# Install dependencies
pip install -r requirements.txt
npm install -g pptxgenjs

# Configure
cp .env.example .env
# Fill in TELEGRAM_BOT_TOKEN and ANTHROPIC_API_KEY

# Run (polling mode)
python main.py
```

### Deploy to Render

1. Push to GitHub
2. Create Web Service on [render.com](https://render.com) → connect repo
3. Set env vars: `TELEGRAM_BOT_TOKEN`, `ANTHROPIC_API_KEY`, `WEBHOOK_SECRET`, `APP_URL`
4. Deploy — webhook registers automatically

## Project Structure

```
├── skills/
│   ├── analyst.py          # Claude brain (Files API + Code Execution)
│   ├── excel_skill.py      # Excel dashboard formatter
│   ├── ppt_skill.py        # PPT formatter (calls build_ppt.js)
│   └── build_ppt.js        # PptxGenJS slide renderer
├── bot/
│   ├── handlers.py         # Telegram message routing
│   └── sessions.py         # In-memory session state
├── main.py                 # Entry point (polling + webhook modes)
├── requirements.txt
├── render.yaml
└── .env.example
```

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `TELEGRAM_BOT_TOKEN` | Yes | From @BotFather |
| `ANTHROPIC_API_KEY` | Yes | Claude API key |
| `LOCAL_MODE` | No | `true` for polling, omit for webhook |
| `WEBHOOK_SECRET` | Render only | Random string for webhook URL |
| `APP_URL` | Render only | Public URL on Render |

## License

Private project.
