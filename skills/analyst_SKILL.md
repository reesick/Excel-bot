---
name: Claude Analyst
description: AI brain layer that analyzes business files using Anthropic Files API + Code Execution
---

# Claude Analyst Skill

## What It Does
Takes any raw business file (CSV, XLS, XLSX) and sends it to Claude via Anthropic's Files API + Code Execution tool. Claude cleans, analyzes, generates charts, writes insights, and returns a complete instruction package ready to pass directly to the Excel and PPT formatting skills.

## API Calls Made

### 1. Upload File
`POST https://api.anthropic.com/v1/files` (multipart/form-data)  
Beta header: `files-api-2025-04-14`

### 2. Analyze with Code Execution
`POST https://api.anthropic.com/v1/messages`  
- Model: `claude-sonnet-4-6`
- Tool: `code_execution_20250825`
- File referenced via `container_upload` content block
- Handles `pause_turn` stop reason for long-running executions

### 3. Download Output Files
`GET https://api.anthropic.com/v1/files/{file_id}/content`  
Downloads chart PNGs + cleaned CSV from Claude's sandbox.

## Input
```python
analyze_and_plan(file_path: str, session_id: str) → dict
```

## Output Dict
```python
{
    "analysis_summary": str,       # 2-3 sentence summary for Telegram
    "file_path": str,              # Path to cleaned CSV
    "charts": [str],               # Paths to chart PNG files
    "insights": [str],             # Plain English insights
    "kpis": [{"label", "value", "icon"}],
    "pivot_instructions": [{"title", "index_col", "value_col", "agg"}],
    "dataset_title": str,
    "dataset_type": str,
    "title": str,                  # For PPT
    "subtitle": str,               # For PPT
    "slides": [{"type", "heading", "content"}],  # For PPT
    "theme": str,                  # For PPT
    "chart_paths": [str],          # For PPT (same as charts)
    "charts_generated": int,
    "output_dir": str,
}
```

## Follow-up Questions
```python
ask_question(question: str, csv_path: str, session_id: str) → str
```

## Dependencies
- `anthropic>=0.40.0`
- `ANTHROPIC_API_KEY` environment variable
