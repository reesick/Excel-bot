---
name: Excel Skill
description: Pure formatting engine that produces professional multi-sheet Excel dashboards
---

# Excel Skill

## What It Does
Receives a complete instruction dict from the Claude Analyst and produces a professional 4-sheet Excel file. Zero analysis logic — only formatting.

## Input
```python
generate_excel(instruction: dict) → str
```

### Input Dict Structure
| Key | Type | Description |
|-----|------|-------------|
| `file_path` | str | Path to cleaned CSV |
| `charts` | list[str] | Paths to chart PNG files |
| `insights` | list[str] | Plain English insight strings |
| `kpis` | list[dict] | KPI cards: `{label, value, icon}` |
| `pivot_instructions` | list[dict] | `{title, index_col, value_col, agg}` |
| `dataset_title` | str | Report title |
| `dataset_type` | str | Data category |
| `output_path` | str | Where to save xlsx |

## Output Sheets
1. **Raw Data** — Cleaned data as formatted table, alternating rows, bold headers
2. **Dashboard** — Title banner, KPI cards, chart PNGs in 2-column grid
3. **Pivot Tables** — One per pivot instruction, styled headers
4. **Insights** — Numbered insight cards with clean typography

## Design
- Palette: Navy `#1F3864`, Blue `#2E75B6`, Light blue `#D6E4F0`
- Font: Arial
- Gridlines: Off on all sheets

## Dependencies
- `pandas`, `openpyxl`
