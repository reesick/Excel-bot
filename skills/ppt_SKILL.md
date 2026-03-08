---
name: PPT Skill
description: Pure formatting engine that renders professional PowerPoint presentations via PptxGenJS
---

# PPT Skill

## What It Does
Receives a complete instruction dict from Claude Analyst and calls `build_ppt.js` (PptxGenJS) via subprocess to render a professional `.pptx` file. Claude controls the full slide structure — this skill just renders it.

## Input
```python
generate_ppt(instruction: dict) → str
```

### Input Dict Structure
| Key | Type | Description |
|-----|------|-------------|
| `title` | str | Presentation title |
| `subtitle` | str | Subtitle / date |
| `slides` | list[dict] | Slide definitions (see types below) |
| `theme` | str | `"blue_white"`, `"dark"`, or `"green_white"` |
| `chart_paths` | list[str] | Paths to chart PNG files |
| `output_path` | str | Where to save pptx |

### Slide Types
| Type | Content Format |
|------|---------------|
| `title` | No content needed |
| `kpi` | `[{label, value, icon}]` |
| `chart` | `{chart_index: int, caption: str}` |
| `insights` | `[str]` (list of insight strings) |
| `table` | `{headers: [str], rows: [[str]]}` |
| `closing` | `{message: str}` |

## Themes
- **blue_white** — Classic corporate (navy headers, white cards)
- **dark** — Navy/charcoal with bright accents
- **green_white** — Professional green palette

## Dependencies
- Node.js 18+
- `pptxgenjs` (npm global: `npm install -g pptxgenjs`)
