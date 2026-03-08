/**
 * build_ppt.js — PptxGenJS slide renderer
 * Called by ppt_skill.py via subprocess.
 * Input:  JSON payload from stdin
 * Output: writes .pptx to payload.output_path
 *
 * Handles slide types: title, kpi, chart, insights, table, closing
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// Read JSON payload from stdin
let inputData = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => { inputData += chunk; });
process.stdin.on("end", () => {
  try {
    const payload = JSON.parse(inputData);
    buildPresentation(payload);
  } catch (e) {
    console.error("Failed to parse JSON input:", e.message);
    process.exit(1);
  }
});

// ─── Theme definitions ────────────────────────────────────────────────────────

const THEMES = {
  "blue_white": {
    bg: "FFFFFF", titleBg: "1F3864", titleFg: "FFFFFF",
    accent: "2E75B6", accentLight: "D6E4F0",
    text: "1A1A2E", subtext: "5B6E8C",
    slide_bg: "F7FAFE", card_bg: "FFFFFF",
    chart: ["1F3864", "2E75B6", "4472C4", "5B9BD5", "9DC3E6", "ED7D31"],
    name: "Blue & White",
  },
  "dark": {
    bg: "1A1A2E", titleBg: "16213E", titleFg: "E8E8E8",
    accent: "0F3460", accentLight: "1A1A3E",
    text: "E8E8E8", subtext: "A0A0B0",
    slide_bg: "1A1A2E", card_bg: "16213E",
    chart: ["0F3460", "E94560", "533483", "1A8FE3", "F39C12", "2ECC71"],
    name: "Dark Theme",
  },
  "green_white": {
    bg: "FFFFFF", titleBg: "2C5F2D", titleFg: "FFFFFF",
    accent: "4CAF50", accentLight: "E8F5E9",
    text: "1B3A1C", subtext: "4A7C59",
    slide_bg: "F5FBF5", card_bg: "FFFFFF",
    chart: ["2C5F2D", "4CAF50", "81C784", "A5D6A7", "388E3C", "66BB6A"],
    name: "Green & White",
  },
};

function getTheme(name) {
  return THEMES[name] || THEMES["blue_white"];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 45, color: "000000", opacity: 0.15 };
}

// ─── Slide builders ───────────────────────────────────────────────────────────

function buildTitleSlide(pres, T, slide) {
  const s = pres.addSlide();
  // Full background
  s.background = { fill: T.titleBg };

  // Accent line
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 2.0, w: 9.0, h: 0.04,
    fill: { color: T.accent },
  });

  // Title
  s.addText(slide.heading || "Business Analysis", {
    x: 0.5, y: 1.2, w: 9.0, h: 0.7,
    fontSize: 32, fontFace: "Calibri", color: T.titleFg,
    bold: true, align: "left",
  });

  // Subtitle from payload
  s.addText(slide.subtitle || "", {
    x: 0.5, y: 2.3, w: 9.0, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: T.subtext,
    italic: true, align: "left",
  });
}

function buildKpiSlide(pres, T, slide) {
  const s = pres.addSlide();
  s.background = { fill: T.slide_bg };

  // Header bar
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.8,
    fill: { color: T.titleBg },
  });
  s.addText(slide.heading || "Key Metrics", {
    x: 0.5, y: 0.1, w: 9, h: 0.6,
    fontSize: 20, fontFace: "Calibri", color: T.titleFg, bold: true,
  });

  const kpis = slide.content || [];
  const cardW = 2.0;
  const cardH = 1.4;
  const gap = 0.25;
  const totalW = kpis.length * cardW + (kpis.length - 1) * gap;
  let startX = (10 - totalW) / 2;

  kpis.forEach((kpi, i) => {
    const x = startX + i * (cardW + gap);
    const y = 1.8;

    // Card background
    s.addShape(pres.ShapeType.rect, {
      x, y, w: cardW, h: cardH,
      fill: { color: T.card_bg },
      shadow: makeShadow(),
      rectRadius: 0.1,
    });

    // Accent bar at top of card
    s.addShape(pres.ShapeType.rect, {
      x, y, w: cardW, h: 0.06,
      fill: { color: T.accent },
    });

    // Icon + Label
    s.addText(`${kpi.icon || "📊"}  ${kpi.label || ""}`, {
      x, y: y + 0.15, w: cardW, h: 0.4,
      fontSize: 10, fontFace: "Calibri", color: T.subtext,
      align: "center", valign: "middle",
    });

    // Value
    s.addText(kpi.value || "", {
      x, y: y + 0.55, w: cardW, h: 0.7,
      fontSize: 22, fontFace: "Calibri", color: T.accent,
      bold: true, align: "center", valign: "middle",
    });
  });
}

function buildChartSlide(pres, T, slide, charts) {
  const s = pres.addSlide();
  s.background = { fill: T.slide_bg };

  // Header bar
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.8,
    fill: { color: T.titleBg },
  });
  s.addText(slide.heading || "Analysis", {
    x: 0.5, y: 0.1, w: 9, h: 0.6,
    fontSize: 20, fontFace: "Calibri", color: T.titleFg, bold: true,
  });

  const content = slide.content || {};
  const chartIdx = content.chart_index !== undefined ? content.chart_index : 0;
  const caption = content.caption || "";

  // Embed chart image if available
  if (charts[chartIdx] && charts[chartIdx].base64) {
    s.addImage({
      data: "image/png;base64," + charts[chartIdx].base64,
      x: 0.5, y: 1.0, w: 9.0, h: 4.2,
    });
  } else {
    s.addText("[ Chart not available ]", {
      x: 1, y: 2, w: 8, h: 1,
      fontSize: 14, fontFace: "Calibri", color: T.subtext,
      align: "center", italic: true,
    });
  }

  // Caption
  if (caption) {
    s.addText(caption, {
      x: 0.5, y: 5.3, w: 9.0, h: 0.4,
      fontSize: 11, fontFace: "Calibri", color: T.subtext,
      italic: true, align: "center",
    });
  }
}

function buildInsightsSlide(pres, T, slide) {
  const s = pres.addSlide();
  s.background = { fill: T.slide_bg };

  // Header bar
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.8,
    fill: { color: T.titleBg },
  });
  s.addText(slide.heading || "Key Findings", {
    x: 0.5, y: 0.1, w: 9, h: 0.6,
    fontSize: 20, fontFace: "Calibri", color: T.titleFg, bold: true,
  });

  const insights = slide.content || [];
  const startY = 1.2;
  const cardH = 0.65;
  const gap = 0.15;

  insights.forEach((insight, i) => {
    const y = startY + i * (cardH + gap);

    // Card background
    s.addShape(pres.ShapeType.rect, {
      x: 0.5, y, w: 9.0, h: cardH,
      fill: { color: T.card_bg },
      shadow: makeShadow(),
      rectRadius: 0.08,
    });

    // Number badge
    s.addShape(pres.ShapeType.rect, {
      x: 0.5, y, w: 0.5, h: cardH,
      fill: { color: T.accent },
      rectRadius: 0.08,
    });
    s.addText(String(i + 1), {
      x: 0.5, y, w: 0.5, h: cardH,
      fontSize: 16, fontFace: "Calibri", color: "FFFFFF",
      bold: true, align: "center", valign: "middle",
    });

    // Insight text
    s.addText(insight, {
      x: 1.2, y, w: 8.1, h: cardH,
      fontSize: 12, fontFace: "Calibri", color: T.text,
      valign: "middle",
    });
  });
}

function buildTableSlide(pres, T, slide) {
  const s = pres.addSlide();
  s.background = { fill: T.slide_bg };

  // Header bar
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.8,
    fill: { color: T.titleBg },
  });
  s.addText(slide.heading || "Data Table", {
    x: 0.5, y: 0.1, w: 9, h: 0.6,
    fontSize: 20, fontFace: "Calibri", color: T.titleFg, bold: true,
  });

  const content = slide.content || {};
  const headers = content.headers || [];
  const rows = content.rows || [];

  if (headers.length === 0) return;

  // Build table data
  const tableRows = [];
  // Header row
  tableRows.push(headers.map(h => ({
    text: h, options: {
      bold: true, fontSize: 10, fontFace: "Calibri",
      color: "FFFFFF", fill: { color: T.titleBg },
      align: "center", valign: "middle",
    }
  })));

  // Data rows
  rows.forEach((row, ri) => {
    tableRows.push(row.map(cell => ({
      text: String(cell), options: {
        fontSize: 10, fontFace: "Calibri",
        color: T.text,
        fill: { color: ri % 2 === 0 ? T.accentLight : T.card_bg },
        align: "center", valign: "middle",
      }
    })));
  });

  const colW = Math.min(9.0 / headers.length, 2.5);

  s.addTable(tableRows, {
    x: 0.5, y: 1.2, w: colW * headers.length,
    fontSize: 10, fontFace: "Calibri",
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: 0.35,
    autoPage: false,
  });
}

function buildClosingSlide(pres, T, slide) {
  const s = pres.addSlide();
  s.background = { fill: T.titleBg };

  s.addText(slide.heading || "Thank You", {
    x: 0.5, y: 1.5, w: 9.0, h: 1.0,
    fontSize: 36, fontFace: "Calibri", color: T.titleFg,
    bold: true, align: "center", valign: "middle",
  });

  const content = slide.content || {};
  const message = content.message || "Generated by Office Bot AI Analyst";

  // Accent line
  s.addShape(pres.ShapeType.rect, {
    x: 3.5, y: 2.7, w: 3.0, h: 0.04,
    fill: { color: T.accent },
  });

  s.addText(message, {
    x: 1.0, y: 3.0, w: 8.0, h: 0.6,
    fontSize: 14, fontFace: "Calibri", color: T.subtext,
    italic: true, align: "center",
  });
}

// ─── Main builder ─────────────────────────────────────────────────────────────

function buildPresentation(payload) {
  const T = getTheme(payload.theme);
  const pres = new pptxgen();

  pres.layout = "LAYOUT_WIDE";
  pres.author = "Office Bot";
  pres.company = "AI Analyst";
  pres.subject = payload.title || "Business Analysis";

  const slides = payload.slides || [];
  const charts = payload.charts || [];

  slides.forEach((slide) => {
    // Inject subtitle from payload into title slide
    if (slide.type === "title") {
      slide.subtitle = payload.subtitle || "";
    }

    switch (slide.type) {
      case "title":
        buildTitleSlide(pres, T, slide);
        break;
      case "kpi":
        buildKpiSlide(pres, T, slide);
        break;
      case "chart":
        buildChartSlide(pres, T, slide, charts);
        break;
      case "insights":
        buildInsightsSlide(pres, T, slide);
        break;
      case "table":
        buildTableSlide(pres, T, slide);
        break;
      case "closing":
        buildClosingSlide(pres, T, slide);
        break;
      default:
        console.error(`Unknown slide type: ${slide.type}`);
    }
  });

  // Save
  const outputPath = payload.output_path || "presentation.pptx";
  const dir = path.dirname(outputPath);
  if (dir && !fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  pres.writeFile({ fileName: outputPath })
    .then(() => {
      console.log(`Saved: ${outputPath}`);
      process.exit(0);
    })
    .catch((err) => {
      console.error("Failed to write pptx:", err.message);
      process.exit(1);
    });
}
