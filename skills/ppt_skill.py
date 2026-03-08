"""
PPT Skill — Professional Presentation Formatter
Pure formatting engine. Receives instruction dict from Claude Analyst,
calls build_ppt.js via subprocess to render a professional .pptx.

Exposes:
  generate_ppt(instruction_dict) → str (output file path)
"""

import json
import logging
import os
import subprocess

logger = logging.getLogger(__name__)

# Path to the Node.js renderer script (same directory as this file)
BUILD_SCRIPT = os.path.join(os.path.dirname(__file__), "build_ppt.js")


def generate_ppt(instruction: dict) -> str:
    """
    Generate a professional PowerPoint presentation.

    Args:
        instruction: Dict with keys:
            - title (str): Presentation title
            - subtitle (str): Subtitle / date line
            - slides (list[dict]): Slide definitions from Claude
              Each slide: {type, heading, content}
              Types: "title", "kpi", "chart", "insights", "table", "closing"
            - theme (str): "blue_white" | "dark" | "green_white"
            - chart_paths (list[str]): Paths to chart PNG files
            - output_path (str): Where to save the pptx

    Returns:
        Path to the generated .pptx file
    """
    output_path = instruction.get("output_path")
    if not output_path:
        output_path = "/tmp/outputs/presentation.pptx"

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Convert chart image paths to base64 for embedding
    chart_paths = instruction.get("chart_paths", [])
    charts_b64 = []
    for cpath in chart_paths:
        if os.path.exists(cpath):
            import base64
            with open(cpath, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("utf-8")
                charts_b64.append({"path": cpath, "base64": b64})
        else:
            charts_b64.append({"path": cpath, "base64": None})

    # Build JSON payload for Node.js script
    payload = {
        "title": instruction.get("title", "Business Analysis"),
        "subtitle": instruction.get("subtitle", "AI-Generated Report"),
        "slides": instruction.get("slides", []),
        "theme": instruction.get("theme", "blue_white"),
        "charts": charts_b64,
        "output_path": output_path,
    }

    # Call build_ppt.js via subprocess
    logger.info("Calling build_ppt.js to render %d slides...", len(payload["slides"]))

    try:
        result = subprocess.run(
            ["node", BUILD_SCRIPT],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            logger.error("build_ppt.js failed (exit %d):\nstdout: %s\nstderr: %s",
                         result.returncode, result.stdout, result.stderr)
            raise RuntimeError(
                f"PPT generation failed: {result.stderr or result.stdout}"
            )

        logger.info("PPT generated successfully: %s", output_path)
        return output_path

    except FileNotFoundError:
        raise RuntimeError(
            "Node.js is not installed or not in PATH. "
            "Install Node.js and run: npm install -g pptxgenjs"
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("PPT generation timed out (60s). Try with fewer slides.")
