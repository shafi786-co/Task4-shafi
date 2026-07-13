# Project 4: Data Visualization (Optional Milestone)

**DecodeLabs Data Analytics Internship — Batch 2026**
Author: Shafi

## Overview

This project turns the findings from Projects 2 and 3 into a boardroom-ready slide deck —
the goal isn't more analysis, it's communicating three decisions that already exist in the
data as clearly and honestly as possible.

## Design principles followed

- **One message per slide.** Each slide has a single action title that states the conclusion,
  not just the topic (e.g. "Revenue Held Flat Across 30 Months," not "Monthly Revenue").
- **Situation-Complication-Resolution structure.** The deck opens with an executive summary
  framed as SCR, then proves each part with a chart, then closes with next steps.
- **Chart type matches the question.** Bar charts for comparing categories (product revenue,
  order status), a line chart for the trend over time. No pie charts.
- **Zero-baseline axis.** Every bar chart starts at 0 so bar height isn't misleading.
- **Direct labeling over legends.** Values are printed on/above the bars instead of requiring
  a separate legend lookup.
- **Color used as a spotlight, not decoration.** Backgrounds are dominant navy or white; a
  single gold accent marks what matters, and red is reserved only for the fulfillment problem.
- **Minimal chartjunk.** No 3D effects, no heavy gridlines, no unnecessary borders.
- **"So What?" and methodology slides.** The deck ends with concrete next steps and closes
  the loop by naming the exact design disciplines (chart selection, data-ink ratio,
  storytelling structure) used to build it.

## Files in this repo

| File | Description |
|---|---|
| `Project4_DataVisualization.pptx` | The 9-slide deck |
| `visualization_data.py` | Python script producing the exact numbers behind each chart |
| `build_deck.js` | pptxgenjs script that generates the deck (for editing/regenerating) |

## Slide-by-slide

1. **Title** — Order Performance: A Data Storytelling Review
2. **Executive Summary** — Situation / Complication / Resolution, in three cards
3. **The Numbers at a Glance** — four stat callouts: $1.26M revenue, 1,200 orders, $1,050 avg order value, 41% Cancelled/Pending
4. **Chairs and Printers Lead, But No Category Carries the Business Alone** — bar chart, product revenue
5. **Revenue Held Flat Across 30 Months** — line chart, monthly revenue, with a peak annotation
6. **Cancelled and Pending Orders Consume 41% of Volume** — bar chart + stat card, the core problem
7. **So What — Three Actions for Next Quarter** — recommendations
8. **How This Was Built** — the three design disciplines behind every chart (Architect / Editor / Storyteller), plus data source and tools
9. **Thank You**

## How to regenerate

```bash
npm install pptxgenjs
node build_deck.js
python3 /mnt/skills/public/pptx/scripts/office/validate.py Project4_DataVisualization.pptx
```

`visualization_data.py` can be run independently to re-check the numbers behind any chart:

```bash
pip install pandas openpyxl
python visualization_data.py
```

## Tools used

pptxgenjs (native PowerPoint charts, not images), Python/pandas for the underlying numbers.
