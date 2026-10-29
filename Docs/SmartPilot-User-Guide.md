# SmartPilot User Guide

> SmartPilot is an AI assistant built into Microsoft Office, working across **Excel / Word / PowerPoint**.
> Describe what you want in plain language and SmartPilot handles data entry, table styling, charting, analysis, receipt recognition, and report generation — no formulas to memorize, no layout to learn.

---

## Table of Contents

1. [Product Overview](#1-product-overview)
2. [Installation & First-Time Setup](#2-installation--first-time-setup)
3. [Quick Start](#3-quick-start)
4. [Working with Tables in Natural Language](#4-working-with-tables-in-natural-language)
5. [Table Styling](#5-table-styling)
6. [Smart Charts](#6-smart-charts)
7. [Data Cleaning](#7-data-cleaning)
8. [Data Analysis](#8-data-analysis)
9. [Formula Assistant](#9-formula-assistant)
10. [Dashboards & Big Screens](#10-dashboards--big-screens)
11. [OCR Image Recognition](#11-ocr-image-recognition)
12. [Mobile Capture & Entry](#12-mobile-capture--entry)
13. [Word / PowerPoint Table Themes](#13-word--powerpoint-table-themes)
14. [PPT Report Generation](#14-ppt-report-generation)
15. [Code Generation (VBA / Python)](#15-code-generation-vba--python)
16. [Skills & Custom Skills](#16-skills--custom-skills)
17. [Knowledge Base (RAG)](#17-knowledge-base-rag)
18. [Thinking Modes & Safety](#18-thinking-modes--safety)
19. [Prompting Tips & Best Practices](#19-prompting-tips--best-practices)
20. [FAQ](#20-faq)
21. [Appendix: Common Commands](#21-appendix-common-commands)

---

## 1. Product Overview

SmartPilot deeply integrates large language models with Office. Key highlights:

- **One assistant, three apps**: a consistent experience across Excel, Word, and PowerPoint — learn once, use everywhere.
- **Natural-language driven**: describe your intent in everyday words; the AI understands and executes.
- **Deterministic engine**: exports, formatting, and filtering run on a stable engine, so results are reproducible and reliable.
- **Safety tiers**: scripted operations require confirmation in advanced modes to prevent accidents.

Typical use cases: routine report processing, data cleaning & analysis, dashboards, presentation generation, receipt/screenshot recognition, and batch file consolidation.

---

## 2. Installation & First-Time Setup

### 2.1 The Ribbon

After installation, a **SmartPilot** tab appears in Excel / Word / PowerPoint with five buttons:

| Button | Purpose |
|---|---|
| **AI Chat** | Main chat panel — where most tasks are done |
| **Image OCR** | Recognize images/screenshots into tables |
| **AI Model** | Configure the language model |
| **OCR Model** | Configure the recognition engine |
| **Mobile** | Pair your phone to snap photos into Excel |

### 2.2 Configure the model (first run)

1. Click **AI Model** to open the settings panel.
2. Enter the model endpoint and API key, select a model, and save.
3. Start chatting right away.

### 2.3 Switch UI language

- Click the 🌐 icon in the top-right, or press `Alt + L`, to toggle between Chinese and English. Ribbon buttons update accordingly.

---

## 3. Quick Start

Three steps to your first result:

1. **Select data**: highlight the range you want to work on.
2. **Open chat**: click **AI Chat** to open the side panel.
3. **Describe your need**: type it in plain language and press Enter.

**Example**: with a data table selected, type:

> Style this table with the Business Blue theme

The table is instantly restyled. The chapters below walk you through every capability.

---

## 4. Working with Tables in Natural Language

No functions or menus to memorize — just say it.

| Goal | Say |
|---|---|
| Sort | `Sort by sales from high to low` |
| Filter | `Keep only the East region data` |
| Total a column | `Add a total sum in the last column` |
| Fill blanks | `Fill blank cells in this column with 0` |
| Merge fields | `Combine column A and B into "Name - Department"` |
| Export | `Export to CSV on the desktop` |

> **Tip**: select the data range before sending, so the AI knows exactly what to work on. Exports default to the desktop.

---

## 5. Table Styling

Built-in **8 color themes × 10 layouts**, freely combinable, with fine-grained control.

### 5.1 Apply a color theme

Just name the theme:

> Style this table with the Business Blue theme

Available themes: `Business Blue`, `Tech Green`, `Vibrant Orange`, `Elegant Purple`, `Minimal Gray`, `Finance Red`, `Ink Black Gold`, `Sky Cyan`.

### 5.2 Switch layouts

Append a layout keyword to the theme:

- `Business Blue three-line table` (academic/formal)
- `Business Blue borderless style` (clean & modern)
- `Business Blue card stripe` / `Business Blue underline header`

Other layouts: standard grid, deep header, first-column focus, total row, compact report, magazine, and more.

### 5.3 Fine-tune details

> Business Blue styling, bold first column, emphasize last row, right-align numbers

Adjustable items: bold first column, emphasize last row, right-align numbers, row height, vertical centering, etc.

---

## 6. Smart Charts

Select data and describe the chart type and dimension:

| Goal | Say |
|---|---|
| Pie | `Use a pie chart to show sales by department` |
| Column | `Use a column chart to compare monthly revenue` |
| Line | `Use a line chart to show gross margin trend` |
| Combo | `Use a combo chart: columns for revenue, line for gross margin on a secondary axis` |

Supported types: pie, doughnut, column, bar, line, area, radar, scatter, bubble, combo, and more.

> **Tip**: for metrics with very different scales (e.g., amount + percentage), ask to "put on a secondary axis". You can also add an insight line, e.g., "add an insight below the chart: revenue peaked in August".

---

## 7. Data Cleaning

Handle duplicates, blanks, and inconsistent formats in one shot:

> Clean this table: remove duplicate rows, fill blanks with 0, standardize date formats

Text normalization works too:

- `Standardize "Male/Female" to M/F`
- `Trim leading and trailing spaces`

---

## 8. Data Analysis

| Capability | Say |
|---|---|
| Merge sheets | `Merge multiple worksheets into one` |
| Split | `Split data into different worksheets by department` |
| Join | `Join the orders table and customers table by customer ID` |
| Diff | `Compare the differences between the two files` |
| Descriptive stats | `Generate descriptive statistics: mean, max, min, standard deviation` |
| Multi-file rollup | `Summarize these 12 branch monthly reports into one master table` |

---

## 9. Formula Assistant

Forget the syntax? Describe the logic in plain language:

> Write a formula: if sales > 10,000 and the customer is a VIP, mark "Key", otherwise "Regular"

The AI generates the nested formula (e.g., IF with AND) and writes it into the cell. Also supports:

- `Calculate the month-over-month growth rate`
- `Write a SUMIFS to sum by condition`

---

## 10. Dashboards & Big Screens

Select data and generate a presentation-ready dashboard in one line:

> Generate a sales data analysis dashboard

The dashboard auto-composes KPI cards, trend charts, and breakdown charts. You can further:

- `Generate a business review big screen`
- `Export the dashboard as a web page/image for presentation`

---

## 11. OCR Image Recognition

Turn paper tables, screenshots, and receipts into editable data.

1. Click **Image OCR**, then drag in or choose an image.
2. The result appears as a structured table; write it back to Excel in one click.

If recognition is off, switch the engine in **OCR Model**. For fixed-layout documents, you can also customize a parsing script to batch-extract fields.

---

## 12. Mobile Capture & Entry

Quickly capture on-site data straight into Excel from your phone:

1. Click **Mobile** to show a QR code.
2. Scan it with your phone to pair.

Once connected, the core use is capture-and-entry: photograph paper tables, receipts, and documents to enter them straight into Excel — snap and go, no manual typing. Simple remote control is also supported.

---

## 13. Word / PowerPoint Table Themes

Table styling works in Word and PowerPoint too. Select a table and type:

> Style with Business Blue

Or specify a layout and theme, e.g., `Vibrant Orange three-line table`.

---

## 14. PPT Report Generation

### 14.1 One-click generation

With data selected:

> Generate a Q3 business review PPT from this data

The AI organizes cover, KPIs, charts, and conclusions, then opens the .pptx. Specify a theme:

> Use the business_blue theme

### 14.2 Step-by-step build (for complex, iterative decks)

1. `Create a PPT project: 2026 Product Launch`
2. Add pages: `Add a cover`, `Add an agenda`, `Add a chart page: revenue trend`, `Add a KPI page`, `Add a timeline page`, `Add a closing page`
3. Adjust anytime: `Change page 3's chart to a combo chart`, `Delete page 5`
4. Final render: `Regenerate the PPT`

---

## 15. Code Generation (VBA / Python)

**VBA macro**

> Write a VBA macro to sort worksheets by name

**Python script**

> Write a Python script to cluster the sales data and plot the result

> **Note**: custom scripts only run in **Deep** or **Expert** mode; lower modes block them for safety.

**OCR parsing script (advanced)**: for fixed-layout documents, customize field extraction for batch recognition:

> Generate an invoice parsing script: extract amount, tax ID, and date

---

## 16. Skills & Custom Skills

SmartPilot ships with about 26 built-in skills grouped by scenario: table styling, smart charts, data cleaning, data analysis, document theming, content generation, and data tools.

You don't need to memorize them — match by intent: **make it pretty, run analysis, work with data, or produce an output**.

### Custom Skills (SkillCreator)

Turn a repeated workflow into your own skill:

> Create a skill: auto-clean the sales table and style it as Business Blue with a total row

Once defined, just say "use my sales skill on this table" to run the whole flow.

---

## 17. Knowledge Base (RAG)

Import company docs, manuals, and policies, then ask questions grounded in them:

1. Import documents into the knowledge base.
2. Ask directly, e.g., `What is our product warranty policy?` — answers cite the source.

---

## 18. Thinking Modes & Safety

Four modes at the top of the AI Chat panel, increasing in capability and autonomy:

| Mode | Best for |
|---|---|
| Quick | Simple edits, sorting, filtering |
| Standard | Everyday analysis and formatting |
| Deep | Complex analysis, custom scripts |
| Expert | High-risk scripts such as custom Python |

**Safety tiers**: custom scripts are blocked at Standard and below; switch to Deep or Expert to run them. For daily use, staying at a lower tier is recommended.

---

## 19. Prompting Tips & Best Practices

### The golden formula

> **Clear goal + data range + desired output**

Example: `Using the selected sales detail, create a pie-chart dashboard of share by department.`

### Practical advice

- **Select data first**, then prompt.
- **Break complex tasks into steps** (e.g., create → add pages → adjust → generate for PPT).
- **Exports are reliable**: export/format/filter run on a deterministic engine — a success message means the file is created and reproducible.
- **Tune modes/engines**: switch OCR engine if recognition is off; switch to Quick mode if generation is slow.

---

## 20. FAQ

**Q: I switched language but the ribbon didn't change?**
A: The language file loads at startup. Save your work and fully close and reopen Office.

**Q: Image recognition is inaccurate?**
A: Switch the recognition engine in **OCR Model** and retry; fixed-layout documents can use a custom parsing script.

**Q: Generation is slow?**
A: Switch the thinking mode to **Quick** or **Standard**.

**Q: Export succeeded but I can't find the file?**
A: Exports default to the desktop — search by file name there.

**Q: My custom script won't run?**
A: Scripts only run in **Deep** or **Expert** mode.

**Q: Are settings shared across the three apps?**
A: The UI and capabilities are consistent; model and language settings are saved per host and can be configured separately.

---

## 21. Appendix: Common Commands

| Category | Common commands |
|---|---|
| Styling | `Style this table with the Business Blue theme` / `Tech Green styling` / `Business Blue three-line table` |
| Data ops | `Sort by sales high to low` / `Keep only East region` / `Add a total in the last column` |
| Charts | `Pie chart of sales by department` / `Combo chart: columns for revenue, line for margin on secondary axis` |
| Cleaning | `Clean this table: remove duplicates, fill blanks with 0, standardize dates` |
| Analysis | `Merge worksheets into one` / `Split by department` / `Diff the two files` |
| Formula | `Write a formula: sales > 10k and VIP marks Key` |
| Dashboard | `Generate a sales analysis dashboard` |
| OCR | Drag image to **Image OCR** or `Recognize this screenshot into a table` |
| PPT | `Generate a Q3 business review PPT from this data` / `Use the business_blue theme` |
| Code | `Write a VBA macro to rename worksheets` / `Write a Python script for clustering` |
| Skills | `Create a skill: auto-clean sales table and style as Business Blue with total row` |

---

_This guide is updated with each product release. A printable one-page cheat sheet is available for desk-side use._
