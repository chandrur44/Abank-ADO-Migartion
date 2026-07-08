# Project Rules

- Do not add "Co-Authored by Claude" in commit messages
- Do not create loose files in the project root; use subfolders
- **AWS CLI safety**: For any `aws` command you intend to run, first display the full command, explain its purpose and potential impact, and wait for explicit approval before executing
- When comparing options, use a table. When showing trade-offs, be explicit about which one you'd pick and why.
- Call out concerns proactively (security, drift, hidden cost). Don't bury them in a footnote — surface them up front.
- Initialize Git and after every response the consist of code change it should commit so that we can revert back easily.

## Memory / Context

- Context lives in a flat `memory/` folder — one Markdown file per topic
- `memory/context.md` is the root: high-level project summary + Change Log
- `memory/MEMORY.md` is the index — one-line pointers to each topic file (e.g. `project_call_cadence.md`, `user_role.md`)
- Before starting a task, read `memory/context.md` first, then read the relevant topic file(s) listed in `MEMORY.md`
- After completing a task, update `memory/context.md`'s Change Log if anything project-wide changed; update or add the relevant topic file for domain-specific changes; keep `MEMORY.md` in sync (one-line entry per file)
- Never create loose context files outside `memory/`

## Excel Output Standard

Whenever generating an Excel (`.xlsx`) deliverable, follow this convention by default — do not ask:

**File location & name**

- Save to `Documents/` folder in the current working directory, create Documents/ folder if not exist
- Filename format: `<Subject> <Identifier>.xlsx` — title case, spaces (not dashes/underscores), no date in filename
  - For AWS account inventories: `AWS Inventory <client>.xlsx` (e.g. `AWS Inventory THE.TEAM.xlsx`)
  - For other artifacts: short descriptive subject + meaningful identifier (env, region, client, etc.)

**Sheet layout (every sheet)**

- Row 1: blank, default row height
- Column A: blank, default column width
- Header row at row 2, starting column B
- Body rows from row 3 onwards
- No title cell, no merged cells, no freeze panes
- Gridlines visible across the rest of the sheet (default Excel look) — table cells additionally have thin black borders

**Styling**

- Font: Calibri 11pt throughout
- Header row: solid navy fill `#002060`, white bold text, centered, wrap-text on, row height 28
- Body cells: left-aligned, **vertically middle-aligned**, wrap-text on, thin black borders all around
- All cells (header + body): **vertically middle-aligned** for consistent visual centring
- Column widths: auto-fit to content, capped at 60 chars
- TOTAL/summary rows: bold

**Multi-sheet workbooks**

- Sheet 1 should be a **Summary** with simple counts (one column for category, one for count, plus TOTAL row)
- Detail sheets follow with descriptive names (e.g. `Resource Inventory`, `IAM Inventory`, `Tags`)
- Tagging/strategy sheets show **unique keys + values used**, not per-resource tag dumps

## Word Output Standard

The standard below is the canonical UI/UX contract for any `.docx` produced
in this project. Treat it as a hard spec — do not invent new colours,
sizes, or layout devices. If a use case is supplied, only the **content**
changes; the **format** stays identical.

### 1. Generation Approach

- Generate the document with **Python + `python-docx`** (already installed
  on the workstation, version 1.2.x).
- Always create a script named `generate_<topic>_document.py` at the
  project root and have it write the `.docx` next to it.
- The script must be **idempotent** — running it again overwrites the
  previous file.
- Use the existing reference script `generate_project_document.py` as the
  template. Reuse its helper functions (`add_h1`, `add_h2`, `add_h3`,
  `add_paragraph`, `add_bullets`, `add_numbered`, `add_table`,
  `add_kpi_strip`, `add_callout`, `add_page_break`, `configure_header_footer`)
  rather than rewriting them.
- Do **not** use Word's built-in heading styles. Build headings as plain
  paragraphs with the run properties below.

### 2. Brand Palette (do not change)

| Token        | Hex       | Use                                     |
| ------------ | --------- | --------------------------------------- |
| `NAVY`       | `#0B2A4A` | All H1 / H3 headings, table-header fill |
| `STEEL_BLUE` | `#1F4E79` | H2 sub-headings only                    |
| `ACCENT`     | `#2E75B6` | Page-header underline only              |
| `LIGHT_GREY` | `#595959` | Cover subtitle, footer, captions        |
| `DARK_TEXT`  | `#1F1F1F` | All body text and table-cell text       |
| `TABLE_ALT`  | `#EAF1F8` | Alternating table row fill (odd rows)   |
| `CALLOUT_BG` | `#F2F6FB` | Callout box fill                        |
| `DIVIDER`    | `#B7C9DB` | Table cell borders                      |

Forbidden: shaded title banners, navy hero blocks behind the cover title,
gradients, drop-shadows, decorative emojis, em-dashes (`—`), en-dashes
(`–`), and ellipsis (`…`). Use ASCII hyphens only.

### 3. Page Setup

- **Page size:** A4 (default).
- **Margins:** top 1.8 cm, bottom 1.8 cm, left 2.0 cm, right 2.0 cm.
- **Body font:** Calibri.
- **Body size:** 11 pt, line-spacing 1.3, justified
  (`WD_ALIGN_PARAGRAPH.JUSTIFY`).
- **Widow/orphan control + `keep_with_next` / `keep_together`** are
  mandatory on every paragraph helper so that two-line tails never get
  stranded on the next page.

### 4. Page Header (every page)

- Single line, 9 pt Calibri Bold, colour `NAVY`.
- **Split layout** using a right-aligned tab stop at 17 cm:
  - Left side: a short project tag, e.g. `OnDemand EDMP`.
  - Right side: the full document title, e.g.
    `Function App Authentication & Naming Architecture`.
- A thin `ACCENT` (`#2E75B6`) bottom border under the header paragraph,
  size 8 (`w:sz=8`).
- Header is the **only place** a horizontal rule appears in the running
  page chrome.

### 5. Page Footer (every page)

- Single line, 9 pt Calibri.
- Left: `Quest Engineering  |  Cloud Platform Group` in `LIGHT_GREY`.
- Right (tab stop at 17 cm): the **current page number only** (e.g. `7`)
  in `NAVY` bold, rendered via a `PAGE` field. Do **not** add `of N`,
  `NUMPAGES`, or any other field.

### 6. Cover Page (page 1)

Strict order — no blank spacer paragraphs, no banner block, no eyebrow:

1. **Document title** — 24 pt, bold, `NAVY`, left-aligned.
   `space_after = 10 pt`. `keep_with_next = True`.
2. **Subtitle paragraph** — italic, 11 pt, `LIGHT_GREY`, justified,
   line-spacing 1.3, `space_before = 0`, `space_after = 10 pt`. One short
   paragraph (3-4 sentences) describing what the document covers.
3. **Document Control** H3 + 2-column table (`Field`, `Value`) listing:
   Document Title, Project, Document Owner, Version, Issue Date, Status,
   Reviewers, Distribution. **Never** include a Classification /
   Confidential / Internal row.
4. **Revision History** H3 + 4-column table (`Version`, `Date`, `Author`,
   `Summary of Changes`). At least the latest baseline row.

No divider rule under the title. No coloured banner. The cover ends with
the Revision History table — start the next section with a page break.

### 7. Table of Contents (page 2)

- H1 `Table of Contents`, then a one-paragraph framing note.
- Manual TOC table with three columns: `#`, `Section`, `Page`. Do **not**
  insert Word's auto-generated TOC field — page numbers are written
  literally so the document looks the same when opened cold.
- Optionally follow with H3 sub-sections like `Audience` and
  `How to Read This Document` using bullet lists.

### 8. Section Headings

| Level | Helper   | Size  | Weight | Colour       | Notes                                                                                                                                                           |
| ----- | -------- | ----- | ------ | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| H1    | `add_h1` | 18 pt | Bold   | `NAVY`       | Text is **upper-cased** in code. No underline. `space_before=14 pt`, `space_after=8 pt`, always `keep_with_next` + `keep_together`. Numbered (`1.`, `2.`, ...). |
| H2    | `add_h2` | 14 pt | Bold   | `STEEL_BLUE` | Used sparingly between H1 and H3.                                                                                                                               |
| H3    | `add_h3` | 12 pt | Bold   | `NAVY`       | Sub-section labels (`6.1`, `6.2`, ...).                                                                                                                         |

Headings are left-aligned. **No** horizontal rule above or below any
heading. Do not page-break before a heading; let `keep_with_next` plus
widow control handle pagination.

### 9. Body Paragraphs

- Calibri 11 pt, `DARK_TEXT`, justified, line-spacing 1.3.
- `space_after = 8 pt`, `space_before = 0`.
- Limit to 4-5 lines per paragraph; break larger blocks into two
  paragraphs or convert to a bullet list.
- Use the helper `add_paragraph(...)` — it sets justification,
  widow-control and spacing automatically.

### 10. Lists

- **Bulleted lists** (`add_bullets`): Word's `List Bullet` style; every
  bullet 11 pt, `DARK_TEXT`, justified, `space_after = 3 pt`,
  line-spacing 1.25.
- A bullet may be a tuple `(lead, rest)`. The lead is rendered **bold,
  `NAVY`**, followed by the rest in normal `DARK_TEXT`. Use the lead-in
  pattern for "Term - explanation" bullets.
- **Numbered lists** (`add_numbered`): Word's `List Number` style with
  the same typography as bullets.
- All list items are bound with `keep_with_next` to the next item except
  the last, so a single bullet can never be orphaned.

### 11. Tables

Tables are the document's primary data device. Build them with `add_table`.

- **Layout:** fixed (`w:tblLayout type="fixed"`) with an explicit
  `w:tblW` set to the sum of the column widths and a `w:tblGrid` whose
  `w:gridCol` widths match. This guarantees the navy header fill reaches
  the full row width.
- **Cell margins:** top/bottom 40 dxa, left/right 80 dxa.
- **Header row:**
  - Fill: `NAVY` (`0B2A4A`).
  - Text: 10.5 pt Calibri Bold, white (`#FFFFFF`), left-aligned.
  - Marked as `tblHeader` so it repeats when a table spans pages.
  - `cantSplit` set on the header row.
- **Data rows:**
  - Text: 10 pt Calibri, `DARK_TEXT`, left-aligned, line-spacing 1.2.
  - Alternating row fill: rows where `r % 2 == 1` get `TABLE_ALT`
    (`#EAF1F8`); even rows stay white.
  - Border: 1 pt `DIVIDER` (`#B7C9DB`) on all four edges.
  - Every row has `cantSplit` so a single row never breaks across pages.
- Always pass explicit `col_widths` (a list of `Cm(...)` values) so the
  fixed layout works.
- Always center-align the table on the page
  (`WD_TABLE_ALIGNMENT.CENTER`).
- Insert a 6 pt spacer paragraph after every table to keep the rhythm.

### 12. Callouts

Use the `add_callout(title, body)` helper for any "Why this matters",
"Note", or "Warning" tile.

- 1x1 table, fill `CALLOUT_BG` (`#F2F6FB`).
- Left edge: thick `ACCENT` border (`w:sz=24`); other edges 1 pt
  `DIVIDER`.
- Title: 11 pt Calibri Bold, `NAVY`.
- Body: 10.5 pt Calibri, `DARK_TEXT`, justified, line-spacing 1.3.
- Followed by a 6 pt spacer paragraph.

### 13. KPI Strip

Use `add_kpi_strip([(label, value), ...])` for the "platform at a
glance" tile rows on the executive summary.

- Single row of 3-6 tiles.
- Top half of each tile: navy fill, white 18 pt bold value, centred.
- Bottom half of each tile: light blue fill (`#EAF1F8`), 9 pt navy bold
  label, centred.
- Both rows have `cantSplit`.
- Followed by a 6 pt spacer paragraph.

### 14. Pagination & Spacing Rules

- Never use blank empty paragraphs to push content. Use `space_before` /
  `space_after` on real paragraphs, or the spacer paragraph emitted by
  `add_table` / `add_callout` / `add_kpi_strip`.
- `add_page_break(doc)` between major sections only — typically before
  every numbered H1 from "1. Executive Summary" onwards.
- Widow-control is on by default in every helper — do not disable it.
- Tables, callouts and KPI strips must not split across pages
  (enforced via `cantSplit` on every row).

### 15. Tone & Content Rules

- Everything is justified — left-rag body text is forbidden.
- No emojis anywhere.
- No `—` or `–` characters; replace with `-`.
- No `…`; replace with three ASCII dots `...`.
- Avoid the word "Confidential" or "Internal" in any chrome (header,
  footer, cover table). Document Control row "Classification" must
  not appear.
- Bullet leads use a hyphen separator, e.g. `Authentication is universal -`.

### 16. Recommended 10-Section Skeleton

When the user asks for a "10-page document about <X>", reuse this skeleton
unless they say otherwise. Each H1 is a numbered section.

1. **Executive Summary** - 3-4 paragraphs + KPI strip + headline-finding
   bullets + one callout.
2. **<Domain> Overview & Scope** - context paragraph, scope / out-of-scope
   bullets, optionally a "components in scope" table.
3. **Solution Architecture** - layered table (`Layer`, `Component`,
   `Responsibility`), numbered request-lifecycle list, design-principle
   callout.
4. **Core Mechanism #1** (e.g. authentication / data model / pipeline) -
   parameters table, salient-fields bullets, per-environment table,
   operational note callout.
5. **Core Mechanism #2** (e.g. policies / rules / interactions) -
   pattern A vs pattern B paragraphs, comparison table, sub-bullets,
   threat-model or assumption callout.
6. **Standards & Conventions** - prefix logic paragraph, abbreviations
   table, naming-pattern table, worked-example table.
7. **Environment / Configuration Matrix** - purpose table, baseline
   matrix, identifiers reference table.
8. **Identity / RBAC / Secrets** - topology table, RBAC assignments
   table, principles bullets.
9. **Risks, Findings & Recommendations** - risk register table
   (`#`, `Finding`, `Impact`, `Likelihood`, `Priority`),
   recommendations bullets with bold leads.
10. **Operational Runbook & Glossary** - first-responder numbered list,
    change-control bullets, glossary table, closing line `- End of
document -` centred in `LIGHT_GREY` italic.

### 17. File Output

- Save the file at the project root with the name
  `<Project>_<Topic>_<Subject>.docx` using only ASCII characters and
  underscores (no spaces).
- Print the absolute path on stdout when generation completes:
  `Document saved to: <path>`.
- Never write a Markdown copy of the same content unless the user asks
  for it.
