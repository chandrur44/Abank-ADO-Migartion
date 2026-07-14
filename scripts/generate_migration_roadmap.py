"""Generate ADO -> GitHub Migration Roadmap Excel workbook.

Excel Output Standard (CLAUDE.md):
- Save to Documents/
- Filename: ADO to GitHub Migration Roadmap Abank.xlsx
- Row 1 blank, Col A blank, header at B2, body from B3
- Calibri 11, navy #002060 header white bold centered wrap, row height 28
- Body left-aligned, vertically middle-aligned, wrap-text, thin black borders
- Auto-fit cols capped at 60, TOTAL rows bold
- Sheet 1 = Summary, detail sheets follow
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DOCUMENTS_DIR = PROJECT_ROOT / "Documents"
DOCUMENTS_DIR.mkdir(exist_ok=True)
OUTPUT = DOCUMENTS_DIR / "ADO to GitHub Migration Roadmap Abank.xlsx"

NAVY = "002060"
WHITE = "FFFFFF"
LIGHT_BLUE = "DCE6F1"
YELLOW = "FFEB9C"
GREEN = "C6EFCE"
ORANGE = "FCE4D6"

thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
BODY_FONT = Font(name="Calibri", size=11)
BODY_BOLD = Font(name="Calibri", size=11, bold=True)
TOTAL_FONT = Font(name="Calibri", size=11, bold=True)

HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

PHASE_FILLS = {
    "Phase 0": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "Phase 1": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Phase 2": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "Phase 3": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "Phase 4": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "Phase 5": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Phase 6": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "Phase 7": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "Phase 8": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
}


def write_sheet(ws, headers, rows, phase_col_idx=None, bold_rows=None):
    bold_rows = bold_rows or set()
    for col_idx, header in enumerate(headers, start=2):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
    ws.row_dimensions[2].height = 28

    for r_off, row in enumerate(rows):
        er = 3 + r_off
        phase_val = row[phase_col_idx] if phase_col_idx is not None and phase_col_idx < len(row) else None
        phase_fill = None
        if phase_val and str(phase_val) in PHASE_FILLS:
            phase_fill = PHASE_FILLS[str(phase_val)]

        for c_off, val in enumerate(row):
            cell = ws.cell(row=er, column=2 + c_off, value=val)
            cell.font = BODY_BOLD if r_off in bold_rows else BODY_FONT
            cell.alignment = BODY_ALIGN
            cell.border = BORDER
            if phase_fill:
                cell.fill = phase_fill

    for col_idx in range(2, 2 + len(headers)):
        letter = get_column_letter(col_idx)
        max_len = len(str(headers[col_idx - 2]))
        for r_off in range(len(rows)):
            val = rows[r_off][col_idx - 2] if col_idx - 2 < len(rows[r_off]) else ""
            if val is None:
                continue
            for line in str(val).split("\n"):
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[letter].width = min(max_len + 2, 60)


def d(offset_weeks, base=date(2026, 7, 14)):
    return (base + timedelta(weeks=offset_weeks)).strftime("%d %b %Y")


wb = Workbook()

# ─────────────────────────────────────────────
# Sheet 1: Summary
# ─────────────────────────────────────────────
ws_sum = wb.active
ws_sum.title = "Summary"
summary_rows = [
    ("Client", "Amalgamated Bank (Abank)"),
    ("Source", "Azure DevOps Server on-prem (abdevopsdev, API 5.1)"),
    ("Target", "GitHub Enterprise Cloud (Standard)"),
    ("Total Repos", "327"),
    ("Active Repos", "309 (excl. 18 empty repos)"),
    ("ADO Projects", "24"),
    ("Migration Approach", "git push --mirror from VDI jump host (GEI not supported for on-prem)"),
    ("Batch Size", "10 repos per batch"),
    ("Total Batches", "33 batches (327 / 10, rounded up)"),
    ("Total Phases", "8 (Phase 0 to Phase 7)"),
    ("Planned Start", d(0)),
    ("Planned End", d(22)),
    ("Total Duration", "~22 weeks"),
]
write_sheet(ws_sum, ["Item", "Detail"], summary_rows)

# ─────────────────────────────────────────────
# Sheet 2: Roadmap (master plan)
# ─────────────────────────────────────────────
ws_road = wb.create_sheet("Roadmap")

roadmap = [
    # S.No, Phase, Activity, Description, Owner, Start, End, Status, Expected Output
    (1,  "Phase 0", "Kickoff & Governance",       "Align stakeholders on migration scope, approach (git mirror), timeline, and decision owners. Get sign-off on: master->main rename, empty repo disposition, collision rename, PR history acceptance.",                                                 "Project Lead",    d(0),  d(1),  "Not Started", "Signed-off scope document, decision log"),
    (2,  "Phase 0", "GitHub Org Setup",            "Provision GitHub Enterprise Cloud org (zeb-ai or abank). Configure Entra SSO + SCIM. Create org-level settings: default branch = main, secret scanning on, push protection on.",                                                                  "Platform Team",   d(0),  d(1),  "Not Started", "GitHub org live, SSO working"),
    (3,  "Phase 0", "VDI Tooling Setup",           "Install embeddable Python on VDI. Install dependencies (requests, python-dotenv, openpyxl). Clone migration repo. Verify .env with ADO PAT and GitHub PAT.",                                                                                      "Engineer",        d(0),  d(1),  "Not Started", "Scripts runnable on VDI"),
    (4,  "Phase 0", "Inventory Finalization",      "Review extracted inventory (327 repos). Identify and confirm: 18 empty repos (skip/migrate), 3 collision renames, stale single-repo projects, wave groupings.",                                                                                   "Engineer + Lead", d(0),  d(1),  "Not Started", "Finalized Repo Inventory sheet with wave assignments"),

    (5,  "Phase 1", "Script Development - Mirror", "Develop git_mirror_migrate.py: iterates repo list, runs git clone --mirror from abdevopsdev and git push --mirror to GitHub. Supports dry-run mode, retry on failure, progress log, resume from last completed repo.",                          "Engineer",        d(1),  d(3),  "Not Started", "git_mirror_migrate.py in scripts/"),
    (6,  "Phase 1", "Script Development - Topics", "Develop apply_topics_teams.py: reads workbook Target Topics + Target Team(s) columns, calls gh api to set topics and team permissions per repo. Dry-run mode.",                                                                                   "Engineer",        d(1),  d(3),  "Not Started", "apply_topics_teams.py in scripts/"),
    (7,  "Phase 1", "Script Development - Ruleset","Develop apply_rulesets.py: creates 3-4 Ruleset JSON templates (default-protect, require-review, file-size, comment-required), attaches to repos by topic. Dry-run mode.",                                                                       "Engineer",        d(1),  d(3),  "Not Started", "apply_rulesets.py + ruleset JSON templates in github-config/"),
    (8,  "Phase 1", "Script Development - Verify", "Develop post_migration_verify.py: for each migrated repo checks branch count, tag count, default branch, topics applied, team access, Ruleset attached. Outputs pass/fail per repo.",                                                           "Engineer",        d(2),  d(3),  "Not Started", "post_migration_verify.py in scripts/"),
    (9,  "Phase 1", "Unit Testing - All Scripts",  "Test all four scripts against zeb-ai sandbox org (already on GitHub). Validate dry-run, retry, resume, error handling. Fix any issues before pilot.",                                                                                            "Engineer",        d(3),  d(4),  "Not Started", "All scripts passing on sandbox org"),

    (10, "Phase 2", "Pilot - Batch 0 (2 repos)",  "Select 2 low-risk repos (1 from ScriptRefactoring + 1 small single-repo project). Run full end-to-end: mirror -> topics -> Ruleset -> verify. Validate with repo owner. Retro findings fed back into scripts.",                                "Engineer",        d(4),  d(5),  "Not Started", "2 repos live on GitHub, verify script green"),
    (11, "Phase 2", "Pilot Retro & Script Fix",   "Review pilot output: any missing branches/tags, topic errors, Ruleset attach failures, team permission issues. Fix scripts. Update runbook.",                                                                                                     "Engineer",        d(5),  d(6),  "Not Started", "Updated scripts, updated runbook"),

    (12, "Phase 3", "Batch 1 (Repos 1-10)",       "Single-repo stale projects (last commit > 2 years). Freeze ADO, mirror, apply topics + teams + Ruleset, verify, archive ADO repos. Notify owners.",                                                                                            "Engineer",        d(6),  d(7),  "Not Started", "10 repos on GitHub, ADO archived"),
    (13, "Phase 3", "Batch 2 (Repos 11-20)",      "Continue single-repo projects (moderately stale, 1-2 years). Same process: freeze -> mirror -> topics/teams/Ruleset -> verify -> archive.",                                                                                                    "Engineer",        d(7),  d(8),  "Not Started", "10 repos on GitHub, ADO archived"),
    (14, "Phase 3", "Batch 3 (Repos 21-30)",      "SSIS ETL project (5 repos) + remaining single-repo projects. Same cutover process.",                                                                                                                                                            "Engineer",        d(8),  d(9),  "Not Started", "10 repos on GitHub, ADO archived"),
    (15, "Phase 3", "Batch 4 (Repos 31-40)",      "Non Binary (2 repos) + active single-repo projects (last commit < 1 year). Increase freeze notice to 48h.",                                                                                                                                    "Engineer",        d(9),  d(10), "Not Started", "10 repos on GitHub, ADO archived"),
    (16, "Phase 3", "Batch 5 (Repos 41-50)",      "Remaining active single-repo projects. Same process.",                                                                                                                                                                                          "Engineer",        d(10), d(11), "Not Started", "10 repos on GitHub, ADO archived"),

    (17, "Phase 4", "Batch 6 (Repos 51-60)",      "ScriptRefactoring project begins. 10 repos from SR. Freeze -> mirror -> topics (project-scriptrefactoring) -> Ruleset -> verify -> archive.",                                                                                                  "Engineer",        d(11), d(12), "Not Started", "10 SR repos on GitHub"),
    (18, "Phase 4", "Batch 7 (Repos 61-70)",      "ScriptRefactoring: next 10 repos.",                                                                                                                                                                                                            "Engineer",        d(12), d(13), "Not Started", "10 SR repos on GitHub"),
    (19, "Phase 4", "Batch 8 (Repos 71-80)",      "ScriptRefactoring: next 10 repos.",                                                                                                                                                                                                            "Engineer",        d(13), d(14), "Not Started", "10 SR repos on GitHub"),
    (20, "Phase 4", "Batch 9 (Repos 81-90)",      "ScriptRefactoring: next 10 repos.",                                                                                                                                                                                                            "Engineer",        d(14), d(15), "Not Started", "10 SR repos on GitHub"),
    (21, "Phase 4", "Batch 10 (Repos 91-106)",    "ScriptRefactoring: final 16 repos. SR project fully migrated.",                                                                                                                                                                                "Engineer",        d(15), d(16), "Not Started", "All 56 SR repos on GitHub"),

    (22, "Phase 5", "Batch 11 (Repos 107-116)",   "AB AppDev begins. First 10 repos (low-traffic, oldest last-commit). Freeze -> mirror -> topics (project-ab-appdev) -> verify -> archive.",                                                                                                     "Engineer",        d(16), d(17), "Not Started", "10 AB AppDev repos on GitHub"),
    (23, "Phase 5", "Batch 12 (Repos 117-126)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(16), d(17), "Not Started", "10 AB AppDev repos on GitHub"),
    (24, "Phase 5", "Batch 13 (Repos 127-136)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(17), d(18), "Not Started", "10 AB AppDev repos on GitHub"),
    (25, "Phase 5", "Batch 14 (Repos 137-146)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(17), d(18), "Not Started", "10 AB AppDev repos on GitHub"),
    (26, "Phase 5", "Batch 15 (Repos 147-156)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(18), d(19), "Not Started", "10 AB AppDev repos on GitHub"),
    (27, "Phase 5", "Batch 16 (Repos 157-166)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(18), d(19), "Not Started", "10 AB AppDev repos on GitHub"),
    (28, "Phase 5", "Batch 17 (Repos 167-176)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(19), d(20), "Not Started", "10 AB AppDev repos on GitHub"),
    (29, "Phase 5", "Batch 18 (Repos 177-186)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(19), d(20), "Not Started", "10 AB AppDev repos on GitHub"),
    (30, "Phase 5", "Batch 19 (Repos 187-196)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(20), d(21), "Not Started", "10 AB AppDev repos on GitHub"),
    (31, "Phase 5", "Batch 20 (Repos 197-206)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(20), d(21), "Not Started", "10 AB AppDev repos on GitHub"),
    (32, "Phase 5", "Batch 21 (Repos 207-216)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (33, "Phase 5", "Batch 22 (Repos 217-226)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (34, "Phase 5", "Batch 23 (Repos 227-236)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (35, "Phase 5", "Batch 24 (Repos 237-246)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (36, "Phase 5", "Batch 25 (Repos 247-256)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (37, "Phase 5", "Batch 26 (Repos 257-266)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (38, "Phase 5", "Batch 27 (Repos 267-276)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (39, "Phase 5", "Batch 28 (Repos 277-286)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (40, "Phase 5", "Batch 29 (Repos 287-296)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (41, "Phase 5", "Batch 30 (Repos 297-306)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (42, "Phase 5", "Batch 31 (Repos 307-316)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (43, "Phase 5", "Batch 32 (Repos 317-326)",   "AB AppDev: next 10 repos.",                                                                                                                                                                                                                    "Engineer",        d(21), d(22), "Not Started", "10 AB AppDev repos on GitHub"),
    (44, "Phase 5", "Batch 33 (Repo 327)",        "AB AppDev: final repo. AB AppDev project fully migrated.",                                                                                                                                                                                     "Engineer",        d(21), d(22), "Not Started", "All 244 AB AppDev repos on GitHub"),

    (45, "Phase 6", "Full Org Audit",             "Run post_migration_verify.py across all 327 repos. Check: every repo has project-* topic, correct default branch, Ruleset attached, correct team access. Fix any gaps.",                                                                        "Engineer",        d(22), d(23), "Not Started", "Audit report - all repos green"),
    (46, "Phase 6", "Security Review",            "Confirm: no repo left public unintentionally, secret scanning enabled org-wide, push protection enabled, no hardcoded secrets in migrated repos (GitHub secret scan alert review).",                                                            "Security",        d(22), d(23), "Not Started", "Security sign-off"),
    (47, "Phase 6", "Consumer Redirect",          "Update all consumer references pointing to abdevopsdev: submodule URLs, CI/CD pipeline clone URLs, package feed references, README badges. Raise PRs per affected repo.",                                                                       "Team Leads",      d(22), d(24), "Not Started", "All consumers redirected to GitHub"),

    (48, "Phase 7", "ADO Server Archive",         "Set all migrated ADO repos to read-only (disable push). Communicate to all teams: ADO is now archive-only. Retain for 90 days as reference.",                                                                                                   "Platform Team",   d(24), d(24), "Not Started", "ADO repos read-only, comms sent"),
    (49, "Phase 7", "Runbook Handover",           "Deliver operational runbook to platform team: how to add new repos, apply topics, request team access, raise Ruleset changes. Conduct handover session.",                                                                                       "Engineer",        d(24), d(25), "Not Started", "Runbook doc + handover session"),
    (50, "Phase 7", "ADO Decommission Planning",  "Plan final ADO Server decommission (90 days post archive). Confirm data retention policy with Abank. Schedule infra team for server removal.",                                                                                                  "Abank Infra",     d(24), d(26), "Not Started", "Decommission plan signed off"),
]

write_sheet(
    ws_road,
    ["S.No", "Phase", "Activity", "Description", "Owner", "Start Date", "End Date", "Status", "Expected Output"],
    roadmap,
    phase_col_idx=1,
)

# ─────────────────────────────────────────────
# Sheet 3: Batch Plan
# ─────────────────────────────────────────────
ws_batch = wb.create_sheet("Batch Plan")
batch_rows = [
    ("Pilot", "Phase 2", "2 repos", "1 from ScriptRefactoring + 1 small single-repo project", "End-to-end process validation", d(4), d(5)),
    ("Batch 1",  "Phase 3", "Repos 1-10",    "Single-repo stale projects (>2 yrs idle)",                      "Low risk warm-up",             d(6),  d(7)),
    ("Batch 2",  "Phase 3", "Repos 11-20",   "Single-repo projects (1-2 yrs idle)",                           "Low risk",                     d(7),  d(8)),
    ("Batch 3",  "Phase 3", "Repos 21-30",   "SSIS ETL (5) + remaining single-repo projects",                 "Low-medium risk",              d(8),  d(9)),
    ("Batch 4",  "Phase 3", "Repos 31-40",   "Non Binary (2) + active single-repo projects",                  "Medium risk, 48h freeze",      d(9),  d(10)),
    ("Batch 5",  "Phase 3", "Repos 41-50",   "Remaining active single-repo projects",                         "Medium risk",                  d(10), d(11)),
    ("Batch 6",  "Phase 4", "Repos 51-60",   "ScriptRefactoring: first 10",                                   "Medium risk",                  d(11), d(12)),
    ("Batch 7",  "Phase 4", "Repos 61-70",   "ScriptRefactoring: next 10",                                    "Medium risk",                  d(12), d(13)),
    ("Batch 8",  "Phase 4", "Repos 71-80",   "ScriptRefactoring: next 10",                                    "Medium risk",                  d(13), d(14)),
    ("Batch 9",  "Phase 4", "Repos 81-90",   "ScriptRefactoring: next 10",                                    "Medium risk",                  d(14), d(15)),
    ("Batch 10", "Phase 4", "Repos 91-106",  "ScriptRefactoring: final 16",                                   "SR fully migrated",            d(15), d(16)),
    ("Batch 11", "Phase 5", "Repos 107-116", "AB AppDev: first 10 (oldest last-commit)",                      "Higher risk, 48h freeze",      d(16), d(17)),
    ("Batch 12", "Phase 5", "Repos 117-126", "AB AppDev: next 10",                                            "Higher risk",                  d(16), d(17)),
    ("Batch 13", "Phase 5", "Repos 127-136", "AB AppDev: next 10",                                            "Higher risk",                  d(17), d(18)),
    ("Batch 14", "Phase 5", "Repos 137-146", "AB AppDev: next 10",                                            "Higher risk",                  d(17), d(18)),
    ("Batch 15", "Phase 5", "Repos 147-156", "AB AppDev: next 10",                                            "Higher risk",                  d(18), d(19)),
    ("Batch 16", "Phase 5", "Repos 157-166", "AB AppDev: next 10",                                            "Higher risk",                  d(18), d(19)),
    ("Batch 17", "Phase 5", "Repos 167-176", "AB AppDev: next 10",                                            "Higher risk",                  d(19), d(20)),
    ("Batch 18", "Phase 5", "Repos 177-186", "AB AppDev: next 10",                                            "Higher risk",                  d(19), d(20)),
    ("Batch 19", "Phase 5", "Repos 187-196", "AB AppDev: next 10",                                            "Higher risk",                  d(20), d(21)),
    ("Batch 20", "Phase 5", "Repos 197-206", "AB AppDev: next 10",                                            "Higher risk",                  d(20), d(21)),
    ("Batch 21", "Phase 5", "Repos 207-216", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 22", "Phase 5", "Repos 217-226", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 23", "Phase 5", "Repos 227-236", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 24", "Phase 5", "Repos 237-246", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 25", "Phase 5", "Repos 247-256", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 26", "Phase 5", "Repos 257-266", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 27", "Phase 5", "Repos 267-276", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 28", "Phase 5", "Repos 277-286", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 29", "Phase 5", "Repos 287-296", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 30", "Phase 5", "Repos 297-306", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 31", "Phase 5", "Repos 307-316", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 32", "Phase 5", "Repos 317-326", "AB AppDev: next 10",                                            "Higher risk",                  d(21), d(22)),
    ("Batch 33", "Phase 5", "Repo 327",      "AB AppDev: final repo",                                         "AB AppDev fully migrated",     d(21), d(22)),
]
write_sheet(
    ws_batch,
    ["Batch", "Phase", "Repo Range", "Contents", "Risk / Notes", "Start Date", "End Date"],
    batch_rows,
    phase_col_idx=1,
)

# ─────────────────────────────────────────────
# Sheet 4: Decisions Needed
# ─────────────────────────────────────────────
ws_dec = wb.create_sheet("Decisions Needed")
decisions = [
    (1, "master -> main rename",      "309 repos are on `master`. GitHub default is `main`. Do we rename during migration or keep `master`?",   "Project Lead + Abank Dev Teams", d(0),  "Rename to `main` during mirror step (recommended) OR keep `master` and set GitHub default branch to `master`"),
    (2, "18 empty repos",             "18 repos have no commits. Skip migration or create empty repos on GitHub to preserve the name?",          "Repo Owners",                    d(0),  "Skip (recommended) OR create empty shell repos on GitHub"),
    (3, "3 name collisions",          "ACH BAI File Generator, BAM, DW_STAGING_LOAD_LEGACY exist in multiple ADO projects. One must be renamed.", "Repo Owners",                   d(0),  "Suffix the less-active copy (e.g. BAM-standalone) OR prefix with project (e.g. ab-appdev-BAM)"),
    (4, "PR history",                 "git mirror does not carry over pull requests or PR comments. Accept loss or engage GitHub PS?",            "Project Lead + Abank Management",d(0),  "Accept loss + keep ADO read-only 90 days (recommended) OR engage GitHub PS for custom exporter (2-3 wks extra)"),
    (5, "325 unprotected repos",      "Only 2 repos had branch policies in ADO. Apply standard Rulesets to ALL repos on GitHub?",               "Project Lead",                   d(1),  "Apply default-protect Ruleset to all repos (recommended) OR migrate as-is with no protection"),
    (6, "GitHub org name",            "What is the GitHub org name to create under GHEC? Should match Abank branding.",                          "Abank IT Management",            d(0),  "Org name to be confirmed by Abank"),
    (7, "ADO archive retention period","How long to keep ADO Server repos as read-only archive after cutover before decommission?",             "Abank IT Management",            d(24), "90 days (recommended) OR 6 months OR permanent archive"),
]
write_sheet(
    ws_dec,
    ["S.No", "Decision", "Context", "Decision Owner", "Needed By", "Options"],
    decisions,
)

# ─────────────────────────────────────────────
# Sheet 5: Scripts to Develop
# ─────────────────────────────────────────────
ws_scripts = wb.create_sheet("Scripts to Develop")
scripts = [
    (1, "Phase 1", "git_mirror_migrate.py",    "Clone each ADO repo with --mirror and push to GitHub. Supports dry-run, retry, resume, branch rename (master->main), progress log. Reads repo list from Repo Inventory sheet.",      "Engineer", d(1), d(3), "Core migration engine"),
    (2, "Phase 1", "apply_topics_teams.py",    "Read Target Topics + Target Teams from inventory. Call gh api to set topics and team permissions per repo. Dry-run mode. Idempotent.",                                                "Engineer", d(1), d(3), "GitHub org hygiene"),
    (3, "Phase 1", "apply_rulesets.py",        "Author 3-4 Ruleset JSON templates. Apply to repos by topic pattern. Supports default-protect, require-review, file-size, comment-required tiers. Dry-run mode.",                    "Engineer", d(1), d(3), "Branch protection replacement"),
    (4, "Phase 1", "post_migration_verify.py", "For each repo: check branch count, tag count, default branch, topics applied, team access, Ruleset attached. Output pass/fail report per repo. Highlight gaps.",                    "Engineer", d(2), d(3), "Quality gate per batch"),
    (5, "Phase 2", "pilot_runner.sh",          "Bash wrapper that runs the full 4-script sequence on 2 pilot repos. Captures logs, outputs summary. Used during Phase 2 pilot only.",                                                "Engineer", d(4), d(5), "Pilot orchestration"),
]
write_sheet(
    ws_scripts,
    ["S.No", "Phase", "Script Name", "Purpose", "Owner", "Start Date", "End Date", "Role in Migration"],
    scripts,
    phase_col_idx=1,
)

# ─────────────────────────────────────────────
# Sheet 6: Per-Batch Cutover Process
# ─────────────────────────────────────────────
ws_proc = wb.create_sheet("Per-Batch Process")
process = [
    (1,  "T-48h",   "Freeze Notice",          "Send email/Teams message to repo owners: ADO push will be disabled in 48h. Share GitHub target URL."),
    (2,  "T-1h",    "Final Sync Check",        "Run git fetch --prune on ADO repos in batch. Note any last-minute commits."),
    (3,  "T-0",     "Freeze ADO",              "Disable push to ADO repos in batch (ADO: Repos > Settings > Disable pushes OR use branch lock). Log freeze time."),
    (4,  "T+0",     "Mirror Clone",            "Run git_mirror_migrate.py --dry-run first. Review output. Then run live. Log any failures."),
    (5,  "T+0",     "Apply Topics + Teams",    "Run apply_topics_teams.py --dry-run. Review. Run live."),
    (6,  "T+0",     "Apply Rulesets",          "Run apply_rulesets.py --dry-run. Review. Run live."),
    (7,  "T+1h",    "Verify",                  "Run post_migration_verify.py. All repos must be green before proceeding."),
    (8,  "T+1h",    "Owner Validation",        "Repo owner clones from GitHub, checks branch list, confirms code integrity. Sign-off."),
    (9,  "T+2h",    "Archive ADO",             "Set ADO repos to read-only (disable). Update Status column in Repo Inventory sheet to Migrated."),
    (10, "T+2h",    "Update Roadmap",          "Mark batch rows as Completed in this workbook. Commit + push updated workbook to GitHub repo."),
    (11, "Post",    "Consumer Redirect",       "Raise PRs or tickets to update any submodule URLs, CI clone URLs, package feed refs pointing to ADO."),
]
write_sheet(
    ws_proc,
    ["Step", "When", "Action", "Detail"],
    process,
)

wb.save(OUTPUT)
print(f"Workbook saved to: {OUTPUT}")
