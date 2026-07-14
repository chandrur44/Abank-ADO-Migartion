"""Generate ADO -> GitHub Migration Roadmap Excel workbook.

Output: Abank Document/ADO to GitHub Migration Roadmap Abank.xlsx
Excel Output Standard per CLAUDE.md.
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ABANK_DOC_DIR = PROJECT_ROOT / "Abank Document"
ABANK_DOC_DIR.mkdir(exist_ok=True)
OUTPUT = ABANK_DOC_DIR / "ADO to GitHub Migration Roadmap Abank.xlsx"

NAVY = "002060"
WHITE = "FFFFFF"
EMPTY_FILL_COLOR = "FFD7D7"   # soft red - empty / unused repos
STALE_FILL_COLOR = "FFF2CC"   # yellow - stale projects

thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
BODY_FONT = Font(name="Calibri", size=11)
BODY_BOLD = Font(name="Calibri", size=11, bold=True)

HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

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
        phase_fill = PHASE_FILLS.get(str(phase_val)) if phase_val else None
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
                max_len = max(max_len, len(line))
        ws.column_dimensions[letter].width = min(max_len + 2, 60)


BASE = date(2026, 7, 14)


def d(offset_days):
    return (BASE + timedelta(days=offset_days)).strftime("%d %b %Y")


wb = Workbook()

# ── Sheet 1: Summary ──────────────────────────
ws_sum = wb.active
ws_sum.title = "Summary"
write_sheet(ws_sum, ["Item", "Detail"], [
    ("Client",                    "Amalgamated Bank (Abank)"),
    ("Source",                    "Azure DevOps Server on-prem (abdevopsdev, API 5.1)"),
    ("Target",                    "GitHub Enterprise Cloud (Standard)"),
    ("Total Repos",               "327"),
    ("Active Repos (to migrate)", "309 (excl. 18 empty repos - skip confirmed)"),
    ("ADO Projects",              "24"),
    ("Migration Approach",        "git push --mirror from VDI (GEI not supported for on-prem ADO Server)"),
    ("Batch Size",                "10 repos per batch"),
    ("Total Batches",             "31 (309 active repos / 10, rounded up)"),
    ("Total Phases",              "8 (Phase 0 to Phase 7)"),
    ("Planned Start",             d(0)),
    ("Planned End",               "31 Aug 2026"),
    ("Total Duration",            "~7 weeks"),
])

# ── Sheet 2: Roadmap ──────────────────────────
ws_road = wb.create_sheet("Roadmap")
roadmap = [
    # S.No, Phase, Activity, Description, Owner, Start, End, Status, Expected Output
    # Phase 0 - Setup (days 0-7, 14-21 Jul)
    (1,  "Phase 0", "Kickoff & Governance",       "Align stakeholders on scope, git-mirror approach, timeline. Sign-off: master->main rename, 18 empty repos skip, 3 collision renames, PR history acceptance.",           "Project Lead",    d(0),  d(3),  "In Progress", "Scope sign-off, decision log"),
    (2,  "Phase 0", "GitHub Org Setup",            "Provision GHEC org. Configure Entra SSO + SCIM. Set org defaults: main branch, secret scanning, push protection.",                                                    "Platform Team",   d(0),  d(5),  "Not Started", "GitHub org live, SSO working"),
    (3,  "Phase 0", "VDI Tooling Setup",           "Embeddable Python on VDI, dependencies installed, migration repo cloned, .env configured.",                                                                            "Engineer",        d(0),  d(2),  "Done",        "Scripts runnable on VDI"),
    (4,  "Phase 0", "Inventory Finalization",      "Confirm 309 active repos, batch groupings, collision rename decisions, empty repo disposition.",                                                                        "Engineer + Lead", d(0),  d(5),  "Not Started", "Finalized Repo Inventory with batch assignments"),
    # Phase 1 - Script Dev (days 7-14, 21-28 Jul)
    (5,  "Phase 1", "Script Dev - Mirror",         "git_mirror_migrate.py: clone --mirror from ADO, push --mirror to GitHub. Dry-run, retry, resume, master->main rename flag.",                                         "Engineer",        d(7),  d(11), "Not Started", "git_mirror_migrate.py"),
    (6,  "Phase 1", "Script Dev - Topics & Teams", "apply_topics_teams.py: set GitHub topics + team permissions per repo from workbook. Dry-run, idempotent.",                                                            "Engineer",        d(7),  d(11), "Not Started", "apply_topics_teams.py"),
    (7,  "Phase 1", "Script Dev - Rulesets",       "apply_rulesets.py: 3-4 Ruleset JSON templates (default-protect, require-review, file-size, comment-required). Attach by topic.",                                     "Engineer",        d(7),  d(11), "Not Started", "apply_rulesets.py + JSON templates"),
    (8,  "Phase 1", "Script Dev - Verify",         "post_migration_verify.py: per-repo pass/fail on branch count, default branch, topics, team access, Ruleset.",                                                        "Engineer",        d(9),  d(12), "Not Started", "post_migration_verify.py"),
    (9,  "Phase 1", "Script Dev - CI/CD Template", "Develop standard GitHub Actions workflow template for identified repos. Single reusable YAML, auto-applied per repo via script.",                                     "Engineer",        d(9),  d(12), "Not Started", "cicd_template.yml + apply_cicd.py"),
    (10, "Phase 1", "Script Dev - Self-Hosted Runner", "Configure org-level self-hosted runner on Abank infra. Register runner with GHEC org. Validate runner picks up workflow jobs.",                                  "Engineer + Infra",d(7),  d(12), "Not Started", "Runner registered, online in GHEC org"),
    (11, "Phase 1", "Unit Testing - All Scripts",  "Test all scripts against zeb-ai sandbox. Validate dry-run, retry, resume, error handling.",                                                                            "Engineer",        d(12), d(14), "Not Started", "All scripts passing on sandbox"),
    # Phase 2 - Pilot (days 14-17, 28-31 Jul)
    (12, "Phase 2", "Pilot - 2 repos",             "1 ScriptRefactoring + 1 stale single-repo. End-to-end: mirror -> topics -> Ruleset -> CI/CD -> verify. Owner validates.",                                            "Engineer",        d(14), d(15), "Not Started", "2 repos on GitHub, verify green"),
    (13, "Phase 2", "Pilot Retro & Fix",           "Review pilot gaps. Fix scripts, update runbook. Get sign-off to scale.",                                                                                               "Engineer",        d(15), d(17), "Not Started", "Updated scripts + runbook, sign-off"),
    # Phase 3 - Stale singles (days 17-22, 31 Jul - 05 Aug)
    (14, "Phase 3", "Batch 1 (Repos 1-10)",        "Stale single-repo projects (>2 yrs idle). 24h freeze -> mirror -> topics + Ruleset -> verify -> archive.",                                                            "Engineer",        d(17), d(18), "Not Started", "10 repos on GitHub, ADO archived"),
    (15, "Phase 3", "Batch 2 (Repos 11-20)",       "Single-repo projects (1-2 yrs idle). Same process.",                                                                                                                  "Engineer",        d(18), d(19), "Not Started", "10 repos on GitHub, ADO archived"),
    (16, "Phase 3", "Batch 3 (Repos 21-30)",       "SSIS ETL (5) + remaining single-repo projects.",                                                                                                                      "Engineer",        d(19), d(20), "Not Started", "10 repos on GitHub, ADO archived"),
    (17, "Phase 3", "Batch 4 (Repos 31-40)",       "Non Binary (2) + active single-repo projects. 24h freeze notice.",                                                                                                    "Engineer",        d(20), d(21), "Not Started", "10 repos on GitHub, ADO archived"),
    (18, "Phase 3", "Batch 5 (Repos 41-50)",       "Remaining active single-repo projects.",                                                                                                                               "Engineer",        d(21), d(22), "Not Started", "10 repos on GitHub, ADO archived"),
    # Phase 4 - ScriptRefactoring (days 24-31, 07-14 Aug)
    (19, "Phase 4", "Batch 6 (Repos 51-60)",       "ScriptRefactoring: first 10 repos. 24h freeze notice.",                                                                                                               "Engineer",        d(24), d(25), "Not Started", "10 SR repos on GitHub"),
    (20, "Phase 4", "Batch 7 (Repos 61-70)",       "ScriptRefactoring: next 10.",                                                                                                                                         "Engineer",        d(25), d(26), "Not Started", "10 SR repos on GitHub"),
    (21, "Phase 4", "Batch 8 (Repos 71-80)",       "ScriptRefactoring: next 10.",                                                                                                                                         "Engineer",        d(26), d(27), "Not Started", "10 SR repos on GitHub"),
    (22, "Phase 4", "Batch 9 (Repos 81-90)",       "ScriptRefactoring: next 10.",                                                                                                                                         "Engineer",        d(27), d(28), "Not Started", "10 SR repos on GitHub"),
    (23, "Phase 4", "Batch 10 (Repos 91-106)",     "ScriptRefactoring: final 16. SR fully migrated.",                                                                                                                     "Engineer",        d(28), d(31), "Not Started", "All 56 SR repos on GitHub"),
    # Phase 5 - AB AppDev (days 31-47, 14-30 Aug)
    (24, "Phase 5", "Batch 11 (Repos 107-116)",    "AB AppDev: first 10 (oldest last-commit). 24h freeze notice.",                                                                                                        "Engineer",        d(31), d(32), "Not Started", "10 AB AppDev repos on GitHub"),
    (25, "Phase 5", "Batch 12 (Repos 117-126)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(32), d(33), "Not Started", "10 AB AppDev repos on GitHub"),
    (26, "Phase 5", "Batch 13 (Repos 127-136)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(33), d(34), "Not Started", "10 AB AppDev repos on GitHub"),
    (27, "Phase 5", "Batch 14 (Repos 137-146)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(34), d(35), "Not Started", "10 AB AppDev repos on GitHub"),
    (28, "Phase 5", "Batch 15 (Repos 147-156)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(35), d(36), "Not Started", "10 AB AppDev repos on GitHub"),
    (29, "Phase 5", "Batch 16 (Repos 157-166)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(36), d(37), "Not Started", "10 AB AppDev repos on GitHub"),
    (30, "Phase 5", "Batch 17 (Repos 167-176)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(37), d(38), "Not Started", "10 AB AppDev repos on GitHub"),
    (31, "Phase 5", "Batch 18 (Repos 177-186)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(38), d(39), "Not Started", "10 AB AppDev repos on GitHub"),
    (32, "Phase 5", "Batch 19 (Repos 187-196)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(39), d(40), "Not Started", "10 AB AppDev repos on GitHub"),
    (33, "Phase 5", "Batch 20 (Repos 197-206)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(40), d(41), "Not Started", "10 AB AppDev repos on GitHub"),
    (34, "Phase 5", "Batch 21 (Repos 207-216)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(41), d(42), "Not Started", "10 AB AppDev repos on GitHub"),
    (35, "Phase 5", "Batch 22 (Repos 217-226)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(42), d(43), "Not Started", "10 AB AppDev repos on GitHub"),
    (36, "Phase 5", "Batch 23 (Repos 227-236)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(43), d(44), "Not Started", "10 AB AppDev repos on GitHub"),
    (37, "Phase 5", "Batch 24 (Repos 237-246)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(43), d(44), "Not Started", "10 AB AppDev repos on GitHub"),
    (38, "Phase 5", "Batch 25 (Repos 247-256)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(44), d(45), "Not Started", "10 AB AppDev repos on GitHub"),
    (39, "Phase 5", "Batch 26 (Repos 257-266)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(44), d(45), "Not Started", "10 AB AppDev repos on GitHub"),
    (40, "Phase 5", "Batch 27 (Repos 267-276)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(45), d(46), "Not Started", "10 AB AppDev repos on GitHub"),
    (41, "Phase 5", "Batch 28 (Repos 277-286)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(45), d(46), "Not Started", "10 AB AppDev repos on GitHub"),
    (42, "Phase 5", "Batch 29 (Repos 287-296)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(46), d(47), "Not Started", "10 AB AppDev repos on GitHub"),
    (43, "Phase 5", "Batch 30 (Repos 297-306)",    "AB AppDev: next 10.",                                                                                                                                                 "Engineer",        d(46), d(47), "Not Started", "10 AB AppDev repos on GitHub"),
    (44, "Phase 5", "Batch 31 (Repos 307-309)",    "AB AppDev: final 3 repos. All 309 active repos migrated.",                                                                                                            "Engineer",        d(47), d(47), "Not Started", "All AB AppDev repos on GitHub"),
    # Phase 6 - Audit (days 44-47, 27-30 Aug) - runs in parallel with tail of Phase 5
    (45, "Phase 6", "CI/CD Identification",        "Identify repos requiring CI/CD pipeline. Document pipeline type and trigger pattern per repo.",                                                                        "Engineer + Lead", d(44), d(46), "Not Started", "CI/CD candidate list"),
    (46, "Phase 6", "CI/CD Pipeline Rollout",      "Apply standard GitHub Actions workflow YAML to identified repos via apply_cicd.py. Validate each pipeline runs on self-hosted runner.",                               "Engineer",        d(44), d(47), "Not Started", "Pipelines active, runner jobs green"),
    (47, "Phase 6", "Full Org Audit",              "post_migration_verify.py across all 309 repos. Every repo: project-* topic, correct default branch, Ruleset, team access.",                                           "Engineer",        d(44), d(46), "Not Started", "Audit report - all green"),
    (48, "Phase 6", "Security Review",             "No public repos unintentionally, secret scanning on org-wide, push protection on, no hardcoded secrets flagged.",                                                     "Security",        d(44), d(46), "Not Started", "Security sign-off"),
    (49, "Phase 6", "Consumer Redirect",           "Update submodule URLs, CI clone URLs, README badges pointing to ADO. PRs raised per affected repo.",                                                                   "Team Leads",      d(44), d(47), "Not Started", "All consumers on GitHub URLs"),
    # Phase 7 - Handover (days 47-48, 30-31 Aug)
    (50, "Phase 7", "ADO Server Archive",          "Set all ADO repos to read-only. Communicate: ADO is archive-only. Retain 90 days.",                                                                                   "Platform Team",   d(47), d(48), "Not Started", "ADO read-only, comms sent"),
    (51, "Phase 7", "Runbook Handover",            "Runbook to platform team: add repo, apply topics, request access, change Ruleset, manage runner. 1h session.",                                                        "Engineer",        d(47), d(48), "Not Started", "Runbook + handover session done"),
    (52, "Phase 7", "ADO Decommission Planning",   "Plan ADO Server decommission 90 days post-archive. Confirm retention policy. Schedule infra team.",                                                                   "Abank Infra",     d(48), d(48), "Not Started", "Decommission plan signed off"),
]
write_sheet(
    ws_road,
    ["S.No", "Phase", "Activity", "Description", "Owner", "Start Date", "End Date", "Status", "Expected Output"],
    roadmap,
    phase_col_idx=1,
)

# ── Sheet 3: Batch Plan ───────────────────────
ws_batch = wb.create_sheet("Batch Plan")
write_sheet(
    ws_batch,
    ["Batch", "Phase", "Repo Range", "Contents", "Risk", "Start Date", "End Date"],
    [
        ("Pilot",    "Phase 2", "2 repos",       "1 ScriptRefactoring + 1 stale single-repo",    "Validation only",   d(14), d(17)),
        ("Batch 1",  "Phase 3", "Repos 1-10",    "Stale single-repo (>2 yrs idle)",              "Low",               d(17), d(18)),
        ("Batch 2",  "Phase 3", "Repos 11-20",   "Single-repo (1-2 yrs idle)",                   "Low",               d(18), d(19)),
        ("Batch 3",  "Phase 3", "Repos 21-30",   "SSIS ETL (5) + single-repo projects",          "Low-medium",        d(19), d(20)),
        ("Batch 4",  "Phase 3", "Repos 31-40",   "Non Binary (2) + active single-repo",          "Medium",            d(20), d(21)),
        ("Batch 5",  "Phase 3", "Repos 41-50",   "Remaining active single-repo",                 "Medium",            d(21), d(22)),
        ("Batch 6",  "Phase 4", "Repos 51-60",   "ScriptRefactoring: first 10",                  "Medium",            d(24), d(25)),
        ("Batch 7",  "Phase 4", "Repos 61-70",   "ScriptRefactoring: next 10",                   "Medium",            d(25), d(26)),
        ("Batch 8",  "Phase 4", "Repos 71-80",   "ScriptRefactoring: next 10",                   "Medium",            d(26), d(27)),
        ("Batch 9",  "Phase 4", "Repos 81-90",   "ScriptRefactoring: next 10",                   "Medium",            d(27), d(28)),
        ("Batch 10", "Phase 4", "Repos 91-106",  "ScriptRefactoring: final 16",                  "SR complete",       d(28), d(31)),
        ("Batch 11", "Phase 5", "Repos 107-116", "AB AppDev: first 10 (oldest)",                 "Higher",            d(31), d(32)),
        ("Batch 12", "Phase 5", "Repos 117-126", "AB AppDev: next 10",                           "Higher",            d(32), d(33)),
        ("Batch 13", "Phase 5", "Repos 127-136", "AB AppDev: next 10",                           "Higher",            d(33), d(34)),
        ("Batch 14", "Phase 5", "Repos 137-146", "AB AppDev: next 10",                           "Higher",            d(34), d(35)),
        ("Batch 15", "Phase 5", "Repos 147-156", "AB AppDev: next 10",                           "Higher",            d(35), d(36)),
        ("Batch 16", "Phase 5", "Repos 157-166", "AB AppDev: next 10",                           "Higher",            d(36), d(37)),
        ("Batch 17", "Phase 5", "Repos 167-176", "AB AppDev: next 10",                           "Higher",            d(37), d(38)),
        ("Batch 18", "Phase 5", "Repos 177-186", "AB AppDev: next 10",                           "Higher",            d(38), d(39)),
        ("Batch 19", "Phase 5", "Repos 187-196", "AB AppDev: next 10",                           "Higher",            d(39), d(40)),
        ("Batch 20", "Phase 5", "Repos 197-206", "AB AppDev: next 10",                           "Higher",            d(40), d(41)),
        ("Batch 21", "Phase 5", "Repos 207-216", "AB AppDev: next 10",                           "Higher",            d(41), d(42)),
        ("Batch 22", "Phase 5", "Repos 217-226", "AB AppDev: next 10",                           "Higher",            d(42), d(43)),
        ("Batch 23", "Phase 5", "Repos 227-236", "AB AppDev: next 10",                           "Higher",            d(43), d(44)),
        ("Batch 24", "Phase 5", "Repos 237-246", "AB AppDev: next 10",                           "Higher",            d(43), d(44)),
        ("Batch 25", "Phase 5", "Repos 247-256", "AB AppDev: next 10",                           "Higher",            d(44), d(45)),
        ("Batch 26", "Phase 5", "Repos 257-266", "AB AppDev: next 10",                           "Higher",            d(44), d(45)),
        ("Batch 27", "Phase 5", "Repos 267-276", "AB AppDev: next 10",                           "Higher",            d(45), d(46)),
        ("Batch 28", "Phase 5", "Repos 277-286", "AB AppDev: next 10",                           "Higher",            d(45), d(46)),
        ("Batch 29", "Phase 5", "Repos 287-296", "AB AppDev: next 10",                           "Higher",            d(46), d(47)),
        ("Batch 30", "Phase 5", "Repos 297-306", "AB AppDev: next 10",                           "Higher",            d(46), d(47)),
        ("Batch 31", "Phase 5", "Repos 307-309", "AB AppDev: final 3",                           "All 309 migrated",  d(47), d(47)),
    ],
    phase_col_idx=1,
)

# ── Sheet 4: Decisions Needed (no Needed By col) ──
ws_dec = wb.create_sheet("Decisions Needed")
write_sheet(
    ws_dec,
    ["S.No", "Decision", "Context", "Decision Owner", "Options"],
    [
        (1, "master -> main rename",        "309 repos on `master`. Rename during migration or keep `master`?",                                                  "Project Lead + Dev Teams",  "Rename to `main` during mirror step (recommended) OR keep `master`"),
        (2, "18 empty repos",               "18 repos have no commits. Skip or create empty shell repos on GitHub?",                                             "Repo Owners",               "Skip (recommended) OR create empty shell repos"),
        (3, "3 name collisions",            "ACH BAI File Generator, BAM, DW_STAGING_LOAD_LEGACY exist in multiple projects. One of each must be renamed.",     "Repo Owners",               "Suffix the less-active copy (e.g. BAM-standalone)"),
        (4, "PR history",                   "git mirror does not carry over PRs or comments. Accept loss or engage GitHub PS?",                                  "Project Lead + Management", "Accept loss + keep ADO read-only 90 days (recommended) OR GitHub PS custom exporter"),
        (5, "325 unprotected repos",        "Only 2 repos had branch policies in ADO. Apply Rulesets to all repos on GitHub?",                                  "Project Lead",              "Apply default-protect Ruleset org-wide (recommended) OR migrate as-is"),
        (6, "GitHub org name",              "What is the GHEC org name? Should match Abank branding.",                                                           "Abank IT Management",       "To be confirmed by Abank"),
        (7, "CI/CD - repos in scope",       "Which repos require a CI/CD pipeline? Identification needed before Phase 6.",                                       "Project Lead + Dev Leads",  "Identify by tech stack (dotnet/python/node) + deployment target"),
        (8, "Self-hosted runner infra",     "What server/VM hosts the org-level self-hosted runner? Who maintains it?",                                          "Abank Infra",               "Dedicated VM (recommended) OR use existing build server"),
        (9, "ADO archive retention period", "How long to keep ADO repos read-only after cutover before decommission?",                                           "Abank IT Management",       "90 days (recommended) OR 6 months OR permanent"),
    ],
)

# ── Sheet 5: Scripts to Develop (no Start/End date) ──
ws_scripts = wb.create_sheet("Scripts to Develop")
write_sheet(
    ws_scripts,
    ["S.No", "Phase", "Script Name", "Purpose", "Owner", "Role in Migration"],
    [
        (1, "Phase 1", "git_mirror_migrate.py",    "Clone --mirror from ADO, push --mirror to GitHub. Dry-run, retry, resume, master->main rename flag.",         "Engineer", "Core migration engine"),
        (2, "Phase 1", "apply_topics_teams.py",    "Set GitHub topics + team permissions per repo from workbook. Dry-run, idempotent.",                           "Engineer", "Org hierarchy + access"),
        (3, "Phase 1", "apply_rulesets.py",        "Create 3-4 Ruleset templates, attach to repos by topic. Dry-run.",                                            "Engineer", "Branch protection"),
        (4, "Phase 1", "post_migration_verify.py", "Per-repo pass/fail: branch count, default branch, topics, team access, Ruleset.",                            "Engineer", "Quality gate per batch"),
        (5, "Phase 1", "apply_cicd.py",            "Apply standard GitHub Actions YAML to identified repos. Validates pipeline triggers on self-hosted runner.",  "Engineer", "CI/CD rollout"),
        (6, "Phase 2", "pilot_runner.sh",          "Orchestrates full 5-script sequence on 2 pilot repos. Captures logs, outputs summary.",                       "Engineer", "Pilot orchestration"),
    ],
    phase_col_idx=1,
)

# ── Sheet 6: Per-Batch Cutover Process ────────
ws_proc = wb.create_sheet("Per-Batch Process")
write_sheet(
    ws_proc,
    ["Step", "When", "Action", "Detail"],
    [
        (1,  "T-24h",  "Freeze Notice",       "Notify repo owners via Teams/email: ADO push disabled in 24h. Share GitHub target URL."),
        (2,  "T-1h",   "Final Sync Check",    "git fetch --prune on ADO repos in batch. Note any last-minute commits."),
        (3,  "T-0",    "Freeze ADO",          "Disable push on ADO repos (Repos > Settings > Disable). Log freeze time."),
        (4,  "T+0",    "Mirror Clone",        "Run git_mirror_migrate.py --dry-run. Review. Run live. Log failures."),
        (5,  "T+0",    "Apply Topics+Teams",  "Run apply_topics_teams.py --dry-run. Review. Run live."),
        (6,  "T+0",    "Apply Rulesets",      "Run apply_rulesets.py --dry-run. Review. Run live."),
        (7,  "T+0",    "Apply CI/CD",         "Run apply_cicd.py on repos in batch flagged for CI/CD. Confirm pipeline queued on runner."),
        (8,  "T+1h",   "Verify",              "Run post_migration_verify.py. All repos must pass before proceeding."),
        (9,  "T+1h",   "Owner Validation",    "Repo owner clones from GitHub, confirms branches and code. Signs off."),
        (10, "T+2h",   "Archive ADO",         "Set ADO repos to read-only. Update Status = Migrated in Repo Inventory."),
        (11, "T+2h",   "Update Roadmap",      "Mark batch Completed in this workbook. Commit + push to GitHub."),
        (12, "Post",   "Consumer Redirect",   "PRs to update submodule URLs, CI clone URLs, README badges pointing to ADO."),
    ],
)

wb.save(OUTPUT)
print(f"Workbook saved to: {OUTPUT}")
