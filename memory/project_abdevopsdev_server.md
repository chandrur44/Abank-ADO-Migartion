---
name: project-abdevopsdev-server
description: Abank on-prem Azure DevOps Server (abdevopsdev) - the PRIMARY migration source, replacing the assumed cloud-only scope
metadata:
  type: project
---

# Abank On-Prem ADO Server (`abdevopsdev`) - Primary Migration Source

## Fact / Decision

Confirmed 2026-07-09: The primary migration source is **Azure DevOps Server on-prem** at `http://abdevopsdev/DefaultCollection/...`, not the cloud org [[project-ado-github-migration]] (which was a warm-up). Server is **fully firewalled**, cannot reach github.com; only path to internet is via the Abank VDI where the user works.

Example repo URL structure:
`http://abdevopsdev/DefaultCollection/AB%20AppDev/_git/ABAPPUTILS`
Server = `abdevopsdev` | Collection = `DefaultCollection` | Project = `AB AppDev` | Repo = `ABAPPUTILS`

**Why:** This is the actual Abank engagement. The zeb-ai work built tooling and process; the real payload is here.

**How to apply:** All future migration work targets `abdevopsdev`. My laptop CANNOT reach it - scripts must run on the VDI, outputs pushed back via the GitHub repo. Any deliverable must account for VDI-only execution.

## Constraints

- **Network:** on-prem server, HTTP (not HTTPS), no outbound to github.com from the server.
- **VDI:** no admin rights - Python via embeddable zip (see conversation for setup steps).
- **API version:** depends on ADO Server version installed - must confirm before writing REST calls (e.g. `api-version=6.0` vs `7.1`).
- **Auth:** PAT (Code Read + Project Read + Policy Read scopes), same shape as cloud.

## Actual Inventory (extracted 2026-07-09 from `Abank Document/ADO Inventory DefaultCollection.xlsx`)

| Metric | Value |
|---|---|
| Projects | 24 |
| Total repos | 327 (all active, 0 disabled) |
| Total size | 1,917 MB |
| Repos with branch policies | 2 (325 unprotected) |
| Name collisions | 3 (ACH BAI File Generator, BAM, DW_STAGING_LOAD_LEGACY) |
| Empty repos (no default branch / no commits) | 18 |
| Repos on `master` | 309 |

**Project breakdown (notable):** AB AppDev (244 repos, active Jun 2026), ScriptRefactoring (56 repos, active Dec 2025), SSIS ETL (5), 20 single-repo projects mostly last active 2020-2024.

**GEI confirmed NOT supported for on-prem ADO Server** (official docs: only ADO Services cloud, Bitbucket Server, GitHub.com, GHES are supported sources). Migration path: `git push --mirror` from VDI as jump host.

## Pending Decisions

- **`master` -> `main` rename**: 309 repos on `master`. Rename during migration or accept `master` as default on GitHub. Needs stakeholder decision before wave 1.
- **18 empty repos**: no commits ever pushed. Confirm skip or migrate as empty repos with owners.
- **3 collisions**: ACH BAI File Generator / BAM / DW_STAGING_LOAD_LEGACY. Need owner sign-off on which gets renamed.
- **325 unprotected repos**: design Ruleset templates to enforce standards on GitHub that were never in ADO. Linked to [[project-branching-cicd]].
- **PR history**: git mirror loses PRs. Decision: accept loss (freeze ADO 90 days as reference) or engage GitHub PS for custom exporter.

## Open Decisions

- **GEI-for-ADO-Server vs. `git mirror` fallback**: preferred path is to get Abank infra to allow outbound HTTPS from `abdevopsdev` (or a jump host) to `api.github.com` + `github.com`. That unlocks GEI + full PR/wiki migration. If denied, fallback = `git clone --mirror` from VDI + `git push --mirror` to GitHub, but **loses PR history, PR comments, wiki** unless a custom exporter is built (2-3 weeks work).
- ADO Server version - confirm >= 2019 for GEI compatibility.
- Full inventory - not yet extracted (no reachability from my machine).

## Standing Adjustments

- Extractor script must be **runnable on VDI-embeddable Python** with only `requests`, `python-dotenv`, `openpyxl` (already the case).
- Any Excel deliverable produced on VDI should be committed to `Documents/` and pushed to GitHub so it appears back in my working copy.

## Related

- [[project-ado-github-migration]] - the zeb-ai cloud work; now demoted to warm-up / process rehearsal.
- [[project-branching-cicd]]
- [[project-databricks-uat-branch]]
