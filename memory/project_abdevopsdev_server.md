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
