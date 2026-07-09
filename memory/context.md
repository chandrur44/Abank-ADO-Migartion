# Project Context

## Summary

Abank engagement. Primary active workstream: **migration from Azure DevOps (cloud) to GitHub (cloud)** across an ADO org with 100+ repos spread over multiple projects. Two adjacent workstreams tied to the same migration:

1. Draft the **current branching policy** and design **automated CI/CD with validation checks** on the GitHub side.
2. Investigate the **Databricks team's request for a new UAT branch** — capture their use case and document the anti-patterns they DO NOT follow that are causing the current mess.

Source system: Azure DevOps (confirmed by Sabreesh Sakthivel, 2026-07-08).
Target system: GitHub Enterprise Cloud (edition TBD - standard GHEC vs. EMU decision pending).

## Change Log

- **2026-07-09** - **Major scope shift**: primary migration source is Abank on-prem ADO Server at `http://abdevopsdev/DefaultCollection/...`, NOT the cloud `dev.azure.com/zeb-ai`. Server is fully firewalled, no outbound to github.com. All extraction/migration scripts must run on the Abank VDI (no admin, embeddable Python). Zeb-ai work is retconned as warm-up/tooling rehearsal. Preferred path: GEI-for-ADO-Server (needs firewall exception from server or jump host to `api.github.com`). Fallback: `git mirror` (loses PRs/wiki). See [[project-abdevopsdev-server]].

- **2026-07-08** - Remote configured: `https://github.com/chandrur44/Abank-ADO-Migartion.git`. Branch renamed `master` -> `main`. All existing commits pushed. Standing rule: push after every local commit. `.env` verified gitignored (not pushed).
- **2026-07-08** - Live check via `scripts/check_wikis_and_tags.py`: **0 wikis across all 6 projects**, **1 tag total** (in `cloud-centralized-pipeline/cloud-centralized-pipeline`). Wikis dropped from migration scope; tags negligible but still preserved automatically by GEI.
- **2026-07-08** - Fixed extractor to resolve orphaned policy repo IDs. Unresolvable IDs now labeled `(deleted repo)` instead of `?` - these are ADO policies pointing at repos that were removed but the policy config was not cleaned up. Recommend ADO-side cleanup pre-migration.
- **2026-07-08** - Ran ADO inventory extractor against `https://dev.azure.com/zeb-ai`. Found **6 projects, 49 repos, 81 branch policies, 0 name collisions**. Original "100+ repos" estimate was high - actual is 49. Output: `Documents/ADO Inventory zeb-ai.xlsx` (Summary / Projects / Repo Inventory / Branch Policies / Collisions sheets). Projects: cloud-centralized-pipeline, ds-dataops, zeb-aws-clients-demo, zeb-databricks, zeb-spiral-poc, zeb-theteam-fka-wasserman-media-group-prod.
- **2026-07-08** - Scope narrowed to Repos + Branch Policies + Access + Topics. Confirmed standard GHEC, retain repo names, use topics for hierarchy. OSS tooling scanned. Delivered `Documents/ADO to GitHub Migration Plan.xlsx` (6 sheets, 30-task 6-phase plan).
- **2026-07-08** - Git initialized (`master` branch). `.gitignore` excludes creds file. Root commit `1fbe9ac` captures CLAUDE.md + memory bootstrap. Standing rule: commit after every response that changes files.
- **2026-07-08** - Memory discipline confirmed: read `context.md` + relevant topic file at task start, update at task end, keep `MEMORY.md` in sync. All CLAUDE.md conventions (Excel, Word, AWS CLI safety, tables-for-comparisons, no-loose-files) acknowledged as standing rules.
- **2026-07-08** - Engagement kicked off. Confirmed source is Azure DevOps (not GitHub Enterprise Server). Produced first deliverable: full ADO -> GitHub component inventory (20 components, migration approach per component, tool mapping to GEI / Actions Importer / manual rebuild). Flagged gaps: Universal Packages and Test Plans have no GitHub equivalent; EMU vs. standard GHEC decision needed early; recommend OIDC over lifted service-connection secrets. Memory folder bootstrapped (previously missing despite CLAUDE.md convention).
