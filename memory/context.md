# Project Context

## Summary

Abank engagement. Primary active workstream: **migration from Azure DevOps (cloud) to GitHub (cloud)** across an ADO org with 100+ repos spread over multiple projects. Two adjacent workstreams tied to the same migration:

1. Draft the **current branching policy** and design **automated CI/CD with validation checks** on the GitHub side.
2. Investigate the **Databricks team's request for a new UAT branch** — capture their use case and document the anti-patterns they DO NOT follow that are causing the current mess.

Source system: Azure DevOps (confirmed by Sabreesh Sakthivel, 2026-07-08).
Target system: GitHub Enterprise Cloud (edition TBD - standard GHEC vs. EMU decision pending).

## Change Log

- **2026-07-08** - Memory discipline confirmed: read `context.md` + relevant topic file at task start, update at task end, keep `MEMORY.md` in sync. All CLAUDE.md conventions (Excel, Word, AWS CLI safety, tables-for-comparisons, no-loose-files) acknowledged as standing rules.
- **2026-07-08** - Engagement kicked off. Confirmed source is Azure DevOps (not GitHub Enterprise Server). Produced first deliverable: full ADO -> GitHub component inventory (20 components, migration approach per component, tool mapping to GEI / Actions Importer / manual rebuild). Flagged gaps: Universal Packages and Test Plans have no GitHub equivalent; EMU vs. standard GHEC decision needed early; recommend OIDC over lifted service-connection secrets. Memory folder bootstrapped (previously missing despite CLAUDE.md convention).
