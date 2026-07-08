---
name: project-ado-github-migration
description: Azure DevOps to GitHub Cloud migration - scope, footprint, component inventory, and open decisions for the Abank engagement
metadata:
  type: project
---

# ADO -> GitHub Cloud Migration

## Fact / Decision

Migrating from **Azure DevOps (cloud)** to **GitHub (cloud)**. Source org: `https://dev.azure.com/zeb-ai`. Actual footprint (extracted 2026-07-08): **6 projects, 49 repos, 81 branch policies, 0 name collisions**. Original stakeholder estimate of "100+ repos" was high. Projects: cloud-centralized-pipeline, ds-dataops, zeb-aws-clients-demo, zeb-databricks, zeb-spiral-poc, zeb-theteam-fka-wasserman-media-group-prod. Confirmed 2026-07-08 by Sabreesh Sakthivel.

**Why:** Client directive to consolidate on GitHub. Scale (100+ repos, multi-project) means wave-based cutover, not big-bang.

**How to apply:** Any recommendation must account for (a) multi-project ADO structure mapping to GitHub org/teams, (b) wave-based rollout, (c) mixed pipeline estate (YAML + Classic + Release), (d) unknown Databricks-team practices that are contributing to current mess (see [[project-databricks-uat-branch]]).

## Component Inventory (delivered 2026-07-08)

20-component map covering: Repos, PRs, Branch Policies, YAML Pipelines, Classic Pipelines, Release Pipelines, Service Connections, Variable Groups, Secure Files, Agents/Pools, Artifacts Feeds, Work Items/Boards, Wikis, Test Plans, Dashboards, Extensions, Permissions/Groups, Entra SSO/SCIM, Audit Logs, Service Hooks.

Tooling per component: **GitHub Enterprise Importer (GEI)** for repos/PRs/wiki/work-items; **GitHub Actions Importer** for YAML pipelines (60-80% auto-convert); manual rebuild for Classic pipelines, service connections (recommend OIDC federation), variable groups, permissions, webhooks.

## Scope Narrowed (2026-07-08)

Active phase focused on: **Repos + Branch Policies + Access + Topics-for-hierarchy**. Pipelines / artifacts / work items / wikis deferred.

**Decisions confirmed:**
- GHEC edition: **Standard** (not EMU).
- Repo naming: **retain original ADO repo names**; use **GitHub Topics** to preserve ADO project hierarchy.
- Collision policy: when the same repo name exists in multiple ADO projects, suffix one (e.g. `shared-lib-projectb`). Discovery must flag every collision before migration.
- Every repo must have **exactly one `project-*` topic** - that is the hierarchy replacement in a flat GitHub org.

**Toolchain (all free / OSS):**
- **GitHub Enterprise Importer** (`gh ado2gh`) - repos/history/PRs
- **`gh` CLI + `gh api`** - topics, teams, permissions, Rulesets
- **ADO REST API + Python** - discovery inventory export
- **ghorg** - pre-migration bulk backup
- Optional: git-filter-repo / BFG (only if pre-migration history cleanup needed)

**Deliverable:** `Documents/ADO to GitHub Migration Plan.xlsx` with sheets: Summary, Migration Plan (30 tasks over 6 phases), Tooling, Repo Inventory (template), Topic Taxonomy, Access Model.

## Open Decisions / Gaps

- **GHEC edition**: standard GHEC vs. **Enterprise Managed Users (EMU)**. EMU gives IdP-managed identities but blocks OSS contribution from same identity. Pick early.
- **Universal Packages**: no GitHub equivalent. Replacement path required before cutover.
- **Test Plans**: no native GitHub equivalent. Options: standalone Azure Test Plans, TestRail, Xray, Zephyr.
- **Dashboards / Analytics**: no 1:1 replacement. Third-party (LinearB, Sleuth, Datadog) or accept reporting gap.
- **Work Items**: migrate to GitHub Issues/Projects v2, or keep in Jira/ServiceNow.
- **Self-hosted agents**: keep VMs / move to Actions Runner Controller on AKS / go GitHub-hosted.

## Standing Recommendations

- **OIDC over lifted secrets** for cloud service connections from day one.
- **Rulesets (JSON, version-controlled)** for branch protection, not click-ops branch protection rules.
- **Wave-based cutover**: freeze -> migrate -> validate -> redirect per wave.
- Pilot with 1-2 low-risk repos before committing GEI at scale.

Related: [[project-branching-cicd]], [[project-databricks-uat-branch]]
