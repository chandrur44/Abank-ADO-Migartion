---
name: project-databricks-uat-branch
description: Databricks team requested a new UAT branch - investigate use case and document anti-patterns they DO NOT follow that caused the current mess
metadata:
  type: project
---

# Databricks UAT Branch Investigation

## Fact / Decision

Databricks team has requested a **new UAT branch**. Sabreesh's instruction (2026-07-08): understand **the use case** driving the request, AND capture **best practices they DO NOT follow** which have caused the current mess.

**Why:** This is not just a branch-creation ticket. The request is a symptom - approving it without fixing the underlying process would entrench the mess further. Findings will feed [[project-branching-cicd]].

**How to apply:** Before recommending yes/no on the new UAT branch, run a short discovery with the Databricks team:
- What environments do they promote through today? (dev -> ? -> prod)
- Why does the existing branch model not serve UAT?
- How are Databricks notebooks / job definitions / DLT pipelines version-controlled today?
- Who has direct-commit rights to shared branches? Are PRs used?
- How are workspace-linked repos configured (Databricks Repos feature)?
- How do they promote notebooks between workspaces (dev/uat/prod)?

## Status

Not started. Interview questionnaire pending.

## Related

- [[project-ado-github-migration]]
- [[project-branching-cicd]]
