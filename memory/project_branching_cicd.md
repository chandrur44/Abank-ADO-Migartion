---
name: project-branching-cicd
description: Current ADO branching policy audit and target GitHub CI/CD with automated validation checks - pending workstream
metadata:
  type: project
---

# Branching Policy & Automated CI/CD (GitHub target)

## Fact / Decision

Requested workstream (Sabreesh, 2026-07-08): **draft the current branching policy** in use on ADO and **design automated CI/CD with validation checks** for the GitHub target.

**Why:** Migration is the forcing function to fix long-standing branching/CI-CD inconsistency across 100+ repos. Databricks team's misuse (see [[project-databricks-uat-branch]]) is one symptom.

**How to apply:** Any branching model recommendation must be enforceable via **GitHub Rulesets + required status checks + CODEOWNERS**, not policy-by-convention. Assume repos have inconsistent practice today - do not over-index on any single repo's setup.

## Status

Not yet started. Blocked on: current-state discovery (need a sample of representative repos + their existing branch policies from ADO).

## Related

- [[project-ado-github-migration]] - migration parent
- [[project-databricks-uat-branch]] - concrete anti-pattern example
