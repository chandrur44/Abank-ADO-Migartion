"""Quick check: list wikis (per project) and tag counts (per repo) in the ADO org."""
from __future__ import annotations

import base64
import os
import sys
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")
ORG_URL = os.environ["ADO_ORG_URL"].rstrip("/")
PAT = os.environ["ADO_PAT"]
S = requests.Session()
S.headers.update({"Authorization": "Basic " + base64.b64encode(f":{PAT}".encode()).decode()})


def api(url, **params):
    params.setdefault("api-version", "7.1")
    r = S.get(url, params=params, timeout=60)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


projects = api(f"{ORG_URL}/_apis/projects", **{"$top": 500}).get("value", [])

print("\n=== WIKIS PER PROJECT ===")
wiki_total = 0
for p in projects:
    wikis = api(f"{ORG_URL}/{p['id']}/_apis/wiki/wikis")
    entries = (wikis or {}).get("value", [])
    wiki_total += len(entries)
    print(f"  {p['name']}: {len(entries)} wiki(s)")
    for w in entries:
        print(f"    - {w.get('name')} ({w.get('type')})  id={w.get('id')}")
print(f"  TOTAL WIKIS: {wiki_total}")

print("\n=== TAGS PER REPO (non-zero only) ===")
tag_total = 0
repo_with_tags = 0
for p in projects:
    repos = api(f"{ORG_URL}/{p['id']}/_apis/git/repositories").get("value", [])
    for r in repos:
        tags = api(
            f"{ORG_URL}/{p['id']}/_apis/git/repositories/{r['id']}/refs",
            **{"filter": "tags/"},
        )
        count = len((tags or {}).get("value", []))
        if count:
            print(f"  {p['name']}/{r['name']}: {count} tag(s)")
            tag_total += count
            repo_with_tags += 1
print(f"  TOTAL TAGS: {tag_total} across {repo_with_tags} repo(s)")
