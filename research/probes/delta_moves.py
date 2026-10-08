"""S2: what does Graph delta report when YOU move or delete mail in Outlook?

Read-only (profile 'graph'). Two phases; the mailbox change is made by you, by hand:

    python research/probes/delta_moves.py --snapshot
    # In Outlook: move one Inbox message to Archive, delete one Inbox message
    #             (to Deleted Items), and mark one Inbox message read/unread.
    python research/probes/delta_moves.py --check

--snapshot runs message delta to completion for inbox, archive and deleteditems and
saves the deltaLinks plus a salted-hash -> folder map to .local/probe-delta-state.json
(plaintext, git-ignored; delta links are opaque cursors, not credentials).
--check replays each deltaLink and reports, per folder: changed vs @removed records
(with removal reason), and for each removed id whether GET by immutable id still
works and where the message now lives. Counts and aliases only.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import secrets
from collections import Counter

from common import LOCAL, emit, graph, graph_pages, need, rows, well_known_ids

STATE = LOCAL / "probe-delta-state.json"
FOLDERS = ("inbox", "archive", "deleteditems")
SELECT = "id,parentFolderId,isRead,changeKey"


def stable_hash(salt: str, value: str) -> str:
    return hashlib.sha256((salt + value).encode()).hexdigest()[:16]


def run_delta(token: str, url_or_path: str) -> tuple[list[dict], str | None, int]:
    records, delta_link, pages = [], None, 0
    kwargs = {} if url_or_path.startswith("https://") else {"$select": SELECT}
    for r in graph_pages(token, url_or_path, max_pages=200,
                         headers={"Prefer": 'IdType="ImmutableId", odata.maxpagesize=500'}, prefer=None, **kwargs):
        pages += 1
        records.extend(rows(r))
        if isinstance(r.payload, dict) and r.payload.get("@odata.deltaLink"):
            delta_link = r.payload["@odata.deltaLink"]
    return records, delta_link, pages


def snapshot(token: str) -> dict:
    salt = secrets.token_hex(8)
    state = {"salt": salt, "folders": {}}
    out = {}
    for alias in FOLDERS:
        recs, link, pages = run_delta(token, f"/me/mailFolders/{alias}/messages/delta")
        state["folders"][alias] = {"delta_link": link,
                                   "ids": {stable_hash(salt, r["id"]): alias for r in recs if "id" in r}}
        out[alias] = {"records": len(recs), "pages": pages, "delta_link": bool(link)}
    LOCAL.mkdir(exist_ok=True)
    STATE.write_text(json.dumps(state), encoding="utf-8")
    return out


def check(token: str) -> dict:
    state = json.loads(STATE.read_text(encoding="utf-8"))
    salt = state["salt"]
    aliases = well_known_ids(token)
    known = {h: f for d in state["folders"].values() for h, f in d["ids"].items()}
    out = {}
    for alias, data in state["folders"].items():
        if not data.get("delta_link"):
            out[alias] = "no_delta_link"
            continue
        recs, _, pages = run_delta(token, data["delta_link"])
        kinds = Counter()
        removed_fate = Counter()
        changed_detail = Counter()
        for rec in recs:
            h = stable_hash(salt, rec.get("id", ""))
            if "@removed" in rec:
                reason = (rec["@removed"] or {}).get("reason", "?")
                kinds[f"removed:{reason}"] += 1
                g = graph(token, f"/me/messages/{rec['id']}", **{"$select": "parentFolderId"})
                if g.ok:
                    removed_fate[f"still_gettable_now_in:{aliases.get(g.payload.get('parentFolderId'), 'custom')}"] += 1
                else:
                    removed_fate[f"get_{g.status}:{g.error_code}"] += 1
            else:
                kinds["added_or_changed"] += 1
                prev = known.get(h)
                changed_detail["was_in:" + (prev or "not_seen_in_snapshot")] += 1
        out[alias] = {"pages": pages, "records": dict(kinds), "removed_items_fate": dict(removed_fate),
                      "changed_items_origin": dict(changed_detail)}
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--snapshot", action="store_true")
    g.add_argument("--check", action="store_true")
    args = ap.parse_args()
    token = need("graph")
    emit({"probe": "delta move/delete semantics (S2)",
          "phase": "snapshot" if args.snapshot else "check",
          "result": snapshot(token) if args.snapshot else check(token)})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
