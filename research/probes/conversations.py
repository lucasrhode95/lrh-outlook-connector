"""R3 + R4: can Graph return a whole conversation across folders, and how good are conversation headers?

Read-only (profile 'graph'). Samples recent conversations from Sent Items and Inbox,
then for each conversation:

R3  GET /me/messages?$filter=conversationId eq '...'   (mailbox-wide)
    - which folders the messages live in (well-known alias or 'custom')
    - whether $orderby is accepted together with the conversationId filter
    - paging behaviour
R4  internetMessageHeaders per message
    - presence of Message-ID / In-Reply-To / References / Thread-Index
    - reply tree from In-Reply-To/References: roots, branch points, orphans
    - forward-prefixed subjects (FW/Fwd/ENC/WG/TR), classified in memory only
    - uniqueBody vs body length ratio (lengths only)

Output: counts and distributions only.
"""

from __future__ import annotations

import argparse
import re
import statistics
from collections import Counter

from common import emit, graph, graph_pages, need, rows, well_known_ids

FWD = re.compile(r"^\s*(fw|fwd|enc|wg|tr|rv)\s*:", re.I)
REPLY = re.compile(r"^\s*(re|res|aw|sv|antw)\s*:", re.I)
HEADER_NAMES = {"message-id", "in-reply-to", "references", "conversation-index", "conversation-topic"}


def sample_conversations(token: str, n: int) -> list[str]:
    seen: list[str] = []
    for alias in ("sentitems", "inbox"):
        r = graph(token, f"/me/mailFolders/{alias}/messages",
                  **{"$top": str(n * 2), "$select": "conversationId", "$orderby": "receivedDateTime desc"})
        for row in rows(r):
            cid = row.get("conversationId")
            if cid and cid not in seen:
                seen.append(cid)
        if len(seen) >= n:
            break
    return seen[:n]


def fetch_conversation(token: str, cid: str, max_messages: int) -> dict:
    flt = f"conversationId eq '{cid}'"
    select = "id,parentFolderId,receivedDateTime,isDraft,internetMessageId,subject"
    ordered = graph(token, "/me/messages", **{"$filter": flt, "$orderby": "receivedDateTime asc",
                                              "$select": select, "$top": "50"})
    msgs, pages, status = [], 0, None
    for r in graph_pages(token, "/me/messages", max_pages=max(1, max_messages // 50 + 1),
                         **{"$filter": flt, "$select": select, "$top": "50"}):
        pages += 1
        status = r.status if r.ok else r.summary()
        msgs.extend(rows(r))
        if len(msgs) >= max_messages:
            break
    return {"messages": msgs[:max_messages], "pages": pages, "status": status,
            "orderby_with_filter": ordered.status if ordered.ok else ordered.summary()}


def header_analysis(token: str, msgs: list[dict]) -> dict:
    ids_by_mid, parent_of, presence = {}, {}, Counter()
    unique_ratios, fwd, reply = [], 0, 0
    for m in msgs:
        r = graph(token, f"/me/messages/{m['id']}",
                  **{"$select": "internetMessageHeaders,internetMessageId,uniqueBody,body"},
                  headers={"Prefer": 'IdType="ImmutableId", outlook.body-content-type="text"'}, prefer=None)
        if not r.ok or not isinstance(r.payload, dict):
            presence["fetch_failed"] += 1
            continue
        p = r.payload
        hdrs = {}
        for h in p.get("internetMessageHeaders") or []:
            name = str(h.get("name", "")).lower()
            if name in HEADER_NAMES:
                hdrs[name] = str(h.get("value", ""))
        if p.get("internetMessageHeaders") is None:
            presence["headers_missing_entirely"] += 1
        for name in HEADER_NAMES:
            presence[name] += name in hdrs
        mid = p.get("internetMessageId") or hdrs.get("message-id")
        if mid:
            ids_by_mid[mid.strip()] = m["id"]
        refs = re.findall(r"<[^>]+>", hdrs.get("references", ""))
        parent = (hdrs.get("in-reply-to") or "").strip() or (refs[-1] if refs else None)
        parent_of[m["id"]] = parent
        subj = m.get("subject") or ""
        fwd += bool(FWD.match(subj))
        reply += bool(REPLY.match(subj))
        body_len = len(((p.get("body") or {}).get("content")) or "")
        uniq_len = len(((p.get("uniqueBody") or {}).get("content")) or "")
        if body_len:
            unique_ratios.append(round(uniq_len / body_len, 3))
    children = Counter()
    roots = orphans = 0
    for mid_id, parent in parent_of.items():
        if not parent:
            roots += 1
        elif parent in ids_by_mid:
            children[parent] += 1
        else:
            orphans += 1  # parent not in this conversation (deleted, not received, or other mailbox)
    return {
        "messages_analysed": len(parent_of),
        "header_presence": dict(presence),
        "roots": roots,
        "orphans_parent_not_in_conversation": orphans,
        "branch_points": sum(1 for c in children.values() if c > 1),
        "max_children": max(children.values(), default=0),
        "forward_prefixed": fwd,
        "reply_prefixed": reply,
        "unique_to_full_body_ratio_median": statistics.median(unique_ratios) if unique_ratios else None,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--conversations", type=int, default=15)
    ap.add_argument("--max-messages", type=int, default=40)
    ap.add_argument("--skip-headers", action="store_true")
    args = ap.parse_args()
    token = need("graph")
    aliases = well_known_ids(token)
    cids = sample_conversations(token, args.conversations)

    per_conv, folder_mix, sizes, spans = [], Counter(), [], Counter()
    totals = Counter()
    for i, cid in enumerate(cids, 1):
        conv = fetch_conversation(token, cid, args.max_messages)
        msgs = conv["messages"]
        folders = Counter(aliases.get(m.get("parentFolderId"), "custom") for m in msgs)
        folder_mix.update(folders)
        sizes.append(len(msgs))
        spans[len(folders)] += 1
        entry = {"conversation": f"c{i}", "message_count": len(msgs), "pages": conv["pages"],
                 "list_status": conv["status"], "orderby_with_filter": conv["orderby_with_filter"],
                 "folders": dict(folders), "drafts": sum(1 for m in msgs if m.get("isDraft"))}
        if not args.skip_headers and len(msgs) >= 2:
            entry["conversation_headers"] = header_analysis(token, msgs)
            t = entry["conversation_headers"]
            totals["messages"] += t["messages_analysed"]
            totals["branch_points"] += t["branch_points"]
            totals["orphans"] += t["orphans_parent_not_in_conversation"]
            totals["conversations_with_branches"] += t["branch_points"] > 0
            totals["forwards"] += t["forward_prefixed"]
            for k, v in t["header_presence"].items():
                totals[f"hdr:{k}"] += v
        per_conv.append(entry)

    emit({
        "probe": "conversations (R3/R4)",
        "conversations_sampled": len(cids),
        "R3_summary": {
            "message_count_distribution": {"min": min(sizes, default=0), "median": statistics.median(sizes) if sizes else 0,
                                           "max": max(sizes, default=0)},
            "conversations_by_number_of_folders": dict(sorted(spans.items())),
            "messages_by_folder": dict(folder_mix),
            "orderby_with_conversation_filter_ok": all(c["orderby_with_filter"] == 200 for c in per_conv),
        },
        "R4_summary": dict(totals),
        "per_conversation": per_conv,
    })
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
