"""Attachment shapes for exports: types, inline images, item/reference attachments.

Read-only (profile 'graph'). Scans recent messages with attachments and reports:
- counts by @odata.type (file / item / reference) and isInline
- inline file attachments whose contentId is referenced as cid: in body vs uniqueBody
- whether /$value works per attachment type (first one of each type only)
- size distribution

Attachment bytes stay in memory and are discarded. Names are never printed.
"""

from __future__ import annotations

import argparse
import statistics
from collections import Counter

from common import emit, graph, http, graph_url, need, rows


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--messages", type=int, default=40)
    args = ap.parse_args()
    token = need("graph")
    r = graph(token, "/me/messages", **{"$filter": "receivedDateTime ge 2000-01-01T00:00:00Z and hasAttachments eq true", "$top": str(args.messages),
                                         "$select": "id", "$orderby": "receivedDateTime desc"})
    messages = rows(r)
    if not r.ok:
        emit({"probe": "attachments", "message_query": r.summary()})
        return 1
    types, inline, sizes = Counter(), Counter(), []
    cid_in_body = cid_in_unique = inline_with_cid = 0
    value_checks: dict[str, int | None] = {}
    # Also look at messages flagged hasAttachments=false that still carry inline images.
    for m in messages:
        mid = m["id"]
        att = graph(token, f"/me/messages/{mid}/attachments",
                    **{"$select": "id,isInline,size,contentType,lastModifiedDateTime"})
        if not att.ok:
            types[f"list_error:{att.status}"] += 1
            continue
        atts = rows(att)
        needs_body = False
        for a in atts:
            kind = str(a.get("@odata.type", "?")).rsplit(".", 1)[-1]
            types[kind] += 1
            inline[f"{kind}:inline={a.get('isInline')}"] += 1
            if isinstance(a.get("size"), int):
                sizes.append(a["size"])
            if kind not in value_checks:
                v = http("GET", graph_url(f"/me/messages/{mid}/attachments/{a['id']}/$value"), token=token,
                         limit=8 * 1024 * 1024, parse_json=False)
                value_checks[kind] = v.status
            needs_body |= bool(a.get("isInline"))
        if needs_body:
            full = graph(token, f"/me/messages/{mid}", **{"$select": "body,uniqueBody"})
            body = ((full.payload or {}).get("body") or {}).get("content") or ""
            uniq = ((full.payload or {}).get("uniqueBody") or {}).get("content") or ""
            for a in atts:
                if not a.get("isInline"):
                    continue
                # contentId is only on fileAttachment; fetch it without bytes via a typed select
                d = graph(token, f"/me/messages/{mid}/attachments/{a['id']}",
                          **{"$select": "microsoft.graph.fileAttachment/contentId"})
                cid = (d.payload or {}).get("contentId") if d.ok else None
                if cid:
                    inline_with_cid += 1
                    cid_in_body += f"cid:{cid}" in body
                    cid_in_unique += f"cid:{cid}" in uniq
    emit({
        "probe": "attachments",
        "messages_scanned": len(messages),
        "attachments_by_type": dict(types),
        "inline_breakdown": dict(inline),
        "value_endpoint_status_by_type": value_checks,
        "inline_with_content_id": inline_with_cid,
        "inline_cid_referenced_in_body": cid_in_body,
        "inline_cid_referenced_in_uniqueBody": cid_in_unique,
        "size_bytes": {"median": statistics.median(sizes) if sizes else None, "max": max(sizes, default=None)},
    })
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
