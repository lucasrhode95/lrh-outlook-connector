"""R2: search bake-off. Graph $search vs Graph Microsoft Search vs Substrate (Outlook top bar).

    python research/probes/search.py "relatório be" "relatorio be" "subject:relatório"
    python research/probes/search.py --no-substrate "..."

Read-only. Profiles: 'graph' (Graph) and, unless --no-substrate, 'search'
(One Outlook Web -> https://outlook.office.com/search/.default).

Per query and backend: HTTP status, result count over N pages, server total /
more-available flags, latency, and distinct conversations. Conversation sets are
compared in memory and only overlap counts are printed. Query strings are echoed
because they are the operator's own test input; result content is not.
"""

from __future__ import annotations

import argparse
import uuid
from typing import Any

from common import account, emit, graph, graph_pages, http, need, rows

from config import SETTINGS

SUBSTRATE_HOSTS = SETTINGS["substrate_urls"]


def graph_dollar_search(token: str, q: str, top: int, pages: int) -> dict[str, Any]:
    term = q.replace('"', '\\"')
    convs, ids, statuses, ms = [], [], [], 0
    more = False
    for r in graph_pages(token, "/me/messages", max_pages=pages,
                         **{"$search": f'"{term}"', "$top": str(top), "$select": "id,conversationId"}):
        statuses.append(r.status if r.ok else r.summary())
        ms += r.elapsed_ms
        for row in rows(r):
            ids.append(row.get("id"))
            convs.append(row.get("conversationId"))
        more = isinstance(r.payload, dict) and bool(r.payload.get("@odata.nextLink"))
    return {"statuses": statuses, "messages": len(ids), "conversations": len(set(convs)),
            "more_available": more, "elapsed_ms": ms, "_convs": set(convs)}


def graph_microsoft_search(token: str, q: str, size: int, pages: int) -> dict[str, Any]:
    msg_ids, statuses, ms, total, more = [], [], 0, None, False
    for page in range(pages):
        body = {"requests": [{"entityTypes": ["message"], "query": {"queryString": q},
                              "from": page * size, "size": size}]}
        r = graph(token, "/search/query", method="POST", body=body)
        statuses.append(r.status if r.ok else r.summary())
        ms += r.elapsed_ms
        if not r.ok:
            break
        containers = [c for v in (r.payload or {}).get("value", []) for c in v.get("hitsContainers", [])]
        hits = [h for c in containers for h in c.get("hits", []) or []]
        total = containers[0].get("total") if containers else total
        more = bool(containers and containers[0].get("moreResultsAvailable"))
        msg_ids.extend(h.get("hitId") or (h.get("resource") or {}).get("id") for h in hits)
        if not more:
            break
    # Microsoft Search hits do not carry conversationId; resolve a bounded number.
    convs, resolve_fail = set(), 0
    for mid in msg_ids[:60]:
        r = graph(token, f"/me/messages/{mid}", **{"$select": "conversationId"})
        if r.ok and isinstance(r.payload, dict):
            convs.add(r.payload.get("conversationId"))
        else:
            resolve_fail += 1
    return {"statuses": statuses, "messages": len(msg_ids), "server_total": total, "more_available": more,
            "conversations_resolved": len(convs), "resolve_failures": resolve_fail, "elapsed_ms": ms, "_convs": convs}


def substrate_body(q: str, start: int, size: int, variant: str, session: tuple[str, str]) -> dict[str, Any]:
    entity: dict[str, Any] = {
        "EntityType": "Conversation", "ContentSources": ["Exchange"],
        "Query": {"QueryString": q, "DisplayQueryString": q},
        "From": start, "Size": size, "RefiningQueries": None, "Sort": [],
        "EnableTopResults": False, "PropertySet": "Optimized",
    }
    if variant == "folder_filter":
        entity["Filter"] = {"Or": [{"Term": {"DistinguishedFolderName": "msgfolderroot"}},
                                   {"Term": {"DistinguishedFolderName": "DeletedItems"}}]}
    if variant == "time_sort":
        entity["Sort"] = [{"Field": "Time", "SortDirection": "Desc"}]
    return {"Cvid": session[0], "LogicalId": session[1], "Scenario": {"Name": "owa.react"},
            "TimeZone": "UTC", "TextDecorations": "Off", "EntityRequests": [entity]}


def substrate_search(token: str, q: str, size: int, pages: int, anchor: str, url: str, variant: str) -> dict[str, Any]:
    convs, ids, statuses, ms, total, more, partial = [], [], [], 0, None, False, None
    session = (str(uuid.uuid4()), str(uuid.uuid4()))  # one search session across pages
    for page in range(pages):
        r = http("POST", url, token=token, json_body=substrate_body(q, page * size, size, variant, session),
                 headers={"X-AnchorMailbox": anchor, "Prefer": 'IdType="ImmutableId"'}, retries_429=0)
        statuses.append(r.status if r.ok else r.summary())
        ms += r.elapsed_ms
        if not r.ok or not isinstance(r.payload, dict):
            break
        sets = [rs for es in r.payload.get("EntitySets", []) for rs in es.get("ResultSets", [])]
        partial = any(es.get("IsPartial") for es in r.payload.get("EntitySets", []))
        results = [x for rs in sets for x in rs.get("Results", []) or []]
        if sets:
            total, more = sets[0].get("Total"), bool(sets[0].get("MoreResultsAvailable"))
        for res in results:
            src = res.get("Source") or {}
            cid = src.get("ConversationId")
            convs.append(cid.get("Id") if isinstance(cid, dict) else cid)
            ids.append(src.get("ImmutableId"))
        if not more:
            break
    return {"url": url, "variant": variant, "statuses": statuses, "conversation_results": len(convs),
            "server_total": total, "more_available": more, "is_partial": partial,
            "immutable_id_present": sum(1 for i in ids if i), "elapsed_ms": ms, "_convs": set(convs),
            "_sample_ids": [i for i in ids if i][:3], "_ids": [i for i in ids if i]}


def find_substrate(token: str, anchor: str) -> tuple[str | None, str | None, list]:
    attempts = []
    for url in SUBSTRATE_HOSTS:
        for variant in ("plain", "folder_filter"):
            res = substrate_search(token, "test", 5, 1, anchor, url, variant)
            attempts.append({"url": url, "variant": variant, "status": res["statuses"]})
            if res["statuses"] and res["statuses"][0] == 200:
                return url, variant, attempts
    return None, None, attempts


def id_alignment(token: str, graph_convs: set, substrate: dict) -> dict[str, Any]:
    """Do Substrate ConversationId / ImmutableId values line up with Graph?"""
    out: dict[str, Any] = {"conversation_id_overlap": len(graph_convs & substrate["_convs"])}
    tried = []
    for source_type in ("ewsId", "restImmutableEntryId", "restId"):
        ids = substrate["_sample_ids"]
        if not ids:
            break
        r = graph(token, "/me/translateExchangeIds", method="POST",
                  body={"inputIds": ids, "sourceIdType": source_type, "targetIdType": "restImmutableEntryId"})
        ok = sum(1 for v in rows(r) if v.get("targetId"))
        tried.append({"sourceIdType": source_type, "status": r.status if r.ok else r.summary(), "translated": ok})
        direct = graph(token, f"/me/messages/{ids[0]}", **{"$select": "id"})
        tried[-1]["immutable_id_directly_gettable"] = direct.status
        if ok:
            break
    out["immutable_id_translation"] = tried
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("queries", nargs="+")
    ap.add_argument("--size", type=int, default=25)
    ap.add_argument("--pages", type=int, default=2)
    ap.add_argument("--no-substrate", action="store_true")
    args = ap.parse_args()
    read = need("graph")
    search_token, anchor, substrate_url, variant, discovery = None, None, None, None, None
    if not args.no_substrate:
        search_token = need("search")
        acc = account("search")
        anchor = f"Oid:{acc.get('oid')}@{acc.get('tid')}"
        substrate_url, variant, discovery = find_substrate(search_token, anchor)

    report = []
    for i, q in enumerate(args.queries, 1):
        entry: dict[str, Any] = {"query": q}
        a = graph_dollar_search(read, q, args.size, args.pages)
        b = graph_microsoft_search(read, q, args.size, args.pages)
        entry["graph_$search"] = {k: v for k, v in a.items() if not k.startswith("_")}
        entry["graph_search_query"] = {k: v for k, v in b.items() if not k.startswith("_")}
        entry["overlap_$search_vs_search_query"] = len(a["_convs"] & b["_convs"])
        if substrate_url:
            c = substrate_search(search_token, q, args.size, args.pages, anchor, substrate_url, variant)
            entry["substrate"] = {k: v for k, v in c.items() if not k.startswith("_")}
            resolved, missing, folders = set(), 0, 0
            for iid in c["_ids"][:100]:
                g = graph(read, f"/me/messages/{iid}", **{"$select": "conversationId,parentFolderId"})
                if g.ok and isinstance(g.payload, dict):
                    resolved.add(g.payload.get("conversationId"))
                else:
                    missing += 1
            entry["substrate_resolved_via_graph"] = {
                "conversations": len(resolved), "unresolvable": missing,
                "raw_conversation_id_equals_graph": len(c["_convs"] & resolved),
                "overlap_$search": len(a["_convs"] & resolved), "overlap_search_query": len(b["_convs"] & resolved)}
            entry["overlap_vs_substrate"] = {"$search": len(a["_convs"] & c["_convs"]),
                                            "search_query": len(b["_convs"] & c["_convs"])}
            if i == 1:
                entry["substrate_id_alignment"] = id_alignment(read, a["_convs"] | b["_convs"], c)
        report.append(entry)
    emit({"probe": "search bake-off (R2)", "substrate_discovery": discovery, "results": report})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
