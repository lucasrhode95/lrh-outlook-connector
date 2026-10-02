"""R1: how big is the mailbox, and what would an eager metadata catalog cost?

Read-only (profile 'read'). Reports:
- recursive folder tree (hidden included): folder count, item totals, top folders by
  size (labelled by well-known alias or 'custom#n', never by name)
- the unexplained hierarchy(23) vs folder-delta(22) difference: which folder is missing
  and its shape (hidden / child count / item count / alias)
- whether mailbox-wide /me/messages $count works and matches the folder sum
- timing of a mailbox-wide metadata walk and an Inbox message-delta bootstrap,
  page-capped, with an extrapolated full-bootstrap estimate

Output: counts, timings and sizes only.
"""

from __future__ import annotations

import argparse
import time

from common import emit, graph, graph_pages, need, rows, well_known_ids

FOLDER_SELECT = "id,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden"
META_SELECT = ("id,conversationId,parentFolderId,subject,from,toRecipients,ccRecipients,receivedDateTime,"
               "sentDateTime,isRead,isDraft,hasAttachments,importance,categories,bodyPreview,changeKey,"
               "internetMessageId,flag")


def walk_folders(token: str) -> tuple[list[dict], int]:
    folders, calls = [], 0
    queue = [None]
    while queue:
        parent = queue.pop()
        path = "/me/mailFolders" if parent is None else f"/me/mailFolders/{parent}/childFolders"
        for r in graph_pages(token, path, max_pages=20, **{"$select": FOLDER_SELECT, "$top": "100",
                                                           "includeHiddenFolders": "true"}):
            calls += 1
            for f in rows(r):
                folders.append(f)
                if f.get("childFolderCount"):
                    queue.append(f["id"])
    return folders, calls


def folder_delta(token: str) -> tuple[list[dict], dict]:
    out, pages, start = [], 0, time.monotonic()
    for r in graph_pages(token, "/me/mailFolders/delta", max_pages=20, **{"$select": FOLDER_SELECT}):
        pages += 1
        out.extend(rows(r))
    return out, {"pages": pages, "elapsed_ms": int((time.monotonic() - start) * 1000)}


def timed_walk(token: str, path: str, max_pages: int, page_size: int, delta: bool) -> dict:
    params = {"$select": META_SELECT}
    headers = None
    if delta:
        headers = {"Prefer": f'IdType="ImmutableId", odata.maxpagesize={page_size}'}
    else:
        params["$top"] = str(page_size)
    count = pages = bytes_ = 0
    start = time.monotonic()
    status = None
    more = False
    for r in graph_pages(token, path, max_pages=max_pages, prefer=None if delta else 'IdType="ImmutableId"',
                         headers=headers, **params):
        pages += 1
        status = r.status if r.ok else r.summary()
        count += len(rows(r))
        bytes_ += len(r.raw)
        more = isinstance(r.payload, dict) and bool(r.payload.get("@odata.nextLink"))
    elapsed = time.monotonic() - start
    return {"path": path, "status": status, "pages": pages, "records": count, "bytes": bytes_,
            "elapsed_s": round(elapsed, 2), "more_pages_remaining": more,
            "records_per_s": round(count / elapsed, 1) if elapsed else None,
            "bytes_per_record": round(bytes_ / count) if count else None}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--max-pages", type=int, default=5, help="page cap for each timed walk")
    ap.add_argument("--page-size", type=int, default=500)
    args = ap.parse_args()
    token = need("read")
    aliases = well_known_ids(token)

    folders, folder_calls = walk_folders(token)
    by_id = {f["id"]: f for f in folders}
    total_items = sum(f.get("totalItemCount") or 0 for f in folders)
    ranked = sorted(folders, key=lambda f: f.get("totalItemCount") or 0, reverse=True)
    custom_n = 0
    top = []
    for f in ranked[:15]:
        label = aliases.get(f["id"])
        if not label:
            custom_n += 1
            label = f"custom#{custom_n}"
        top.append({"folder": label, "items": f.get("totalItemCount"), "unread": f.get("unreadItemCount"),
                    "hidden": f.get("isHidden"), "children": f.get("childFolderCount")})

    delta_rows, delta_meta = folder_delta(token)
    delta_ids = {f["id"] for f in delta_rows if "@removed" not in f}
    hier_ids = set(by_id)
    missing = [{"alias": aliases.get(i, "custom"), "hidden": by_id[i].get("isHidden"),
                "children": by_id[i].get("childFolderCount"), "items": by_id[i].get("totalItemCount"),
                "parent_is_root_level": by_id[i].get("parentFolderId") not in hier_ids}
               for i in hier_ids - delta_ids]

    count_resp = graph(token, "/me/messages", headers={"ConsistencyLevel": "eventual"},
                       **{"$count": "true", "$top": "1", "$select": "id"})
    mailbox_count = count_resp.payload.get("@odata.count") if isinstance(count_resp.payload, dict) else None

    walk = timed_walk(token, "/me/messages", args.max_pages, args.page_size, delta=False)
    inbox_delta = timed_walk(token, "/me/mailFolders/inbox/messages/delta", args.max_pages, args.page_size, delta=True)
    est = None
    if walk["records_per_s"] and total_items:
        est = round(total_items / walk["records_per_s"] / 60, 1)

    emit({
        "probe": "catalog (R1)",
        "folders": {"count": len(folders), "listing_calls": folder_calls, "total_items": total_items,
                    "hidden": sum(1 for f in folders if f.get("isHidden")), "top_by_items": top},
        "folder_delta": {**delta_meta, "records": len(delta_rows),
                         "in_hierarchy_not_in_delta": missing,
                         "in_delta_not_in_hierarchy": len(delta_ids - hier_ids)},
        "mailbox_count": {"status": count_resp.status if count_resp.ok else count_resp.summary(),
                          "@odata.count": mailbox_count,
                          "difference_vs_folder_sum": (total_items - mailbox_count) if isinstance(mailbox_count, int) else None},
        "timed_mailbox_metadata_walk": walk,
        "timed_inbox_delta_bootstrap": inbox_delta,
        "estimated_full_metadata_bootstrap_minutes": est,
    })
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
