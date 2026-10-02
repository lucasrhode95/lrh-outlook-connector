"""OWS write contracts (read state, flag, categories, conversation read state, move, soft delete).

Graph cannot write mail with the usable first-party clients (research §2), so these
contracts are exercised through OWS with the 'write' profile and verified read-only
through Graph with the 'read' profile.

    python research/probes/ows_mutations.py \\
        --case read,flag,categories,conversation-read "A TEST EMAIL HALPRIO190" \\
        --case move-roundtrip "A TEST EMAIL 20A9LK" \\
        --case delete "Some newsletter subject" \\
        --case move-to-deleted "A TEST EMAIL 777AVCQPZL"

Without --i-authorize-mutations it only locates the targets and prints the plan (read-only).

Safety rules:
- Targets are exact subjects named by the operator. Only the copy in the Inbox is touched,
  and only when exactly one Inbox copy matches.
- No send. No hard delete. 'delete' = DeleteItem with MoveToDeletedItems.
- read/flag/categories/conversation-read toggle and then restore the original value.
- Writes are never retried after an ambiguous transport result.

Output: OWS status codes and Graph-observed state (folder alias, isRead, flag, category
count). No subjects beyond the case labels, no ids.
"""

from __future__ import annotations

import argparse
from typing import Any

from common import emit, graph, need, ows, ows_items, ows_result, rows, well_known_ids

OPS = {"read", "flag", "categories", "conversation-read", "move-roundtrip", "delete", "move-to-deleted"}


def to_ows_id(graph_id: str) -> str:
    """Graph immutable REST id -> OWS ItemId (base64 alphabet swap, research §4.2)."""
    return graph_id.replace("-", "/").replace("_", "+")


def returned_item(r) -> dict[str, Any]:
    items = ows_items(r)
    inner = items[0].get("Items") if items else None
    return inner[0] if isinstance(inner, list) and inner else {}


def observed(read: str, graph_id: str, aliases: dict[str, str]) -> dict[str, Any]:
    r = graph(read, f"/me/messages/{graph_id}", **{"$select": "parentFolderId,isRead,flag,categories"})
    if not r.ok:
        return {"graph_get": r.status, "error_code": r.error_code}
    p = r.payload
    return {"folder": aliases.get(p.get("parentFolderId"), "custom"), "isRead": p.get("isRead"),
            "flag": (p.get("flag") or {}).get("flagStatus"), "category_count": len(p.get("categories") or [])}


def update(write: str, item_id: str, field_uri: str, props: dict[str, Any]) -> dict[str, Any]:
    body = {
        "ItemChanges": [{
            "__type": "ItemChange:#Exchange",
            "ItemId": {"__type": "ItemId:#Exchange", "Id": item_id},
            "Updates": [{"__type": "SetItemField:#Exchange",
                         "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": field_uri},
                         "Item": {"__type": "Message:#Exchange", **props}}],
        }],
        "ConflictResolution": "AlwaysOverwrite", "MessageDisposition": "SaveOnly",
        "SuppressReadReceipts": True, "SendCalendarInvitationsOrCancellations": "SendToNone",
    }
    return ows_result(ows(write, "UpdateItem", "UpdateItemRequest", body))


def move(write: str, item_id: str, target: str) -> tuple[dict[str, Any], str]:
    body = {"ToFolderId": {"__type": "TargetFolderId:#Exchange",
                           "BaseFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": target}},
            "ItemIds": [{"__type": "ItemId:#Exchange", "Id": item_id}], "ReturnNewItemIds": True}
    r = ows(write, "MoveItem", "MoveItemRequest", body)
    new = returned_item(r).get("ItemId")
    new_id = new.get("Id") if isinstance(new, dict) and new.get("Id") else item_id
    return {**ows_result(r), "ows_id_changed": new_id != item_id}, new_id


def conversation_read(write: str, conversation_id: str, is_read: bool) -> dict[str, Any]:
    r = ows(write, "ApplyConversationAction", "ApplyConversationActionRequest", {"ConversationActions": [{
        "__type": "ConversationAction:#Exchange", "Action": "SetReadState",
        "ConversationId": {"__type": "ItemId:#Exchange", "Id": to_ows_id(conversation_id)},
        "ContextFolderId": {"__type": "TargetFolderId:#Exchange",
                            "BaseFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "inbox"}},
        "IsRead": is_read, "SuppressReadReceipts": True}]})
    return ows_result(r)


def run_case(read: str, write: str, ops: list[str], target: dict[str, Any], aliases: dict[str, str]) -> dict[str, Any]:
    gid, item = target["id"], to_ows_id(target["id"])
    start = observed(read, gid, aliases)
    steps: dict[str, Any] = {"initial": start}

    def step(name: str, result: dict[str, Any]) -> bool:
        result["graph_after"] = observed(read, gid, aliases)
        steps[name] = result
        return result.get("response_class") == "Success"

    if "read" in ops:
        orig = bool(start.get("isRead"))
        step("read_toggle", update(write, item, "message:IsRead", {"IsRead": not orig})) and \
            step("read_restore", update(write, item, "message:IsRead", {"IsRead": orig}))
    if "flag" in ops:
        orig = "Flagged" if start.get("flag") == "flagged" else "NotFlagged"
        other = "NotFlagged" if orig == "Flagged" else "Flagged"
        flag = lambda status: {"Flag": {"__type": "FlagType:#Exchange", "FlagStatus": status}}  # noqa: E731
        step("flag_toggle", update(write, item, "item:Flag", flag(other))) and \
            step("flag_restore", update(write, item, "item:Flag", flag(orig)))
    if "categories" in ops and start.get("category_count") == 0:
        step("categories_set", update(write, item, "item:Categories", {"Categories": ["Connector Test"]})) and \
            step("categories_clear", update(write, item, "item:Categories", {"Categories": []}))
    elif "categories" in ops:
        steps["categories"] = "skipped: message already has categories"
    if "conversation-read" in ops:
        orig = bool(start.get("isRead"))
        step("conversation_read_toggle", conversation_read(write, target["conversationId"], not orig)) and \
            step("conversation_read_restore", conversation_read(write, target["conversationId"], orig))
    if "move-roundtrip" in ops:
        res, item = move(write, item, "archive")
        if step("move_to_archive", res):
            res, item = move(write, item, "inbox")
            step("move_back_to_inbox", res)
    if "delete" in ops:
        r = ows(write, "DeleteItem", "DeleteItemRequest", {
            "ItemIds": [{"__type": "ItemId:#Exchange", "Id": item}], "DeleteType": "MoveToDeletedItems",
            "SendMeetingCancellations": "SendToNone", "AffectedTaskOccurrences": "AllOccurrences",
            "SuppressReadReceipts": True})
        step("delete_soft", ows_result(r))
    elif "move-to-deleted" in ops:
        res, item = move(write, item, "deleteditems")
        step("move_to_deleteditems", res)
    return steps


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--case", nargs=2, action="append", metavar=("OPS", "SUBJECT"), required=True,
                    help=f"comma-separated ops from {sorted(OPS)} and an exact subject")
    ap.add_argument("--i-authorize-mutations", action="store_true")
    args = ap.parse_args()
    cases = []
    for ops, subject in args.case:
        op_list = [o.strip() for o in ops.split(",") if o.strip()]
        unknown = set(op_list) - OPS
        if unknown:
            ap.error(f"unknown ops: {sorted(unknown)}")
        if {"delete", "move-to-deleted"} <= set(op_list):
            ap.error("choose either delete or move-to-deleted for one message")
        cases.append((op_list, subject))

    read = need("read")
    aliases = well_known_ids(read)
    inbox_id = next((k for k, v in aliases.items() if v == "inbox"), None)
    located: list[dict[str, Any]] = []
    for i, (ops, subject) in enumerate(cases, 1):
        esc = subject.replace("'", "''")
        hits = rows(graph(read, "/me/messages", **{"$filter": f"subject eq '{esc}'",
                                                  "$select": "id,parentFolderId,conversationId"}))
        inbox_hits = [h for h in hits if h.get("parentFolderId") == inbox_id]
        located.append({"case": f"c{i}", "ops": ops, "matches_total": len(hits),
                        "inbox_matches": len(inbox_hits), "_target": inbox_hits[0] if len(inbox_hits) == 1 else None})

    plan = [{k: v for k, v in c.items() if not k.startswith("_")} for c in located]
    if not args.i_authorize_mutations:
        emit({"probe": "OWS write contracts", "mode": "plan only (no changes)", "cases": plan})
        return 0
    if any(c["_target"] is None for c in located):
        emit({"abort": "every case needs exactly one Inbox match", "cases": plan})
        return 1

    write = need("write")
    results = {c["case"]: run_case(read, write, c["ops"], c["_target"], aliases) for c in located}
    emit({"probe": "OWS write contracts", "mode": "executed", "cases": plan, "results": results})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
