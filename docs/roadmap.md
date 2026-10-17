# Outlook Connector Roadmap

Open work only. What the connector must do, its settled decisions and its ownership boundaries are in the
[requirements](requirements.md); how it is built is in the [architecture](architecture.md); dated live
evidence is in the [research record](api-research.md). Delete an item once it is done.

Work is split by what it needs: **implementation** is code still to write; **decisions** wait for the
user; **live verification** checks something on a real mailbox before it is built or relied on, and
becomes implementation only if the check succeeds.

Snapshot **2026-10-17**.

## Implementation

### Bugs

None open.

### Improvements

#### Passive HTML as an allow-list

Draft bodies and signatures pass one deny-list checker,
[`passive_html.py`](../src/outlook_connector/service/passive_html.py), and it reads markup with Python's
HTML parser, not a browser's, so markup the two tokenize differently could pass. Example: malformed markup
that Python reads as an attribute value but a browser reads as a new tag with an event handler. The one
save test is one account and one client ([research §4.5](api-research.md#45-what-outlook-keeps-of-saved-html)).
Candidates: an allow-list of elements and attributes (for example `p`, `div`, `span`, `table`, `a href`,
`img src`, inline `style`), or wider evidence of what Outlook strips, checked against browser parsing.

#### A purpose-named calendar profile

Calendar windows read Graph through the shared `teams.graph` profile, because `outlook.graph` has no
calendar scope ([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)). If
the Teams profile changes, `list_calendar_events` falls back to Outlook Web and loses descriptions,
`contains` and `include_body`. Plan: a calendar profile in `lrh-m365-auth` on the same client and
resource. The alternative, `Calendars.Read` on `outlook.graph`, needs its preauthorization checked first.

### Enhancements

#### Attach local files to drafts

Build the design recorded in [requirements §12.1](requirements.md#121-drafts-sending-and-signatures): an
`attachments` list on `create_draft`, sent as non-inline `FileAttachment:#Exchange` entries in the same
`CreateItem` call that already carries the signature's inline images
([`ows_mapping.py`](../src/outlook_connector/remote/ows_mapping.py); shape in
[research §4.2](api-research.md#42-item-writes)), with Graph read-back confirming each file by name and
size, and `CreateAttachment` on the saved draft for more than one request's worth. Settle while building:
where the send limit comes from ([research §4.6](api-research.md#46-message-size-limits) has two
candidate values), and whether a single file over the one-request limit is worth
`CreateAttachmentFromLocalFile` slices ([research §4.3](api-research.md#43-reads-and-other-actions)),
which are untested.

## Decisions

### Default conversation scope

`get_conversation` leaves out Deleted Items and Junk Email unless `scope.deleted_items=true`, the same
default as listing and search, so a conversation can look shorter than it does in Outlook. Example: you
reply to a thread and then delete the original; `get_conversation` returns only your reply, and the
original appears with `scope={"deleted_items": true}`. Confirm this default, or include deleted mail in
conversations by default, since a conversation is usually read for its whole history.

## Live verification

Checks against a real mailbox. Each needs the user's authorization for the live read it makes; record the
dated result in the research record.

### Fresh-account consent

The `m365.search` and `teams.graph` calendar grants are verified on one account only. Check them on a new
account before relying on them.

### Inline image placeholders

Attachment markers replace the bracketed alt text that Exchange's text body shows for an inline image
([research §3.6](api-research.md#36-attachments-and-inline-images)). Check what that text shows for an image
without alt text, and for two images with the same alt text; today such an image's marker follows the text.

### Mail search route

Compare Graph Search API with the current Graph `$search` route on identical queries
([research §3.4](api-research.md#34-mail-search)): terms, limits, hydration, request counts and latency,
before considering a change.

### Graph-first event hydration

`get_calendar_event` reads through Outlook Web (`GetItem`) only, while listing reads Graph first. Check the
Graph event's detail fields (attendees, response status, join link, attachment metadata) and the id
conversion ([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)). If they
hold, hydrate through Graph (`GET /me/events/{id}`) with Outlook Web as the fallback.

### Graph event search

`search_calendar_events` uses Search v2, which takes no date restriction and cannot page past offset
1,000. Check whether Graph Search (`POST /search/query`, `entityTypes=["event"]`;
[research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)) restricts dates on the
server and can target the description alone. If so, "budget review in the week of 2 November" could be one
ranked search limited to that week.

### Other calendars

Only the signed-in user's default calendar is listed and searched. Out of reach: secondary calendars the
user created ("What is on my Project Beta calendar next week?"), and calendars shared with or delegated to
the user ("What does Alex have on Thursday?"). Check reading the user's own secondary calendars with the
`teams.graph` token's calendar scope, and shared and delegated calendars with its `Calendars.Read.Shared`
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)).

### Reply headers for branch-aware conversations (parked)

A conversation is exactly Exchange's conversation: one flat list, oldest first, with nothing saying who
answered whom. Example: Alice sends a proposal, Bob replies to all, and Carol forwards it to Dan, who
replies to Carol; if Exchange groups all four, `get_conversation` returns them by date, and Dan's reply
reads as if it answered Bob. A forward that Exchange gives its own conversation id is not returned at all.
Check the quality of the RFC 5322 `Message-ID` / `In-Reply-To` / `References` headers, and of Exchange's
`ConversationId` and `Thread-Index`, in real mail. If they hold, building the reply tree becomes an
enhancement: detect branches, including forwards (looking across conversation ids where a forward left
the conversation), and offer a merged chronological view where each message is labeled with its branch.

### Event attachment download

`get_calendar_event` lists attachment metadata, but `download_attachment` is mail-only
([research §6.2](api-research.md#62-events)). Check an event-scoped download route, so an agenda PDF
attached to a meeting can be read.

### Event to Teams meeting chat

`get_calendar_event` returns a Teams join link but no chat id ([research §6.2](api-research.md#62-events)),
so an agent finds the meeting chat in the Teams connector by name and time. Look for a supported lookup
from the event or its join link to the chat; reading an id out of the join URL is not established.
