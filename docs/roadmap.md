# Outlook Connector Roadmap

Open work only. What the connector must do, its settled decisions and its ownership boundaries are in the
[requirements](requirements.md); how it is built is in the [architecture](architecture.md); dated live
evidence is in the [research record](api-research.md). Delete an item once it is done.

Work is split by what it needs:

- **Implementation** is code still to write.
- **Decisions** wait for the user.
- **Live verification** checks something on a real mailbox before it is built or relied on. It becomes
  implementation only if the check succeeds.

Each item opens with a fictional example of the problem today, then says what the item changes, decides
or checks.

Snapshot **2026-11-02**.

## Implementation

### Bugs

None open.

### Improvements

None open.

### Enhancements

#### Attach local files to drafts

You ask the agent: "Draft a reply to the team and attach `report.pdf`." `create_draft` takes no files
today. The draft is saved without the PDF, and you attach it by hand in Outlook.

**Change:** build the design in
[requirements §12.1](requirements.md#121-drafts-sending-and-signatures):

- An `attachments` list of local paths on `create_draft`.
- The files go as non-inline `FileAttachment:#Exchange` entries in the same `CreateItem` call that already
  carries the signature's inline images
  ([`ows_mapping.py`](../src/outlook_connector/remote/ows_mapping.py); shape in
  [research §4.2](api-research.md#42-item-writes)).
- Graph read-back confirms each file by name and size.
- `CreateAttachment` on the saved draft adds files beyond what one request holds.

Settle while building:

- Where the send limit comes from. [Research §4.6](api-research.md#46-message-size-limits) has two
  candidate values: 35 MB from the bounce message and 36 MiB from MailTips.
- Whether one file over the one-request limit, about 21 MiB, is worth `CreateAttachmentFromLocalFile`
  slices ([research §4.3](api-research.md#43-reads-and-other-actions)). They are untested.

## Decisions

### Reading text and inline image placement

A newsletter has a table, a quoted reply and two inline charts. `get_message` shows its text in one of two
ways today:

- **Local rendering.** The message has inline files, so the connector parses the source HTML itself and
  puts `[Image: chart.png]` where each image sits.
- **Exchange text with a list.** If the parser gives up, for example on an Outlook-only conditional
  comment around an image, the text is Exchange's own conversion. The charts then appear under "Inline
  attachments without a body position".

A message without inline files always reads Exchange's text. So two similar messages can read
differently.

Exchange's text cannot simply carry the image positions. Its `[cid:…]` placeholders are not reliable
([2026-10-18 conversion results](api-research.md#html-to-text-body)):

- An encoded CID can differ between the full and unique bodies.
- Ordinary prose can contain the same literal marker.
- A mail gateway's hidden preview can duplicate markers.

**Decide** one approach:

- Keep Exchange's reading text and place images in it, validated against the source HTML.
- Render text locally from the source structure. The current local rendering is not a settled design.

### Default conversation scope

You reply to a thread and then delete the original. `get_conversation` returns only your reply. The
original appears with `scope={"deleted_items": true}`. Conversations leave out Deleted Items and Junk Email
by default, the same as listing and search, so a conversation can look shorter than it does in Outlook.

**Decide:** keep this default, or include deleted mail in conversations by default, since a conversation
is usually read for its whole history.

## Live verification

Checks against a real mailbox. Each needs the user's authorization for the live read it makes. Record the
dated result in the research record.

### HTML-to-text conversion across mail clients

A colleague pastes a screenshot into a reply in Outlook desktop. Nobody knows yet which CID Outlook
desktop gives it, or how Exchange's text shows it. The
[2026-10-18 matrix](api-research.md#html-to-text-body) covers only these sources:

- Drafts created through Outlook Web Service (OWS), and their self-delivery.
- Binary paste in Outlook Web.
- Gmail MIME replies written through an integration.

**Check** the same cases with:

- Outlook desktop.
- Apple Mail.
- Gmail's browser and mobile editors.

Cover their generated CIDs, native quotation and `uniqueBody` behavior. Also check:

- VML-only images without an HTML fallback.
- Other percent-escape spellings in CIDs.
- Whether other deployments' mail filters expose or duplicate hidden previews.

### CID attachment carry-over in replies

You ask for a reply draft to a message whose quoted history holds several inline images. `create_draft`
returns `history_intact: false` and "2 inline image(s) of the original are missing or changed". The quoted
HTML still points at both images, but the reply has no files for them. In the test, the two lost files
were the edge cases: a URI-encoded CID and a CID with literal angle brackets
([research §3.6](api-research.md#html-to-text-body)).

**Check** where the files get lost before choosing a policy:

- Isolate URI decoding and MIME `Content-ID` normalization, with valid CID variants.
- Compare replies made in native clients.

### Fresh-account consent

A second user installs the connector and runs `outlook-connector auth` for the first time. The calendar
grants have only been checked on one account. If the new account is refused:

- `m365.search`: `search_calendar_events` fails.
- `outlook.calendar` (the Teams client): `list_calendar_events` falls back to Outlook Web, without descriptions.

**Check** both grants on a new account before relying on them.

### Mail search route

`search_messages` for `forecast` misses a message that mentions the word only inside an attached PDF. The
current route, Graph `$search`, matches attachment names, not attachment content. Graph Search API
(`/search/query`) searches attachment content but does not say which attachment matched
([research §3.4](api-research.md#34-mail-search)).

**Compare** the two routes on identical queries before considering a change:

- Terms and recall.
- Limits.
- Hydration: `$search` hits need an extra id translation.
- Request counts.
- Latency.

### Graph-first event hydration

`get_calendar_event` reads only through Outlook Web (`GetItem`), while listing reads Graph first. An
agenda over 2,048 characters comes back cut, with `body_truncated: true`, because that is the default
shape's cap ([research §4.3](api-research.md#43-reads-and-other-actions)). Graph read one event's
description in a test, but its other detail fields are unchecked.

**Check** the Graph event and the id conversion
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)):

- Attendees.
- Response status.
- Join link.
- Attachment metadata.

If they hold, hydrate through Graph (`GET /me/events/{id}`), with Outlook Web as the fallback.

### Graph event search

You look for "budget review" meetings in the week of 2 November. `search_calendar_events` uses Search v2:

- It takes no date restriction, so it returns matches from every year.
- It cannot page past offset 1,000.

Today the agent must use `list_calendar_events` for that week with `contains` instead, which scans the
window.

**Check** whether Graph Search (`POST /search/query`, `entityTypes=["event"]`;
[research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)) can:

- Restrict dates on the server.
- Target the description alone.

If so, that request could be one ranked search limited to that week.

### Other calendars

You ask: "What is on my Project Beta calendar next week?" or "What does Alex have on Thursday?" Neither
works today. Only the signed-in user's default calendar is listed and searched. Out of reach:

- Secondary calendars the user created.
- Calendars shared with or delegated to the user.

**Check** with the `outlook.calendar` token
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)):

- The user's own secondary calendars, with its calendar scope.
- Shared and delegated calendars, with its `Calendars.Read.Shared`.

### Reply headers for branch-aware conversations (parked)

Alice sends a proposal. Bob replies to all. Carol forwards it to Dan, who replies to Carol. If Exchange
groups all four, `get_conversation` returns them by date, and Dan's reply reads as if it answered Bob. A
forward that Exchange gives its own conversation id is not returned at all. A conversation is exactly
Exchange's conversation: one flat list, oldest first, with nothing saying who answered whom.

**Check** the quality of these in real mail:

- RFC 5322 `Message-ID`, `In-Reply-To` and `References` headers.
- Exchange's `ConversationId` and `Thread-Index`.

If they hold, building the reply tree becomes an enhancement:

- Detect branches, including forwards. Look across conversation ids where a forward left the
  conversation.
- Offer a merged chronological view, with each message labeled by its branch.

### Event attachment download

A meeting carries `agenda.pdf`. `get_calendar_event` lists it, but `download_attachment` takes only
message ids, so the agent cannot read the file
([research §6.2](api-research.md#62-events)).

**Check** for an event-scoped download route.

### Event to Teams meeting chat

`get_calendar_event` for "Weekly sync" returns a Teams join link but no chat id
([research §6.2](api-research.md#62-events)). The agent then finds the meeting chat in the Teams connector
by name and time, and can pick the wrong chat when two share a name.

**Check** for a supported lookup from the event, or its join link, to the chat. Reading an id out of the
join URL is not established.
