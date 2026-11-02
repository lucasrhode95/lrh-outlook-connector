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

#### Hydrate events through Graph

You ask the agent whether you accepted Thursday's review, and to read its attached `agenda.pdf`.
`get_calendar_event` reads through Outlook Web (`GetItem`). It returns no response of yours, and the PDF is
listed but cannot be downloaded: `download_attachment` takes only message ids.

Graph does both. `GET /me/events/{id}` takes every event id the connector returns, after the usual id swap.
It matched Outlook Web on attendees, join links and text bodies, and it adds your response
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)).

**Change:**

- Hydrate through Graph under `outlook.calendar`, with no Outlook Web fallback. Remove the `GetItem` event
  read in [`ows_calendar.py`](../src/outlook_connector/remote/ows_calendar.py).
- List an event's attachments from `/me/events/{id}/attachments`; Graph's `hasAttachments` can say `false`
  for an occurrence that has a file.
- Let `download_attachment` take an event id, through `/me/events/{id}/attachments/{id}/$value`.

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

- Where the send limit comes from. [Research §4.6](api-research.md#46-message-size-limits) puts the send
  ceiling between 34 and 35.3 MiB: a 34 MiB file was delivered, and 35.3 MiB bounced.
- Whether to upload every file with `CreateAttachmentFromLocalFile` in 14 MiB slices. They worked live: a
  34 MiB file went up in three slices and read back byte for byte. That reaches Outlook's own send ceiling,
  while base64 in `CreateItem` stops near 21 MiB per file.

#### Secondary calendars

You ask: "What is on my Project Beta calendar next week?" `list_calendar_events` reads only the default
calendar, so the answer is empty. Graph lists your other calendars and reads each one's window, with
descriptions ([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)).

**Change:** list calendars from `GET /me/calendars`, and let `list_calendar_events` take a calendar,
read through `/me/calendars/{id}/calendarView`.

#### Reply tree for branch-aware conversations (parked)

Alice sends a proposal. Bob replies to all. Carol forwards it to Dan, who replies to Carol. `get_conversation`
returns them as one flat list, oldest first, so Dan's reply reads as if it answered Bob.

The headers hold up: `In-Reply-To` always equals the last `References` entry, and in a sample of recent mail
16 conversations had two answers to the same parent
([research §3.5](api-research.md#35-conversations)). Exchange splits off notification chains whose subject
changes, even when their headers link them.

**Change**, when it is unparked:

- Build the reply tree from `In-Reply-To`, across conversation ids where a forward or a subject change
  left the conversation.
- Offer a merged chronological view, with each message labeled by its branch.

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

### Meeting chat from the join link

`get_calendar_event` for "Weekly sync" returns a Teams join link but no chat id. The agent finds the meeting
chat in the Teams connector by name and time, and can pick the wrong chat when two share a name.

No supported lookup works with the connector's tokens: Graph's `onlineMeetings` answers 403. The join link
holds a `19:meeting_…@thread.v2` thread id, and in 54 of 56 tested links it equaled the meeting chat's id
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)). Microsoft does not
document that format.

**Decide:** expose that thread id as the event's likely Teams chat id, labeled as read from the link, or keep
the chat lookup to names and times.

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

Cover their generated CIDs, native quotation and `uniqueBody` behavior. Also check whether other
deployments' mail filters expose or duplicate hidden previews.

### Fresh-account consent

A second user installs the connector and runs `outlook-connector auth` for the first time. The calendar
grants have only been checked on one account. If the new account is refused:

- `m365.search`: `search_calendar_events` fails.
- `outlook.calendar` (the Teams client): `list_calendar_events` fails.

**Check** both grants on a new account before relying on them.

### Shared and delegated calendars

You ask: "What does Alex have on Thursday?" Alex shared a calendar with you. Only your own calendars are in
reach today. Your secondary calendars read fine through Graph
([research §3.8](api-research.md#38-calendar-through-the-teams-client-graph-profile)), but the tested
mailbox had no calendar shared by or delegated from someone else.

**Check**, on a mailbox that has one, with the `outlook.calendar` token's `Calendars.Read.Shared`:

- A calendar another user shared, as it appears in `GET /me/calendars`.
- A delegate's read of another user's calendar through `/users/{address}/calendar`.
