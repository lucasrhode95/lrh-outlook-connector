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

#### Read Exchange's text and list inline images

A newsletter has a table, a quoted reply and two inline charts. `get_message` reads it in one of two ways
today. If the message has inline files, the connector renders the source HTML itself and puts
`[Image: chart.png]` where each chart sits. If not, or if its parser gives up on a construct such as an
Outlook-only conditional comment, it reads Exchange's own text. So two similar messages can read
differently, and the renderer in
[`html_body.py`](../src/outlook_connector/remote/html_body.py) keeps collecting edge cases.

**Change** (decided 2026-11-02):

- Every message reads Exchange's text conversion. Exchange's own image placeholders (`[cid:…]` or
  `[alt text]`) stay where Exchange wrote them, untouched
  ([research §3.6](api-research.md#html-to-text-body)).
- Inline images are listed after the body as `[Image: name]`, like other attachments, with no position.
- Remove the local renderer and its ordered body parts where nothing else needs them. Exports then
  include every inline file the message lists, not only those the selected body references.

The Teams connector keeps placing `[Image: name]` in its text, and that is no inconsistency: Teams serves
message bodies only as HTML, with no server text, so it must read that HTML anyway, and each image is a
typed `<img>` tag. The shared part stays the same in both: the `[Image: name]` marker, safe unique names
and `download_attachment` by name.

### Enhancements

#### Attach local files to drafts

You ask the agent: "Draft a reply to the team and attach `report.pdf`." `create_draft` takes no files
today. The draft is saved without the PDF, and you attach it by hand in Outlook.

**Change:** build the design in
[requirements §12.1](requirements.md#121-drafts-sending-and-signatures):

- An `attachments` list of local paths on `create_draft`.
- After `CreateItem` saves the draft, each file uploads through `CreateAttachmentFromLocalFile` in 14 MiB
  slices, with a `CancellationId` ([research §4.6](api-research.md#46-message-size-limits); decided
  2026-11-02). One path for every size, up to Outlook's own send ceiling.
- Graph read-back confirms each file by name and size.
- A failed upload leaves no half-built draft: the draft moves to Deleted Items and the call reports which
  file failed.

Settle while building where the send limit comes from. Research §4.6 puts the send ceiling between 34 and
35.3 MiB: a 34 MiB file was delivered, and 35.3 MiB bounced.

#### Include deleted mail in conversations

You reply to a thread and then delete the original. `get_conversation` returns only your reply; the original
appears only with `scope={"deleted_items": true}`.

**Change** (decided 2026-11-02): conversations include Deleted Items and Junk by default.
`scope={"deleted_items": false}` still leaves them out. Listing, search, exports by folder or date and
read-state changes keep today's default.

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

None open.

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
