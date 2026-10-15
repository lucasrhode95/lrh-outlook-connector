# Outlook API research

What Microsoft does not document about the services this connector uses, or documents too vaguely to build
on without testing: which clients may do what, Outlook Web's private services (OWS, Cloud Settings, Search
v2), and the places where Microsoft Graph behaved differently from its documentation. Documented Graph
behavior is not repeated here; [Microsoft's Graph documentation](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
is the reference for routes, parameters and schemas.

**Product scope:** [requirements](requirements.md) · **How it is built:** [architecture](architecture.md) ·
**Probes:** [`research/`](../research/README.md)

## Contents

1. [Summary](#1-summary)
2. [Sign-in and clients](#2-sign-in-and-clients)
3. [Microsoft Graph](#3-microsoft-graph): [ids](#31-ids), [mailbox size](#32-mailbox-size-and-listing-cost),
   [counts](#33-counts-and-batches), [mail search](#34-mail-search), [conversations](#35-conversations),
   [attachments](#36-attachments-and-inline-images), [delta](#37-delta-after-moves-and-deletes),
   [calendar](#38-calendar-through-the-teams-client-graph-profile)
4. [OWS](#4-ows-owaservicesvc): [transport](#41-transport), [item writes](#42-item-writes),
   [reads and other actions](#43-reads-and-other-actions), [inbox rules](#44-inbox-rules),
   [saved HTML](#45-what-outlook-keeps-of-saved-html), [message size limits](#46-message-size-limits)
5. [Native signatures](#5-native-signatures-outlook-cloud-settings): [the block in a draft](#51-the-signature-block-in-a-draft),
   [reads](#52-reads), [writes](#53-writes-and-lifecycle), [corporate add-in](#54-corporate-signature-add-in-not-used)
6. [Search v2](#6-search-v2-substrate): [request and paging](#61-request-and-paging), [events](#62-events),
   [mail](#63-mail-conversations-parked), [Teams](#64-teams-messages-and-transcripts),
   [other routes](#65-other-search-routes-seen-in-outlook-web)
7. [Rejected alternatives and identifiers](#7-rejected-alternatives-and-identifiers)
8. [Mailbox side effects of research](#8-mailbox-side-effects-of-research)

**How to read it.** Every finding carries its evidence and date:

| Label | Meaning |
|---|---|
| **Live** | Called by our own process with app-owned tokens; no browser credentials. |
| **Browser** | Seen in an Outlook Web capture taken around a controlled action. The capture is not kept. |
| **Source** | Present in Outlook Web JavaScript. Not proof of use. |

This is an anonymized record of one tested account and tenant, not a capability matrix. Tenant policy, grants,
clients and Microsoft endpoints change; rerun the probes in `research/` before adopting a routing choice
elsewhere, and keep a missing permission, a bad request shape, a transport failure and an empty result apart.
Probes are read-only unless the owner authorized a write for its exact target and content; the writes made
are listed in §8. No tokens, cookies, message text, addresses or real ids are kept here.

## 1. Summary

- **Mail reads → Graph** with the Outlook Mobile client (`Mail.Read`); Graph also reads back OWS changes to
  verify them (§3).
- **Mail changes and inbox rules → OWS**, Outlook Web's private JSON RPC, with the One Outlook Web client.
  The tested clients cannot write mail through Graph, and the Graph profile has no `MailboxSettings.*` (§2, §4).
- **Native signatures → Outlook Cloud Settings**, with the same Outlook profile (§5).
- **Calendar → Graph `calendarView`** with the Teams client's Graph token, which carries `Calendars.*`; OWS
  `GetCalendarView` is the fallback and Search v2 finds events by topic (§3.8, §6.2).
- **Mail search → Graph `$search`**, which matched Outlook's own top-bar search on recall (§3.4).
- **Folders only are cached.** 25.6k items, two thirds in Junk; a full metadata mirror takes about
  12 minutes (§3.2).
- **One id model.** Graph immutable ids survive moves and become OWS ids by a base64 alphabet swap (§4.2);
  `$search` hits are the exception and need `translateExchangeIds` (§3.1).
- **Delta reports a soft delete as `deleted`.** Only a 404 on a GET by id means the item is gone (§3.7).

## 2. Sign-in and clients

Which Microsoft clients are granted which permissions, and the `AADSTS65002` refusals, are in the shared
[sign-in research](https://github.com/lucasrhode95/lrh-m365-auth/blob/main/docs/api-research.md). For this
connector, none of the tested usable clients could write mail through Graph, so mail changes, inbox rules and
signatures use the Outlook profile. The split in
[architecture §6](architecture.md#6-capability-routing-and-portability) follows from that and must be
re-derived for another tenant.

Live, 2026-10-03 and 2026-10-04:

- The Graph token (Outlook Mobile) carried 18 scopes: `Mail.Read`, `Mail.Read.Shared`, `Files.ReadWrite.All`,
  `Sites.ReadWrite.All`, `People.Read(.All)`, `Presence.Read.All`, `User.Read(.All)`, `User.ReadBasic.All` and
  others, but **no `Calendars.*`** and no `MailboxSettings.*`.
- The Outlook token (One Outlook Web) carried 74 Outlook scopes, among them `Mail.ReadWrite(.All/.Shared)`,
  `Mail.Send(.Shared)` and `MailboxSettings.ReadWrite`.
- The two clients do not share a sign-in: each needs its own device code.
- The Search v2 token (One Outlook Web, `https://outlook.office.com/search/.default`) carried
  `SubstrateSearch-Internal.ReadWrite` (2026-10-10).
- The Teams client's Graph token (`teams.graph`) carried `Calendars.Read`, `Calendars.Read.Shared`,
  `Calendars.ReadWrite` and `Calendars.ReadWrite.Shared` (2026-10-12); see §3.8.

Not tested: shared mailboxes, and an Online Archive (the account reports `HasArchive=false`).

## 3. Microsoft Graph

The connector reads folders, messages, conversations, search hits, attachments, MIME and the profile through
Graph with immutable ids requested. Below are only the places where Graph needed something its documentation
does not say.

### 3.1 Ids

Live, 2026-10-03 and 2026-10-04:

- `$search` ignores the immutable-id preference and returns regular ids; a GET by a regular id keeps it regular.
- `translateExchangeIds` from `restId` to `restImmutableEntryId` turns them into the ids listings return
  (3 of 3). Passing an id that is already immutable fails the whole call with 400 `InvalidArgument`, so
  translate only search ids.
- Graph's `changeKey` is accepted as the OWS change key (§4.2).

### 3.2 Mailbox size and listing cost

Live, 2026-10-02 to 2026-10-04, on the tested mailbox:

| Measure | Result |
|---|---|
| Folders | 23, one hidden. Folder delta returns 22; the missing one is hidden, empty and top-level. |
| Items | 25,607: Junk 16,837, Deleted Items 5,170, Sent Items 1,205, Archive 1,036, Inbox 139, 18 others about 1,220. `$count` gives 25,466. |
| Metadata walk (`$top=500`, 17 fields) | 37 records/s, about 1.7 KB each: a full mirror takes about 11.5 minutes and 45 MB |
| Inbox delta bootstrap | 139 records, one page, 1.6 s |

So only folders are cached, and message metadata is read on demand.

**Listing without a folder.** `/me/messages` pages newest first across every folder, Junk and Deleted Items
included (86% of this mailbox), so filtering after paging kept 13 of 100 rows. Counting first (one `$batch` of
per-folder counts for the window) and then listing the in-scope folders that have mail, merged newest first,
filled every page: with 100 per page and no window, 96 messages in 4.7 s (four copies of self-sent mail folded
into one) against 13 in 1.0 s, and a 30-day JSONL export in 11.1 s against 28.3 s. Both methods returned the
same ids in the same order.

### 3.3 Counts and batches

Live, 2026-10-04:

- One `$batch` holding each folder's date-window count and newest message answered for 16 of 16 folders in
  0.9 s.
- No filter excludes meeting mail from a count: `not isof('microsoft.graph.eventMessage')` failed with
  `ErrorInvalidUrlQueryFilter`, and `meetingMessageType` is not a property of `message`.

### 3.4 Mail search

Live, 2026-10-02, three pages of 25 per backend:

| Query | Graph `$search` | Graph `/search/query` | Search v2 (§6.3) |
|---|---|---|---|
| `relatório be` | 75 messages → 52 conversations, more available | total 238 | 75 conversations, total 340 |
| `subject:relatório` (complete) | 63 messages → **50 conversations** | total 65 | **50 conversations**, total 65 |

- All three fold accents (`relatório` matches `relatorio`).
- After resolving Search v2's `ImmutableId`s through Graph, `subject:relatório` returned the identical 50
  conversations from Graph `$search` and Search v2; all 52 `$search` conversations for `relatório be` were in
  Search v2's top 75.
- `/search/query` is not used: its total comes from another engine and ignores the folder rules.
- `$search` takes KQL `attachment:<name>` for attachment file names. Attachment *content* is searched only by
  `/search/query` with `entityTypes=["message"]`, whose hits do not say which attachment matched.
- `/search/query` with `message` found a synthetic marker in its subject and, separately, in its body (2 hits
  each, consistent with the Sent Items and Inbox copies, 2026-10-10).
- Three one-hit `search_messages` calls took 3.4 s, 0.8 s and 0.8 s end to end (2026-10-10). `$search` and
  `/search/query` have not been timed on identical queries.

### 3.5 Conversations

Live, 2026-10-02, 25 recent conversations with 192 messages; and 2026-10-13:

- The `conversationId` filter returns every folder (21 of 25 conversations spanned 2–4 folders).
- Adding `$orderby` answers **400 `InefficientFilter`**, so the connector sorts.
- Graph pages a conversation listing; one-message pages followed to the end returned the same five messages as
  one read (2026-10-13).
- Sizes: median 4, maximum 40 messages.
- Headers: `Message-ID` on 137, `Thread-Index` 136, `In-Reply-To`/`References` 113; 55 messages (drafts and the
  user's own) have none. 6 of 25 conversations branch; 9 parents are not in the mailbox.
- `uniqueBody` is a median 10% of `body`.

### 3.6 Attachments and inline images

Live, 2026-10-02, 40 messages and 176 attachments; and later checks:

- Types: `fileAttachment` 168, `itemAttachment` 8, `referenceAttachment` 0. Sizes: median 11 KB, maximum
  27.5 MB.
- 125 of 168 file attachments are inline. All are referenced as `cid:` in `body`, but only 47 in
  `uniqueBody`.
- Selecting `microsoft.graph.fileAttachment/contentId` in the attachment listing returns every inline file's
  content id, in an ordinary request and inside a `$batch` alike, so no per-image lookup is needed
  (2026-10-08).
- The text body Graph returns shows an inline image as its alt text in brackets at its place
  (`[synthetic pixel]`), and a message whose only attachment is inline reports `hasAttachments=false`
  (2026-10-15). An image without alt text was not checked.

### 3.7 Delta after moves and deletes

Live, 2026-10-02: a delta snapshot taken before the §4.2 changes and read after them.

| Folder | Delta records | Meaning |
|---|---|---|
| Inbox | 2 × `@removed`, `reason: "deleted"` | Both were **soft** deletes, still readable by id in Deleted Items |
| Inbox | 2 × changed | Property edits, and a move out and back |
| Archive | 0 | The round trip netted out: delta reports net state, not history |
| Deleted Items | 2 × added | The same immutable ids, so moves match by id |

### 3.8 Calendar through the Teams client Graph profile

The Outlook Graph profile has no calendar scope; an `event` search with it answered 403 (2026-10-10).
Adding `Calendars.Read` to that first-party client is untested and may fail with `AADSTS65002`. The Teams
client's Graph token already carries `Calendars.*` (§2). Live, read-only, 2026-10-12 and 2026-10-13:

- `GET /me/calendarView` with a window, `$top=100`, selected fields and `Prefer: outlook.body-content-type="text"`
  returned one week in one page: 16 events (11 occurrences, 4 single events, 1 exception), each with attendees
  and a text body.
- **Order:** without `$orderby`, `calendarView` is not in start order, not even within one page. With
  `$orderby=start/dateTime asc` it is, and next links keep that order (page sizes 1, 2, 3 and 100 returned
  the same 16 events, no duplicates).
- Graph's immutable event ids, mapped as in §4.2, equal the ids OWS `GetCalendarView` returns for the same
  week, and hydrate through OWS `GetItem`.
- Start and end times from `GetItem` and from both listings denote the same instants.
- `POST /search/query` with `entityTypes=["event"]` and a word from one event's description returned 12 hits
  in one page; the hydrated hit had the word in its description, not its subject. Whether it can restrict dates
  or search the description alone is untested.
- `GET /me/events/{id}?$select=subject,body` read one event. The other detail fields were not compared.

Not tested: shared or delegated calendars (the token carries `Calendars.Read.Shared`), time zones other than
UTC, and consent on a fresh account.

## 4. OWS (`/owa/service.svc`)

Outlook Web's private JSON RPC, with the One Outlook Web token. Live unless stated.

### 4.1 Transport

- `POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail`.
- Envelope: `{"__type": "<Action>JsonRequest:#Exchange", "Header": {"__type": "JsonRequestHeaders:#Exchange",
  "RequestServerVersion": "V2018_01_08"}, "Body": {"__type": "<Action>Request:#Exchange", …}}`.
- URL-encoded JSON up to 2,048 characters goes in the `X-OWA-UrlPostData` header with an empty body; longer
  payloads go in the body.
- Headers: the bearer token, `Action`, `X-OWA-ActionSource`, correlation ids,
  `X-AnchorMailbox: AAD-SMTP:<user>`, `Prefer: IdType="ImmutableId"`. **No cookies or canary.**
- Answers: `Body.ResponseMessages.Items[]` with `ResponseClass` (`Success`, `Warning`, `ClientError`,
  `ServerError`, `ServerTransientError`, `Throttled`) and `ResponseCode`; some reads answer directly in `Body`.
  Error headers: `x-owa-error`, `x-owa-returncode`.

### 4.2 Item writes

Live, 2026-10-02 to 2026-10-06: a self-send, four user-named Inbox messages and synthetic drafts; every result
read back through Graph.

| Operation | Request | Result |
|---|---|---|
| Ids | A Graph immutable id with `-`→`/` and `_`→`+` is the OWS `ItemId`; the same for `conversationId` | Accepted on all, and ids survive moves |
| New mail | `CreateItem`, `MessageDisposition: "SendAndSaveCopy"`, `ComposeOperation: "newMail"` | Sent; one copy in Sent Items |
| Draft | `CreateItem`, `SaveOnly`; inline images as `Attachments` of `FileAttachment:#Exchange` (`Content` base64, `ContentType`, `ContentId`, `IsInline`) | Saved in Drafts with its inline attachments (2026-10-07) |
| Reply with history | `CreateItem` `ReplyToItem` / `ReplyAllToItem` with `NewBodyContent` `BodyType: "HTML"`, `SaveOnly` | The draft quotes the original's HTML (lists, tables, bold) and carries its inline images byte for byte. A `Text` body makes Exchange flatten the quoted original. |
| Send an existing draft | `SendItem` | **Not supported:** HTTP 500 `OwaOperationNotSupportedException`, nothing sent |
| Send an existing draft unchanged | `UpdateItem` with `Updates: []`, the draft's Graph `changeKey` as `ChangeKey`, `ConflictResolution: "NeverOverwrite"`, `SendAndSaveCopy`, `SavedItemFolderId` Sent Items | Sent, copy in Sent Items. If the draft changed after it was read: `ErrorIrresolvableConflict` (also with `AutoResolve`), nothing sent, the draft kept. |
| Clear Cc/Bcc | `UpdateItem` `SaveOnly`, `SetItemField` `message:CcRecipients` / `message:BccRecipients` with an empty list | Both read back empty |
| Read state | `UpdateItem`, `SetItemField` `message:IsRead`, `ConflictResolution: "AlwaysOverwrite"`, `SaveOnly`, `SuppressReadReceipts: true` | Done |
| Flag | `SetItemField` `item:Flag`, `Flag: {"__type": "FlagType:#Exchange", "FlagStatus": "Flagged" or "NotFlagged"}` | Done |
| Categories | `SetItemField` `item:Categories` | Done |
| Conversation read state | `ApplyConversationAction`, `Action: "SetReadState"`, `ContextFolderId` a distinguished folder | Done (the connector sets it per message instead) |
| Move | `MoveItem`, `ToFolderId` a `DistinguishedFolderId` or folder id, `ReturnNewItemIds: true` | Done; the Graph immutable id is unchanged |
| Soft delete | `DeleteItem`, `DeleteType: "MoveToDeletedItems"` | In Deleted Items under the same id |

A message saved through OWS reads back through Graph with a trailing CRLF on its text (2026-10-08).

### 4.3 Reads and other actions

- Proven but unused for ordinary mail data (2026-10-01): `FindFolder`, `GetFolder`, `FindItem`,
  `FindConversation`, `GetConversationItems`, `GetTimeZone`. `ExecuteSearch` answered 400 for every variant
  tried.
- `GetItem` hydrates calendar events: `BaseShape: "Default"`, `BodyType: "Text"` and `AdditionalProperties`
  `item:Body` and `item:Attachments` return the agenda, attachment metadata and required and optional attendees
  (2026-10-10). The default shape caps the body at 2,048 characters and says so in `Body.IsTruncated`. The text
  body is not HTML-entity-decoded.
- `GetCalendarView` with a UTC `TimeZoneContext` returns a window's events directly in `Body`, without
  descriptions, and the same events as Graph's `calendarView` (§3.8).
- `GetOwaUserConfiguration` answers with null legacy signature contents; it is not the modern signature list
  (§5.2).
- Browser: `FindConversation` rows carry cross-folder `GlobalItemIds` and `GlobalMessageCount`.
- **File attachments (Browser, 2026-10-07; untested from our process):** Outlook Web uploads with
  `POST /owa/service.svc/CreateAttachmentFromLocalFile`. The metadata travels as URL-encoded JSON in
  `X-OWA-UrlPostData` (`CreateAttachmentJsonRequest:#Exchange` → `CreateAttachmentRequest:#Exchange`:
  `ParentItemId {Id, ChangeKey}`, `Attachments: [{"__type": "FileAttachment:#Exchange", "Content": "",
  "ContentType", "IsInline", "Name", "Size"}]`, `IncludeContentIdInResponse: true`, `SliceNumber: 0`,
  `TotalSlices: 1`); the bytes are the request body. The answer gives `AttachmentId`, `ContentId`,
  `LastModifiedTime` and `IsInline`.
- Unusable: the B2 route `/messageservice/ows/…` answered 500. `/outlookgatewayb2/graphql` is Outlook's own
  GraphQL, not Microsoft Graph.

### 4.4 Inbox rules

Graph's `messageRules` needs `MailboxSettings.*`, which the Graph profile lacks (§2). Over OWS, the EWS name
`GetInboxRules` answers `OwaOperationNotSupportedException`; Outlook Web's own rule actions work. Captured in
Outlook on the web on 2026-10-04 (create, edit, disable, enable, reorder and delete throwaway rules), then
replayed from our process the same day: every action works as captured.

**A different envelope.** Rule actions send the request object itself, without the `…JsonRequest` wrapper and
`Body`:

```json
{"__type": "GetInboxRuleRequest:#Exchange",
 "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08"},
 "UseServerRulesLoader": true}
```

The answer has no `ResponseMessages`: it carries `WasSuccessful`, `ErrorCode` (0 on success), `ErrorMessage`,
`UserPrompt` and `IsUserError`. URL, headers and the `X-OWA-UrlPostData` rule are as in §4.1. A
`TimeZoneContext` in the header is not needed.

| Action | Request (besides `__type` and `Header`) | Answer |
|---|---|---|
| `GetInboxRule` | `UseServerRulesLoader: true` | `InboxRuleCollection.InboxRules`, in priority order (`Priority` 1 first) |
| `NewInboxRule` | `InboxRule` with `Name` and only the conditions and actions used, for example `From` / `SentTo` as `[{"__type": "PeopleIdentity:#Exchange", "DisplayName", "SmtpAddress", "RoutingType": "SMTP"}]`, `SubjectContainsWords`, `MoveToFolder: {"DisplayName", "RawIdentity": <folder id>}`, `MarkAsRead`, `StopProcessingRules` | The created rule, `Priority` 1 (a new rule goes first) |
| `SetInboxRule` | `InboxRule` with `Identity: {"DisplayName": <name>, "RawIdentity": <rule id>}` and the fields to set; `Force: false` | `WasSuccessful` |
| `EnableInboxRule`, `DisableInboxRule`, `RemoveInboxRule` | `Identity` only | `WasSuccessful` |
| `SetInboxAndSweepRules` | `EnableDisableInboxRules`: one `{"__type": "EnableDisableInboxRuleRequest:#Exchange", "Header", "Identity", "IsEnabled"}` per rule, **every rule, in the new order** | `WasSuccessful`; priorities renumbered |

- **Rule id:** `<mailbox GUID>\<20-digit number>`, stable across edits and renames (a rename sends the new name
  in `Identity.DisplayName` with the same `RawIdentity`).
- **Edits are partial:** fields left out are kept; a field sent as `null` is cleared. Outlook always resends
  `Name`, the sender or recipient condition, `StopProcessingRules` and `Identity`.
- **Order** has no `Priority` field: it is the order of `SetInboxAndSweepRules`, which also carries each
  rule's `IsEnabled`.
- **Folders:** a request's `MoveToFolder.RawIdentity` takes the Graph folder id as is. An answer gives a
  different form, `<organisation path>/<mailbox GUID>:\<folder name>`, so a rule's target is matched by name.
- **A rule has 90 fields** (conditions, `ExceptIf…` exceptions, actions, descriptions, `InError`,
  `SupportedByTask`, `RuleProvider`). Inactive ones are not null but carry sentinel values
  (2026-10-07): `NullInboxRuleMessageFlag` (`FlaggedForAction`, `RequestedAction` and exceptions),
  `NullInboxRuleMessageType` (`MessageTypeMatches`), `NullImportance` (`MarkImportance`, `WithImportance`),
  `NullSensitivity` (`WithSensitivity`), plus `DescriptionTimeFormat` and `DescriptionTimeZone` metadata.
- **Creation reports a different identity:** the `RawIdentity` in the `NewInboxRule` answer did not match the
  same rule in a fresh `GetInboxRule` (2026-10-07). Find a created rule by reading the list back.
- Outlook Web's toggle can send `DisableInboxRule` twice; harmless. It also calls `GetMailboxByIdentity` and
  `/ows/v1.0/OutlookOptions/MailForwardingNotification`, neither needed for rules.

### 4.5 What Outlook keeps of saved HTML

Live, 2026-10-07 and 2026-10-13, synthetic messages sent to the owner's mailbox and an authorized personal
address; read back through Graph and inspected in Outlook Web, Gmail and (by the owner) Outlook desktop:

- A fragment and the equivalent full document saved to the same 548-character HTML.
- Unclosed paragraphs, lists and tables were repaired; all text kept.
- Inline styles, `<head>` style rules and classes persisted; Outlook Web's dark theme changes displayed colors.
- A remote HTTPS image kept its URL and loaded in both clients (Gmail through its image proxy); no attachment
  was made.
- Tables, lists, links, accented and Japanese text survived draft, Sent Items and received copies.
- An external copy gained the organization's disclaimer, absent from the work-mailbox copies: a transport
  footer.
- **Unfinished markup** (2026-10-13, one synthetic draft never sent): a body ending in an unfinished `img` tag
  with an `onerror` attribute, followed by a signature block. Outlook dropped the `onerror` attribute and the
  script text, but the unfinished tag absorbed the signature's opening `div`: its `id` and
  `data-signature-name` moved onto the `img`, and the signature's closing tag became an empty `div`.
- A signature image whose bytes were a valid PNG rendered and saved in Outlook ("Save as Picture",
  2026-10-10); an earlier red X came from a malformed 26-byte PNG, not from a missing image.

### 4.6 Message size limits

Live, 2026-10-15: synthetic self-drafts, each carrying non-inline `FileAttachment:#Exchange` files of random
bytes (`IsInline: false`), saved through the connector's own OWS client. Two of them were sent to the
owner's address.

- **Request size.** One OWS request body holds about 30,000,000 bytes. A `CreateItem` carrying a 21.3 MiB
  file (about 29.8 MB of base64) saved. One carrying 21.5 MiB (about 30.1 MB) failed with **HTTP 500**,
  `Microsoft.AspNetCore.Http.BadHttpRequestException` or no error code at all, and saved nothing (checked
  through Graph). The 500 looks like an unknown outcome, but nothing was saved: the request was refused
  before Exchange read it. Larger sizes (23, 30 MiB) failed the same way.
- **Adding files to a draft.** `CreateAttachment` with `ParentItemId {Id, ChangeKey}` and `Attachments: [...]`
  (the same `FileAttachment` shape, `Content` base64) adds a file to an existing draft and returns the new
  `RootItemChangeKey` in `Attachments[0].AttachmentId`. The same request-size cap applies to each call.
- **Stored size.** Adding 20 MiB to a 21.3 MiB draft failed with `ErrorMessageSizeExceeded` ("The message
  exceeds the maximum supported size.", "Cannot save changes made to an item to store."). The draft kept its
  earlier state. Adding 14 MiB (35.3 MiB in total) saved. Graph's stored size (`Integer 0x0E08`) matched the
  file sizes.
- **Send.** Sending the 35.3 MiB draft (`UpdateItem` with `SendAndSaveCopy`) was accepted at once. Two seconds
  later a non-delivery report reached the Inbox: too large, "The limit is 35 MB. Your message is 36 MB." A
  copy stayed in Sent Items. The send limit is therefore learned from a later NDR, not from the send call.
  A 1 MiB attachment sent and arrived normally.
- **MailTips.** Graph `POST /me/getMailTips` with `MailTipsOptions: "maxMessageSize"` for the user's own
  address returned 37,748,736 bytes (36 MiB), using the Graph mail profile. That figure matches the
  stored-size limit rather than the 35 MB that the NDR states.

## 5. Native signatures (Outlook Cloud Settings)

Graph has no documented signature API
([mailboxSettings](https://learn.microsoft.com/en-us/graph/api/resources/mailboxsettings?view=graph-rest-1.0),
[roaming signatures](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-organizationconfig?view=exchange-ps#-postponeroamingsignaturesuntillater)),
so signatures use Outlook Cloud Settings with the Outlook profile. No extra scopes or vendor tokens are needed.

### 5.1 The signature block in a draft

Source, from the Outlook Web bundles of 2026-10-07 (minified names change between releases; the constants
below are what matter):

- The block is a client convention: an element with id `Signature`, selector `#Signature` and the attribute
  `data-signature-name`, which the editor reads into its model and writes back. Outlook Web's own drafts
  carried `<div id="Signature">`.
- The current signature is the **first** `Signature` container. Inserting a signature replaces that block in
  place; without one it goes at the start of the body for new mail and replies, else at the cursor, wrapped as
  `<div id='Signature'><div>{html}</div></div>`. Plain-text bodies get it after two newlines.
- A `Signature` container left with one empty paragraph loses its id.
- Received HTML has its ids prefixed `x_` (`x_Signature`). In a reply, the prefix is stripped only before the
  reply header (`divRplyFwdMsg`), so the quoted sender's signature never matches.
- `<div id=appendonsend></div>` is a separate placeholder for content appended on send.

### 5.2 Reads

`GET /ows/v1/OutlookCloudSettings/settings/` with headers `x-outlook-client: owa` and
`Accept: application/json`. Live, 2026-10-07:

| Purpose | Request | Result |
|---|---|---|
| Names and defaults | `?settingname=roaming_signature_list,roaming_new_signature,roaming_reply_signature` | Three settings: the list `type: BlobArray` (comma-joined names), the new-message and reply defaults `type: String` |
| Contents | `?settingname={URL-encoded name}`, header `x-islargesetting: true` | One entry per format by `secondaryKey`: `htm`, `txt` and sometimes `rtf` |

- Entries carry `name`, `value`, `type`, `source`, `scope` (the account), `metadata`, `parentSetting`,
  `Timestamp`, `itemClass`, `id` and `secondaryKey`; large ones also `nameBase64Encoded`.
- Repeated reads returned identical values and timestamps.
- A listed name can return `[]`: a leftover reference after a deletion. Resolve contents before trusting a
  name.
- Adding a signature does not change the defaults.
- An image in a signature is stored inline as a `data:image/png;base64,…` source (61,483-byte PNG in one test).
- Variants that failed: a name in the path (`/settings/{name}`) answered 404; `parentsetting=…` answered 403.

### 5.3 Writes and lifecycle

Live, 2026-10-07 and 2026-10-08, with synthetic signatures only; originals verified unchanged afterwards.

| Action | Request | Result |
|---|---|---|
| Create or update | `PATCH /ows/v1/OutlookCloudSettings/settings/account`, JSON array of per-format settings, `x-islargesetting: true` | 200; readable at once and added to the name list |
| Set or clear a default | Same route, `String` settings, `x-islargesetting: false` | 200; new-message and reply defaults change independently; an empty value clears one |
| Delete | `DELETE /ows/v1/OutlookCloudSettings/settings/account`, body `{"name": …}`, `Content-Type: application/json`, `x-islargesetting: true` | 200; contents `[]`, name gone from the list |

- A content setting: `itemClass: "RoamingSetting"`, `name`, account `scope`, `secondaryKey` `htm` or `txt`,
  `type: "Blob"`, `value`, `parentSetting: "roaming_signature_list"`, `metadata: "encoding:utf-8"`, a UTC .NET-ticks
  `timestamp`, `value@is.Large: true`. A default: `name` and `secondaryKey` both `roaming_new_signature` or
  `roaming_reply_signature`, `type: "String"`. Headers also `x-outlook-client: owa`, `x-overridetimestamp: false`.
- A DELETE without the JSON content type answered 415; with only a query name, 400.
- Saved HTML is normalized: Outlook adds wrappers and CRLFs and turns `id="Signature"` into `id="x_Signature"`.
  Visible text, the text format and embedded PNG bytes are kept exactly.
- **Deleting the selected default leaves the pointer** on the missing name. A reader must check that the
  default's contents exist and report it missing, not fall back.
- **Names are exact, case-sensitive keys.** Writing the same content under a new name makes a second entry;
  defaults do not follow a rename. Spaces, accents, `+` and `&` round-trip. A name with a comma answered
  `InvalidRequest` (the list is comma-joined).
- An image signature placed in a draft as a CID with an inline attachment kept the reference and byte-identical
  image through Graph read-back and sending (2026-10-08, 2026-10-10).

Wire-format leads came from the Outlook Web source and an
[open HTTP implementation](https://gist.github.com/EionRobb/21dd89b05417b247bface6373fb380fb); the calls above
establish the behavior for this account.

### 5.4 Corporate signature add-in (not used)

Browser and Source, 2026-10-07. The organization's officeatwork Mail Signature add-in supplies its own signature
in the compose client: the HTML is in Outlook Web's first `CreateItem` request, and OWS does not generate it.
The add-in renders a Nunjucks template package (`.ofawmsig` from a SharePoint library) from the sender's
profile, language, audience and recipients, uploads its images with `CreateAttachmentFromLocalFile` (§4.3), and
inserts the result with Office.js `body.setSignatureAsync`. Its selected template is stored in add-in roaming
settings (`SaveExtensionSettings`), not as native signature HTML, and `UserConfigurations` writes in the same
capture held unrelated add-in preferences. Because its output depends on recipients, it has no stored
equivalent; drafts created by the connector are not changed by it (owner, 2026-10-07). A `GetItem`
`NormalizedBody` shows images as 42-byte placeholder GIFs, so attachments, not that body, are the source of
image bytes.

## 6. Search v2 (Substrate)

`POST https://outlook.office.com/searchservice/api/v2/query` with the Search v2 token (§2) and
`X-AnchorMailbox: Oid:<object id>@<tenant id>`. Live from our process; no browser cookies or tokens.
The `teams.microsoft.com` host answered 405 for this path.

### 6.1 Request and paging

```json
{"Cvid": "<guid>", "LogicalId": "<guid>", "Scenario": {"Name": "owa.react"}, "TimeZone": "UTC",
 "TextDecorations": "Off",
 "EntityRequests": [{"EntityType": "Event", "ContentSources": ["Exchange"],
   "Query": {"QueryString": "<text>", "DisplayQueryString": "<text>"},
   "From": 0, "Size": 25, "Sort": [], "PropertySet": "Optimized"}]}
```

- The response is `EntitySets[].ResultSets[]` with `Total`, `MoreResultsAvailable` and `Results[].Source`.
- **Reuse `Cvid` and `LogicalId` across the pages of one search.** A new pair per page answered 400 after the
  first page.
- **Offset limit** (2026-10-13): one session answered pages starting up to 975 and refused a page starting at
  1,000 with 400; a page starting at 990 with 25 rows was answered, so the limit is on where a page starts.
- **No date restriction:** `start:`, `meeting start:` and `meeting received:` terms returned no result set
  (2026-10-13).
- `ContentSources` is required: a Teams message request without it answered 400.

### 6.2 Events

`EntityType: "Event"`, `ContentSources: ["Exchange"]` (2026-10-10 and 2026-10-12):

- A broad query returned 25 rows a page, `Total=703` and `MoreResultsAvailable=true`; offset 5 with the same
  `Cvid`/`LogicalId` returned five more.
- `Source` fields: `Subject`, `Start`, `End`, `Organizer`, `DisplayTo`, `Location`, `IsMeeting`, `IsCancelled`,
  `IsAllDayEvent`, `IsOrganizer`, `IsResponseRequested`, `MyResponseType`, `LegacyFreeBusyStatus`,
  `Sensitivity`, `HasAttachments`, `ItemId.Id`, `ImmutableId`, `UID` and series-master ids. Requesting
  `Fields` did not add `OnlineMeeting` or `EventConversationId`.
- **`Total` is unreliable for recurring events:** one series returned ten occurrence rows over two pages while
  each response said `Total=1`.
- **`HasAttachments` is unreliable:** rows said `false` for an event whose `GetItem` returned an attachment.
- `Source.ItemId.Id` hydrates through OWS `GetItem` (§4.3), which returns the agenda, attendees, attachment
  metadata and a Teams join URL, but no Teams chat id.
- A minimal `AnswerEntityRequests` for events answered 500.

### 6.3 Mail conversations (parked)

`EntityType: "Conversation"`, `ContentSources: ["Exchange"]` works (2026-10-02). `Source` carries
`ConversationId` (a different encoding from Graph's), `ImmutableId` (readable through Graph), `ParentFolderId`,
`From`, `UniqueSenders`, `MessageCount`, `Preview` and `HasAttachments`; the result set also has
`TotalWithoutCollapsing`. It matched Graph `$search` on recall (§3.4), so mail search stays on Graph.

### 6.4 Teams messages and transcripts

The same token answers `EntityType: "Message"`, `ContentSources: ["Teams"]` (2026-10-10). The Teams connector
uses this route; its findings are in the
[Teams research](https://github.com/lucasrhode95/lrh-teams-connector/blob/main/docs/api-research.md#message-search).
`EntityType: "Transcript"` with `ContentSources: ["Teams"]` answered 200 with no `EntitySets`; with Exchange,
OneDrive or SharePoint sources, or as `MeetingTranscript` or `MeetingInsights`, it answered 400. Transcript
search is neither shown nor ruled out.

### 6.5 Other search routes seen in Outlook Web

Browser and Source, 2026-10-10:

- Outlook Web's global search submitted `POST /searchservice/api/v2/query` for `Conversation`/`Exchange` with
  `From=0`, `Size=25`. Search-box suggestions are separate (`/search/api/v1/suggestions`, `People`, `File`,
  `Event`, `Message`), as are `/search/api/v1/recommendations`, `/search/api/v2/init`,
  `/search/api/v2/resources` and the `/search/api/v1/events` telemetry.
- Every captured offset was 0; no continuation request was captured.
- Persona calls beside search (`livepersonacard/configuration`, `peopleInsights/owareadingpane`,
  `personaphoto`, `personacards/preparePersona`) are profile support, not search.
- Outlook Web also calls its own GraphQL (`/outlookgatewayb2/graphql`) for calendar groups, tabs, working
  hours, auth settings and meeting responses.

## 7. Rejected alternatives and identifiers

- **Browser-token tools** ([outlook-cli](https://github.com/yusufaltunbicak/outlook-cli),
  [owa-piggy](https://github.com/damsleth/owa-piggy)): useful route vocabulary, unacceptable authentication.
- **Outlook REST v2:** retired; see Microsoft's [Graph comparison](https://learn.microsoft.com/en-us/outlook/rest/compare-graph).
- **`GetAccessTokenForResource` over OWS:** a historical token-minting path; do not use
  ([Black Hat 2019](https://i.blackhat.com/USA-19/Thursday/us-19-Jaiswal-Preventing-Authentication-Bypass-A-Tale-Of-Two-Researchers.pdf)).
- **EWS:** follow Microsoft's [retirement guidance](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online).
- The Exchange resource app id `00000002-0000-0ff1-ce00-000000000000` is not a client id.
- `ab0455a0-8d03-46b9-b18b-df2f57b9e44c` appears in Outlook Web's consumer-auth configuration; unverified, do
  not use.
- Outlook Web deep links (Browser): `/mail/inbox/id/<id>` and `/owa/?viewmodel=ReadMessageItem&ItemID=<id>`.

## 8. Mailbox side effects of research

What research left in the tested mailbox. Nothing was permanently deleted, no write was retried, and no probe
used a browser cookie, token or canary.

| Date | Writes |
|---|---|
| 2026-10-01 | One synthetic self-send. |
| 2026-10-02 | On four user-named Inbox messages: read and flag toggles, a category set and cleared, a conversation read toggle, a move round trip, two soft deletes (two messages remain in Deleted Items). |
| 2026-10-04 | One throwaway inbox rule created, edited, renamed, disabled, enabled, reordered and removed; the rule list ended as it started. |
| 2026-10-07 | Four test messages sent, with personal recipients as authorized; sent and received copies remain. Five HTML test messages sent to the work mailbox and Gmail; copies remain, read states restored. 105 drafts moved to Deleted Items. One throwaway rule created and removed; the original rules unchanged. Four synthetic signatures created and deleted; defaults switched and restored. |
| 2026-10-08 | Six synthetic signatures created and removed, defaults restored. One synthetic self-send. Synthetic drafts moved to Deleted Items. |
| 2026-10-10 | One synthetic self-send for the signature-image check; one copy remains in Inbox and one in Sent Items. One synthetic self-send for the search checks. |
| 2026-10-13 | One synthetic draft with unfinished markup, never sent, moved to Deleted Items. |
| 2026-10-15 | Size-limit check (§4.6): three synthetic self-drafts with random-byte attachments saved. Two were sent to the owner's address (1 MiB delivered, 35.3 MiB bounced with an NDR); the third draft, both Sent Items copies, the received copy and the NDR were moved to Deleted Items. Four oversized saves were refused and left nothing. |
