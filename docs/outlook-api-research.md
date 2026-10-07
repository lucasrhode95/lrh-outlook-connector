# Outlook API research

**Evidence dates:** 2026-10-01 (first standalone probes) · 2026-10-02 (browser capture review, probe suite) ·
2026-10-04 (inbox-rule replay, H7 live check, read-only probes) · 2026-10-06 (version-bound draft sends) ·
2026-10-07 (connector HTML/replies, inbox-rule blockers, bulk/partial failures, signature capture analysis)

**Product scope:** [Requirements v4](outlook-requirements-v4.md) · **How it is built:** [Architecture](architecture.md) · **Probes:** [`research/`](../research/README.md)

This is the single record of what Microsoft's APIs allow and how they behave for this tenant (Landis+Gyr, one user mailbox). It keeps evidence and conclusions only. How to re-run the probes and take captures is in `research/README.md`.

| Label | Meaning |
|---|---|
| **STANDALONE** | Called by our own process with app-owned tokens. No browser credentials. |
| **BROWSER** | Seen in an Outlook Web capture. Request/response observations are identified separately where material. |
| **SOURCE** | Present in Outlook Web JavaScript. Not proof of use. |

## 1. Summary

- **Reads → Microsoft Graph** via the Outlook Mobile client (`Mail.Read`). Every read the product needs is proven (§3).
- **Writes → OWS** (Outlook Web's private JSON RPC) via the One Outlook Web client. Graph mail write and send scopes are denied to every usable client (§2), so OWS fills exactly that gap. Send and all mutations are proven (§4).
- **Search → Graph `$search`.** It matched Outlook's own top-bar search (Substrate) on recall. Substrate works with our token, but it is parked (§3.3, §5).
- **Cache folders only.** The mailbox has 25.6k items, two thirds of them Junk, and a full metadata mirror takes about 12 minutes (§3.2).
- **The id model is simple.** Graph immutable ids survive moves, and they become OWS ids by a base64 alphabet swap (§4.2). Exception: `$search` returns regular ids; they need `translateExchangeIds` (§3.1).
- **Delta `@removed` with `reason: "deleted"` is also reported for soft deletes.** Resolve it with a GET by id. Only a 404 means the item is gone (§3.6).

## 2. Authentication (STANDALONE)

Device-code sign-in against `login.microsoftonline.com/organizations`, then refresh tokens. A client's refresh token can request other resources it is preauthorized for. `.default` reveals the full preauthorized Graph scope set.

| Client | Resource / scope | Result |
|---|---|---|
| **C** Outlook Mobile `27922004-5251-4030-b22d-91ecd9a37ea4` | Graph `Mail.Read` (`.default`: 18 scopes) | **Granted.** The only mail scopes are `Mail.Read` and `Mail.Read.Shared`. Silent refresh works. The set also includes `User.Read` (checked 2026-10-04): `/me` and `/me/photos/48x48/$value` work with this sign-in, so the UI shows the user's name and photo without new permissions. |
| C | Graph `Mail.Send`, `Mail.ReadWrite`, `Mail.ReadWrite.Shared` | **Denied: AADSTS65002** |
| **A** One Outlook Web `9199bf20-a13f-4107-85dc-02114787ef48` | `https://outlook.office.com/.default` | **Granted.** Includes `Mail.ReadWrite(.All/.Shared)` and `Mail.Send(.Shared)` for the Outlook resource. |
| A | `https://outlook.office.com/search/.default` | **Granted** silently from A's refresh token. Audience `66a88757-258c-4c72-893c-3e8bed4d6899`, scope `SubstrateSearch-Internal.ReadWrite`. This is the same client and scope Outlook Web uses (BROWSER). |
| A | Graph `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` (`.default`: 26 scopes, none `Mail.*`) | **Denied: AADSTS65002** |
| **B** Outlook desktop / M365 `d3590ed6-52b3-4102-aeff-aad2292ab01c` | Graph `Mail.Read` | **Denied: AADSTS65002** |
| D older Outlook Web `bc59ab01-8403-45c6-8796-ac3ef710b3e3` | — | Not tested |

- [`AADSTS65002`](https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes) means the client is not preauthorized for that resource. It is final for that client/scope pair. Denied pairs are listed in `research/probes/common.py` `DENIED` and are never requested again.
- **Conclusion:** no usable client can write mail through Graph. The Graph/OWS split in [architecture §6](architecture.md) follows from this table and must be re-derived for another tenant.

## 3. Microsoft Graph: reads (STANDALONE, client C)

Every call sent `Prefer: IdType="ImmutableId"`.

### 3.1 Proven routes

| Route | Result / note |
|---|---|
| `/me/mailFolders` (+ recursive `childFolders`, `includeHiddenFolders=true`) and 11 well-known aliases | 200. `archive` is the normal Archive folder, not Online Archive. |
| `/me/messages`, `/me/mailFolders/{id}/messages` with `$select`, `$filter` (date, sender), `$top`, `$orderby` | 200. `$orderby` requires the ordered property to appear first in `$filter` (otherwise `InefficientFilter`). |
| `/me/messages?$count=true` with `ConsistencyLevel: eventual` | 200 |
| `GET /me/messages/{id}`: text and HTML body, `uniqueBody`, `internetMessageHeaders` | 200. Reading does not change `isRead`. |
| `/me/messages/{id}/$value` (MIME) | 200 |
| `/me/messages?$filter=conversationId eq '…'` | 200, across all folders (§3.4) |
| `$search` (`subject:`, `from:`, `to:`, body, `attachment:`, `received:`) | 200. Field terms must be quoted. **Ignores `Prefer: IdType="ImmutableId"`** (live 2026-10-03): hits carry the regular id (`AQMk…`), not the immutable one (`AAkALg…`) that listings return; a GET by a regular id returns it unchanged (the header does not convert it; live 2026-10-03), but `POST /me/translateExchangeIds` (`restId` → `restImmutableEntryId`) does, with the read sign-in. Outlook Web accepted the regular id as a reply target. |
| `POST /search/query` (message entity) | 200. Gives a server `total`. |
| Attachments: list, item (`$select` must not include `@odata.type`), `$value`, `$select=microsoft.graph.fileAttachment/contentId` | 200. The cast also works on the list: `/me/messages/{id}/attachments?$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId` returns each file attachment's `contentId` in one request (live 2026-10-04, §3.5). |
| `/me/mailFolders/{id}/messages?$count=true&$top=1&$select=receivedDateTime&$orderby=receivedDateTime desc&$filter=receivedDateTime ge …` in one `$batch`, `ConsistencyLevel: eventual` | 200 for 16 of 16 folders in one request (0.9 s): each sub-response carries `@odata.count` and, when the count is not 0, the newest message's `receivedDateTime` (5 of 5). One batch gives each folder's count and newest date (live 2026-10-04, H30). |
| `$filter` by message type, for counts without meeting mail | **400** (live 2026-10-04, H28): `not isof('microsoft.graph.eventMessage')` → `ErrorInvalidUrlQueryFilter`; `meetingMessageType eq 'none'` → property not found on `microsoft.graph.message`. Same with a `receivedDateTime` window added. A `$count` cannot leave out meeting mail. |
| `/me/mailFolders/delta`, `/me/mailFolders/{id}/messages/delta` (`odata.maxpagesize`) | 200, `deltaLink`, no-change replay returns 0 |
| `/me/translateExchangeIds` | 200. Converts the regular ids `$search` returns (`AQMk…` and `AAMk…`) into immutable ids (live 2026-10-03, read sign-in); the results equal the ids listings return (3 of 3, live 2026-10-04). An input that is already immutable fails the whole call (HTTP 400 `InvalidArgument`, expected `EntryId`). Not needed for OWS (§4.2). |

**Sign-in (live 2026-10-03):** `auth write` asked for its own device code right after `auth read`: One Outlook Web does not reuse Outlook Mobile's sign-in, so two sign-ins are needed. The read token carries 18 Graph scopes (mail: `Mail.Read`, `Mail.Read.Shared`; re-checked 2026-10-04: `Content.Process.User`, `Family.Read`, `FileStorageContainer.Selected`, `Files.ReadWrite.All`, `Mail.Read`, `Mail.Read.Shared`, `People.Read`, `People.Read.All`, `Presence.Read.All`, `ProtectionScopes.Compute.User`, `Sites.ReadWrite.All`, `User.Read`, `User.Read.All`, `User.ReadBasic.All`, `UserAuthenticationMethod.ReadWrite`, `email`, `openid`, `profile`; **no `Calendars.*`**, so `/me/calendarView` was not tried, roadmap X10); the write token 74 Outlook scopes (mail: `Mail.ReadWrite(.All/.Shared)`, `Mail.Send(.Shared)`).

Not tested: shared mailboxes (no target supplied). Online Archive: the account reports `HasArchive=false`, and Graph does not support it.

### 3.2 Mailbox size and catalog cost (`catalog.py`)

| Measure | Result |
|---|---|
| Folders | 23 (1 hidden). Folder delta returns 22; the missing one is a hidden, empty, top-level folder. |
| Items | 25,607: Junk 16,837 · Deleted 5,170 · Sent 1,205 · Archive 1,036 · Inbox 139 · 18 custom ≈ 1,220. `$count` gives 25,466 (likely non-message items). |
| Metadata walk (`$top=500`, 17 fields) | 37 records/s, ~1.7 KB/record → full mirror ≈ 11.5 min, ≈ 45 MB |
| Inbox delta bootstrap | 139 records, 1 page, 1.6 s |

**Conclusion:** cache folders only, and fetch message metadata on demand.

**Mailbox-wide paging (H7, 2026-10-04, read-only).** Listing without a folder means `/me/messages`, newest first, Junk and Deleted Items included (86% of the mailbox); the connector drops them after paging. One page of 100, compared with counting first (one `$count` `$batch` for the window) and then listing only the in-scope folders with mail in parallel, merged newest first:

| Window | `/me/messages`, filter after paging | Count first, then per folder |
|---|---|---|
| Last 7 days | 13 of 100 kept (83 Junk/Deleted dropped), 1.3 s | 87 of 87, 2.7 s, 5 of 16 folders |
| Last 30 days | 13 of 100, 1.3 s | 100 of 100, 4.1 s, 9 folders |
| No window | 13 of 100, 1.0 s | 100 of 100, 10.7 s, 13 folders |
| No window, 25 per page | 2 of 25, 0.4 s | 25 of 25, 2.6 s, 13 folders |

**Conclusion:** counting first fills every page, and per message delivered it is faster (filtering after paging needs about 8 pages for 100 messages), but each page is slower, most of all without a window. The cursor design across folders is still open (roadmap H7).

**Live check of the built per-folder listing (H7, 2026-10-04, read-only).** Service layer, default scope; Junk, Deleted Items and hidden folders hold 86.1% of the mailbox (threshold 66%). Whole-mailbox forced with `PER_FOLDER_SHARE = 2.0`. Graph requests counted at the transport (`$batch` = 1, its sub-requests in brackets):

| Case | Per folder: time, items, complete, requests | Whole mailbox, first page | Whole mailbox, paged to the same items |
|---|---|---|---|
| 25, no window | 3.2 s, 21, no, 15 (+16) | 0.4 s, 5, 1 | 5 pages, 2.5 s |
| 100, no window | 4.7 s, 96, no, 18 (+16) | 1.0 s, 13, 1 | 6 pages, 7.5 s |
| 25, last 7 days | 2.1 s, 21, no, 7 (+16) | 0.4 s, 5, 1 | 5 pages, 2.5 s |
| 100, last 7 days | 2.6 s, 75, yes, 8 (+16) | 0.8 s, 13, 1 | 5 pages, 3.9 s |
| JSONL range export, last 30 days | 11.1 s, 185 messages, 22 (+214) | | 28.3 s, 185 messages, 20 (+198) |

Both methods return the same message ids in the same order in all four cases (0 positions differ). Pages hold 21 of 25 and 96 of 100 because 4 copies (self-sent mail in Inbox and Sent Items) are folded into `also_in`, the same for both methods. The per-folder note names the left-out share and the biggest folders. A first run measured the same within 0.5 s (export 10.8 s vs 36.6 s).

### 3.3 Search (`search.py`)

Three pages of 25 per backend.

| Query | Graph `$search` | Graph `/search/query` | Substrate (§5) |
|---|---|---|---|
| `relatório be` | 75 msgs → 52 conversations, more available | total 238 | 75 conversations, total 340 |
| `subject:relatório` (complete set) | 63 msgs → **50 conversations** | total 65 | **50 conversations**, total 65 |

- All three backends fold accents (`relatório` ≡ `relatorio`).
- After resolving Substrate's `ImmutableId`s through Graph, `subject:relatório` returned the **identical 50 conversations** from both. For `relatório be`, all 52 `$search` conversations appear in Substrate's top 75.
- **Conclusion:** adopt Graph `$search`, and group hits by `conversationId` locally. `/search/query` is not used: its total comes from another engine and ignores the folder rules.

### 3.4 Conversations and reply headers (`conversations.py`)

Sample: 25 recent conversations, 192 messages.

| Measure | Result |
|---|---|
| Conversation filter | Returns every folder: 21 of 25 conversations span 2–4 folders |
| + `$orderby` | **400 `InefficientFilter`** → sort locally |
| Size | median 4, max 40 messages |
| Headers | `Message-ID` 137, `Thread-Index` 136, `In-Reply-To`/`References` 113. **55 messages have no headers at all** (drafts and the user's own messages). |
| Reply tree | 6 of 25 conversations branch (18 branch points). 9 parents are not in the mailbox. |
| `uniqueBody` / `body` | median length ratio 0.10 |

**Conclusion:** `get_conversation` = conversation filter + local sort. Branch detection is feasible for received mail. The user's own messages need a fallback.

### 3.5 Attachments (`attachments.py`)

40 messages, 176 attachments.

| Measure | Result |
|---|---|
| Types | `fileAttachment` 168, `itemAttachment` 8, `referenceAttachment` 0 |
| Inline | 125 of 168 file attachments. All are referenced as `cid:` in the full body; only 47 in `uniqueBody`. |
| `$value` | Works for file attachments, and for item attachments (returns MIME → `.eml`) |
| Sizes | median 11 KB, max 27.5 MB |

**Conclusion:** export non-inline attachments by default. Include inline images only when the rendered body references them.

**Content ids in the listing (2026-10-04, read-only, one message from the last 30 days with an inline image):** `GET /me/messages/{id}/attachments?$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId` → 200; `contentId` came back on the inline file attachment (1 of 1), alongside `@odata.type`, `@odata.mediaContentType` and the selected fields. One listing request per message gives every image's content id, so the separate per-attachment lookup (`attachment_content_ids`) is unnecessary (roadmap H37, H19).

### 3.6 Delta around moves and deletes (`delta_moves.py`)

Snapshot taken before the §4.2 mutations, checked after them.

| Folder | Delta records | Meaning |
|---|---|---|
| Inbox | 2× `@removed`, `reason: "deleted"` | Both were **soft** deletes, still GET-able by id in Deleted Items |
| Inbox | 2× changed | Property edits, and a move out and back (net: a change) |
| Archive | 0 | The round trip netted out. Delta reports net state, not history. |
| Deleted Items | 2× added | The same immutable ids, so moves are matched by id |

## 4. OWS (`/owa/service.svc`): writes (STANDALONE, client A)

### 4.1 Transport contract

- Route: `POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail`.
- Envelope: `{"__type":"<Action>JsonRequest:#Exchange","Header":{"__type":"JsonRequestHeaders:#Exchange","RequestServerVersion":"V2018_01_08"},"Body":{"__type":"<Action>Request:#Exchange",…}}`.
- If the URL-encoded JSON is ≤ 2,048 characters, it goes in the `X-OWA-UrlPostData` header with an empty body. Otherwise it goes in the body.
- Required headers: bearer token, `Action`, `X-OWA-ActionSource`, correlation ids, `X-AnchorMailbox: AAD-SMTP:<user>`, `Prefer: IdType="ImmutableId"`. **No cookies or canary.**
- Response: `Body.ResponseMessages.Items[]` with `ResponseClass` / `ResponseCode`. Some read actions return results directly in `Body`.
- Error headers: `x-owa-error`, `x-owa-returncode`. Result classes: `Success`, `ClientError`, `ServerError`, `ServerTransientError`, `Throttled`.
- Writes are never retried after an ambiguous result.

### 4.2 Proven write contracts

The send is a self-send. The mutations ran on four user-named Inbox messages. Every result was verified through Graph.

| Operation | Request essentials | Result |
|---|---|---|
| Id mapping | Graph immutable id with `-`→`/`, `_`→`+` is accepted as the OWS `ItemId`. The same applies to `conversationId`. | Accepted on all four, and ids survive moves |
| Send | `CreateItem`, `MessageDisposition:"SendAndSaveCopy"`, `ComposeOperation:"newMail"` (full body in `research/probes/common.py` `ows_message`) | `NoError`. One copy in Sent Items, as requested |
| Reply with history (2026-10-04) | `CreateItem` `ReplyToItem` with `NewBodyContent` `BodyType:"HTML"`, `SaveOnly` | The draft quotes the original's HTML (lists, tables, bold kept) and carries its inline images with the same bytes; no `[cid:…]` text. A `Text` body makes Exchange flatten the quoted original instead. |
| Send an existing draft (2026-10-04) | `SendItem` | **Not supported over OWS:** HTTP 500 `OwaOperationNotSupportedException`; nothing sent. |
| Send an existing draft (2026-10-04) | `UpdateItem` on the draft, one `SetItemField` (its subject), `MessageDisposition:"SendAndSaveCopy"`, `SavedItemFolderId` Sent Items | `NoError`. That exact draft is sent; the copy lands in Sent Items and the draft leaves Drafts. |
| Send an existing draft, unchanged and bound to the version read (2026-10-06, self-sends) | `UpdateItem` with **no field updates** (`Updates: []`), the draft's Graph `changeKey` as the `ItemId`'s `ChangeKey`, `ConflictResolution:"NeverOverwrite"`, `SendAndSaveCopy`, `SavedItemFolderId` Sent Items | Current change key: `NoError`, sent, the copy in Sent Items (also with one subject `SetItemField`). Graph's `changeKey` is accepted as the OWS change key. After the draft's subject was changed: `ErrorIrresolvableConflict` with `NeverOverwrite` and with `AutoResolve`, with or without a field update; nothing sent, the draft and its new subject kept. |
| Clear draft recipients (2026-10-06) | `UpdateItem` / `SaveOnly`, `SetItemField` `message:CcRecipients` (then `message:BccRecipients`) with an empty list | `NoError`; Graph reads both back empty. `DeleteItemField` is not needed. |
| Read state | `UpdateItem`, `SetItemField`, `FieldURI:"message:IsRead"`, `ConflictResolution:"AlwaysOverwrite"`, `MessageDisposition:"SaveOnly"`, `SuppressReadReceipts:true` | `NoError` |
| Flag | `UpdateItem`, `FieldURI:"item:Flag"`, `Flag:{__type:"FlagType:#Exchange",FlagStatus:"Flagged"\|"NotFlagged"}` | `NoError` |
| Categories | `UpdateItem`, `FieldURI:"item:Categories"`, `Categories:[…]` | `NoError` |
| Conversation read state | `ApplyConversationAction`, `Action:"SetReadState"`, `ContextFolderId` = a distinguished folder | `NoError` |
| Move | `MoveItem`, `ToFolderId` = `DistinguishedFolderId` (or a folder id), `ReturnNewItemIds:true` | `NoError`. The Graph immutable id is unchanged. |
| Delete (soft) | `DeleteItem`, `DeleteType:"MoveToDeletedItems"` (or `MoveItem` to `deleteditems`) | `NoError`. The item is in Deleted Items under the same id. |

### 4.3 Other OWS knowledge (for the "no Graph" scenario)

- **Read actions proven STANDALONE (2026-10-01), parked while Graph serves reads:** `FindFolder`, `GetFolder`, `FindItem`, `FindConversation`, `GetItem` (`Default` shape caps the body at 2,048 characters), `GetConversationItems`, `GetTimeZone`. `ExecuteSearch` returned 400 with the variants tried. Search would use Substrate instead (§5).
- **BROWSER:** `FindConversation` rows carry cross-folder `GlobalItemIds` / `GlobalMessageCount`, and search results page through a server search folder.
- **Other routes:**
  - `CreateAttachmentFromLocalFile` uploads draft attachments (BROWSER). It is relevant only if send-with-attachments is added.
  - The B2 route `/messageservice/ows/…` returned 500 and is not usable.
  - `/outlookgatewayb2/graphql` is Outlook's internal GraphQL, not Microsoft Graph.
- **Inbox rules (2026-10-04, read attempts only, nothing changed):** the read sign-in (Graph) has no `MailboxSettings.*` scope, so Graph's `messageRules` is out of reach. The write sign-in's Outlook token carries `MailboxSettings.ReadWrite`. Over OWS, `GetInboxRules` (the EWS name) returned `OwaOperationNotSupportedException`; `GetInboxRule` (the name Outlook Web's settings use) exists but returned `NullReferenceException` with an empty request body, because the rule actions use a different envelope (captured 2026-10-04, §4.4).

### 4.4 Inbox rules (captured 2026-10-04 in Outlook on the web; owner's throwaway rule only)

Two captures (2026-10-04). In the first, the owner captured Outlook on the web's rules page (`Settings → Mail → Rules`) while creating, editing, disabling and deleting a throwaway rule ("ZZ connector test rule": From `nobody@example.invalid`, Subject includes `zz-connector-test`, Move to a folder, Mark as read, Stop processing more rules), then reordering two rules and back. In the second, two more throwaway rules: Sent to, Subject includes, Move to Archive; disabled, re-enabled, conditions cleared, moved to a user folder, renamed, deleted. Shapes below are from that capture with every value redacted; the capture itself (it held tokens) was deleted after analysis. **Replayed by the connector on 2026-10-04 (STANDALONE, write token): every action works as captured** (see "Replay" at the end of this section).

**A different envelope from item actions (§4.1–4.2).** The rule actions send the request object itself, with no `…JsonRequest` wrapper and no `Body`:

```json
{"__type": "GetInboxRuleRequest:#Exchange",
 "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08",
            "TimeZoneContext": {"__type": "TimeZoneContext:#Exchange",
                                "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange", "Id": "<Windows time zone>"}}},
 "UseServerRulesLoader": true}
```

This explains the 2026-10-04 `NullReferenceException`: the connector's `Ows.call` wraps every body in `{"__type": "<Action>JsonRequest", "Header", "Body"}`. The answer is also different: no `ResponseMessages`/`Items`, but `WasSuccessful` (bool), `ErrorCode` (0 on success), `ErrorMessage`, `UserPrompt`, `IsUserError`. Same transport otherwise: `?action=<Action>&app=Mail`, `Action` header, the payload in `X-OWA-UrlPostData` when short (the reorder call, being long, went in the body), `X-AnchorMailbox`, bearer token.

| Action | Request (besides `__type` and `Header`) | Answer |
|---|---|---|
| `GetInboxRule` | `UseServerRulesLoader: true` | `InboxRuleCollection.InboxRules`: every rule, in priority order (`Priority` 1 = first) |
| `NewInboxRule` | `InboxRule`: `Name` and only the conditions and actions used, e.g. `From: [{"__type": "PeopleIdentity:#Exchange", "DisplayName", "SmtpAddress", "RoutingType": "SMTP"}]`, `SubjectContainsWords: [..]`, `MoveToFolder: {"DisplayName", "RawIdentity": <folder id>}`, `MarkAsRead: true`, `StopProcessingRules: true` | `InboxRule`: the created rule (all fields, `Identity`, `Priority` 1: a new rule goes first) |
| `SetInboxRule` | `InboxRule` with `Identity: {"DisplayName": <rule name>, "RawIdentity": <rule id>}` and the fields being set; `Force: false` | `WasSuccessful` … |
| `DisableInboxRule` / `RemoveInboxRule` | `Identity: {"DisplayName": <rule id>, "RawIdentity": <rule id>}` | `WasSuccessful` … |
| `SetInboxAndSweepRules` | `EnableDisableInboxRules`: one `{"__type": "EnableDisableInboxRuleRequest:#Exchange", "Header", "Identity", "IsEnabled"}` per rule, **every rule, in the new order** | `WasSuccessful` … |

Observed semantics:

- **Rule id:** `<mailbox GUID>\<20-digit number>`, stable across edits, including a rename (`SetInboxRule` with a new `Name`, `Identity.DisplayName` set to the new name, the same `RawIdentity`).
- **Edit is partial:** the `SetInboxRule` request carried `Name`, `From`, `SubjectContainsWords`, `StopProcessingRules` and `Identity`, but not `MoveToFolder` or `MarkAsRead`; the rule read back afterwards still had both. Fields left out are kept. **A field sent as `null` is cleared** (second capture: `SubjectContainsWords: null` and later `SubjectOrBodyContainsWords: null` each removed that condition, read back as null). Outlook always resends `Name`, the sender/recipient condition, `StopProcessingRules` and `Identity`, plus the fields that changed.
- **Order:** reordering sends `SetInboxAndSweepRules` with all rules in the new order; there is no `Priority` field in the request. It also carries each rule's `IsEnabled`, so the same call can enable or disable rules.
- **Enable:** `EnableInboxRule`, the same shape as `DisableInboxRule` (`Identity` only); read back `Enabled: true` (second capture). Outlook's toggle sometimes sends `DisableInboxRule` twice in a row (both captures); harmless, the rule stays disabled.
- **Folders:** in requests, `MoveToFolder.RawIdentity` is a base64 folder id starting `AQMk` (116–120 characters; Archive and a user folder), the same family as the regular ids `$search` returns (H12); **the Graph folder id that `list_folders` returns is accepted as is** (replay 2026-10-04; in this mailbox all 18 folder ids are `AQMk…` with no `-` or `_`, so their OWS form (§4.2) is the same string and was not a separate case). In answers, `MoveToFolder.RawIdentity` is a different form: `<organisation path>/<mailbox GUID>:\<folder name>` (136–140 characters, the folder's own name only, e.g. `:\Archive`), so reading a rule's target folder means matching by name (ambiguous when two folders share a name) or by `DisplayName`.
- **Other conditions and flags (second capture):** `SentTo` has the same `PeopleIdentity` shape as `From` (read back with `Address` and `AddressOrigin` added). `NewInboxRule` from the UI also carried `DisplayAlert: "Default"` or `PlaySound: "Default"` depending on the options shown; neither is needed.
- **A rule has 90 fields** (conditions, `ExceptIf…` exceptions, actions, `Description` texts, `InError`, `SupportedByTask`, `RuleProvider`). The owner's 8 rules use only `MoveToFolder` (8), `SentTo` (6), `SubjectContainsWords` (5), `From` (1), `SubjectOrBodyContainsWords` (1), all with `StopProcessingRules`, all enabled, none in error.
- Outlook also calls `GetMailboxByIdentity` after each change and `/ows/v1.0/OutlookOptions/MailForwardingNotification` on the rules page; neither is needed to manage rules.

**Replay (2026-10-04, STANDALONE, `Ows.call_request`; one throwaway rule, no other rule touched).** Same URL and headers as the item actions, bearer write token, no cookies.

- `GetInboxRule` with `UseServerRulesLoader: true`: `WasSuccessful: true`, `ErrorCode: 0`; answer keys `ErrorCode`, `ErrorMessage`, `Header`, `InboxRuleCollection`, `IsUserError`, `UserPrompt`, `WasSuccessful`. 8 rules, 90 fields each (`Identity` is `{DisplayName, RawIdentity}`; `Enabled`, `Priority`). **`TimeZoneContext` is not needed:** without it, the same 8 rules in the same order.
- `NewInboxRule` (`Name`, `SentTo` one `PeopleIdentity`, `SubjectContainsWords`, `MoveToFolder: {DisplayName, RawIdentity: <Graph folder id of Archive>}`, `StopProcessingRules: true`; no `__type` on `InboxRule`): success in 2.3 s; the answer also carries `InboxRule`. Read back: 9 rules, the new one `Priority` 1, enabled, its target folder returned as `…:\<folder name>`.
- `SetInboxRule` (`Identity`, `Name`, `SentTo`, `StopProcessingRules`, `SubjectContainsWords: null`, `SubjectOrBodyContainsWords: [..]`, `Force: false`): the subject condition cleared, the subject-or-body one set; `MoveToFolder`, `StopProcessingRules` and `SentTo` unchanged.
- `SetInboxRule` renaming it (`Identity.DisplayName` = new name, same `RawIdentity`): renamed, **id unchanged**, other fields kept.
- `DisableInboxRule`, then `EnableInboxRule` (`Identity` only): read back `Enabled` false, then true.
- `SetInboxAndSweepRules` with all 9 rules (each `{__type: EnableDisableInboxRuleRequest, Header, Identity, IsEnabled}`, payload in the body), the test rule last, then first: 0.7–0.8 s each; read back in exactly that order, `Priority` renumbered 1…9, every other rule's relative order and `Enabled` kept.
- `RemoveInboxRule`: gone. The final list equals the starting one: 8 rules, same ids, order, priorities and enabled states.
- Every write was sent once and answered `WasSuccessful: true`; no unclear answer, nothing retried.

### 4.5 Connector live validation (2026-10-07)

Owner-authorized dedicated test messages, self-sends and two personal recipient addresses. The
production MCP tools and connector service/transport were used with the app's encrypted token cache;
no browser credentials or plaintext probe-token cache. Live harness output recorded counts, statuses, field names
and verification booleans only. Private recovery ids were kept in an encrypted, ignored local file.

**W7 — HTML and replies.**

- HTML fragment draft creation and body editing passed mandatory Graph read-back without findings.
  Tables, lists, links, line breaks, accented and Japanese text were retained in the saved draft,
  Sent Items and received self-copy. Existing-draft sends returned `sent`, without reconstruction.
- A dedicated inline-image fixture was saved through `Ows.call("CreateItem", ...)`, using the
  production new-message mapping plus an `Attachments` array containing a synthetic PNG
  `FileAttachment:#Exchange` (`Content`, `ContentType`, `ContentId`, `IsInline`). Graph confirmed one
  inline attachment and the expected body CID. This was fixture setup, not a newly implemented
  attachment-upload tool. Recipient-only editing through `edit_draft` retained it without findings.
- Text reply and HTML reply-all drafts through `create_draft` both returned `verified: true` and
  `history_intact: true`. The connector's existing checks confirmed quoted structure and inline-image
  bytes. Default reply recipients selected the original's personal To recipient; reply-all also kept
  its personal Cc recipient and omitted the sending account. The work account was explicitly added
  back for self-copy inspection before both drafts were sent.
- All four sent copies and all four received self-copies retained tables and lists. The inline
  original and both replies in each folder retained matching CIDs and PNG bytes identical to the
  fixture. The text reply retained a literal angle-bracket expression; the HTML reply retained its
  own ordered list in addition to the quoted history.
- The owner subsequently confirmed on 2026-10-07 that the received test emails looked correct and
  HTML rendered properly. Individual recipient/client combinations were not specified. Automated
  recipient-side inspection was unavailable: the Gmail connector was unauthenticated and browser
  automation failed to initialize. Fragment/full-document comparisons, malformed accepted HTML,
  CSS and remote images remain deferred research. No claim of pixel-for-pixel rendering equivalence
  is made.

**W8 — signature discovery.** Minimal text, repeated minimal HTML and empty HTML drafts returned
only the submitted text or an empty body, with no extra signature text or images. The empty draft
was saved and correctly flagged as empty. Automatic signature extraction through server draft
creation was not demonstrated. The subsequent owner-supplied Web capture establishes that signature
HTML and images were supplied by the compose client (§4.6). It also exposes the corporate template
source and add-in insertion code. Native signature list/default/content retrieval was subsequently
proved with app-owned authentication (§4.7). Automatic draft integration remains unimplemented.

**W9 — real rule shapes expose two implementation blockers.**

- `list_rules` marked all eight baseline rules read-only. An authorized throwaway rule with an
  invalid sender, unique subject marker, Archive destination and stop-processing action was created
  once and appeared in fresh read-back with the intended conditions and destination.
- OWS supplies description metadata (`DescriptionTimeFormat`, `DescriptionTimeZone`) and inactive
  enum strings: `NullInboxRuleMessageFlag` for `FlaggedForAction`, `RequestedAction` and their
  `ExceptIf` counterparts; `NullInboxRuleMessageType` for `MessageTypeMatches` and its exception;
  `NullImportance` for `MarkImportance`, `WithImportance` and its exception; `NullSensitivity` for
  `WithSensitivity` and its exception. The adapter currently treats all 13 fields as unsupported
  behavior. The throwaway rule therefore became read-only too, preventing its tool update lifecycle.
- The `NewInboxRule` answer's `RawIdentity` differed from the same rule's `GetInboxRule` identity.
  `create_rule` retained the creation identity, could not find it in read-back, and returned `failed`
  even though the rule existed. The write was not retried. Creation needs reconciliation against
  fresh list identity when the reported identity does not match, not only when no id is returned.
- Replaying the now-stale creation confirmation was rejected without another write. Reordering the
  unsupported collection was refused. Complete public-tool edit/clear, disable/enable, delete and
  successful reorder validation remains pending the fixes.
- Cleanup removed only the newly created rule through `OwsRules.delete_rule`. Fresh read-back
  confirmed the original eight rules' ids, relative order, priorities, enabled states and full
  revisions exactly matched the starting snapshot. Existing rules were never rewritten.

**H21 — conversation expansion beyond 100 messages.** 101 unsent reply drafts plus two original
copies formed a 103-message conversation. `set_read_state` through the MCP surface returned
`done: 103` for unread and then read, with zero ordinary detailed results and one compact-results
note each time. Independent Graph summaries confirmed all 103 states, and initial states were
restored. No truncation occurred; the 1,000-message boundary remains synthetic coverage.

**H17 — controlled partial failures with real writes.** On 60 dedicated drafts (three chunks), a
locally injected pre-send error in chunk two yielded `done: 40, failed: 20` with default continuation.
With `continue_on_error=false`, it yielded `done: 20, failed: 40`, including 20 `not sent`; the third
chunk was never attempted. Fresh Graph states matched both results. A completed real 20-item write
followed by injected response loss and failed read-back yielded `unknown: 20`; an independent read
confirmed the changes. All initial states were restored. These are injected failures, not evidence
of natural Microsoft outages; flag/move/delete failure variants remain synthetic coverage.

### 4.6 Outlook Web signature capture (2026-10-07)

**Scope and confidence.** Offline review of the owner-supplied HAR: 105 entries, 96 captured response
bodies, 37 requests with body text and eight with `X-OWA-UrlPostData`. Request start times span
09:05:00–09:06:47 America/Sao_Paulo (106.407 seconds). No captured credential was replayed and no
mailbox changes were made during this analysis. Only schema, aggregate counts and conclusions are
recorded here. Entry numbers below are zero-based positions in the supplied capture.

The owner reports also adding, editing and deleting a signature. Those actions are owner-reported;
this HAR does not expose an identifiable native signature CRUD sequence. It contains compose,
add-in settings, template retrieval, inline uploads and draft save traffic. This limits what can be
claimed about native settings endpoints, without contradicting the owner's actions.

**Main finding — client-supplied signature content.** The initial `CreateItem` request already
contains HTML with `id="Signature"`; OWS does not generate that block in its response. Later,
`UpdateItem` submits a signature block with four `cid:` images after four successful inline uploads.
This directly establishes client-supplied content for the captured draft. The loaded officeatwork
Mail Signature add-in source contains the rendering and Office.js insertion path described below.
Attributing every initial text-only signature byte to that add-in, or establishing native signature
storage, would require additional evidence.

**Useful network contracts (BROWSER, requests and responses):**

| Entry | Operation | Observed result and use |
|---|---|---|
| 8 | `GET /ows/v1/OutlookCloudSettings/settings/` | HTTP 200, empty array; no signature content in this response. |
| 29 | OWS `CreateItem`, `SaveOnly` | Submitted HTML already contains a text-only signature block; saved successfully. |
| 32, 87, 102 | OWS `GetItem` | Draft read-back before/after inline uploads and body save; final item is still a draft with four inline attachments. |
| 33 | Outlook beta `translateExchangeIds` | HTTP 200; the compose/add-in flow translates an item identifier. |
| 50, 64 | `LoadExtensionCustomProperties` | Successful, empty custom-property objects. |
| 56, 59 | `PATCH /api/beta/users/{user}/mailFolders/Inbox/UserConfigurations/{configuration}` | HTTP 200; `DictionaryData` is serialized XML, not signature HTML. See settings distinction below. |
| 58 | OWS `SanitizeHtml` | Input HTML 2,471 characters; response is a JSON string containing 1,755 characters of HTML. No standalone sanitization contract tested. |
| 61, 66, 70, 73 | `POST /owa/service.svc/CreateAttachmentFromLocalFile` | Four HTTP 200 / `NoError` uploads; inline attachment IDs and content IDs returned. |
| 67 | Vendor `GET /api/appSettings/mailSignature` | HTTP 200; corporate signature configuration, library references, licensing and language settings. |
| 74 | Graph `GET /v1.0/me` | HTTP 200; profile properties available for template substitution. |
| 76 | Graph `POST /v1.0/me/getMemberGroups` | HTTP 200; request uses `securityEnabledOnly: false`. Group-based selection is present in the client source; this request alone does not prove its final choice. |
| 82 | Graph `GET /v1.0/drives/{drive}/list/items` | `$expand=fields,driveItem`, `$filter=fields/DocIcon eq 'ofawmsig'`; one template package returned. |
| 88 | SharePoint package download | HTTP 200; ZIP-format `.ofawmsig`, 16,815 bytes. |
| 89 | OWS `SaveExtensionSettings` | Successful; selected template reference persisted in add-in roaming settings. |
| 90, 91 | Graph `GET /v1.0/shares/{encoded-url}/driveItem?$expand=listItem` | Two image file lookups. The capture uses a doubled slash after `v1.0`; that is observed spelling, not a required contract. |
| 96, 97 | SharePoint image downloads | Two complete PNG response bodies, 10,836 and 206,953 bytes. |
| 101 | OWS `UpdateItem`, `SaveOnly` | HTML containing two tables and four CID images saved successfully. No send operation captured. |

**Corporate template pipeline.** The vendor settings response contains three SharePoint library
definitions and 15 content-language entries. The downloaded package contains `template.njk`
(7,559 bytes), `metadata.json` (4,385 bytes), `images.json` (671 bytes), and five embedded PNGs.
This is a Nunjucks template with configuration and assets, rather than a ready-to-send HTML file.
The client source constructs a context from sender/user profile, language, audience, compose type,
item type, coercion type and configurable form fields before rendering it.

The template references `user.givenName`, `surname`, `jobTitle`, `companyName`, `streetAddress`,
`postalCode`, `city`, `country`, `businessPhones[0]`, `mobilePhone` and `mail`. It uses
`field(...).value`, `image(...).cid`, `mail.isPlainText`, filters such as `removeEmpty`, `join`,
`lower` and `trim`, and the custom `loadImageFileFromSharePoint(...)` helper for additional images.
There are country-dependent branches and optional phone/image fields. URLs, field values and
personal data are deliberately omitted from this record.

The captured metadata has ten form elements: language and audience pickers, switches, single-line
text fields and a multi-line text field. Its automatic-insertion switches are:

| Event | Configured |
|---|---|
| New message, forward, reply | `true` |
| Recipients changed, sender changed | `true` |
| New appointment, message send | `false` |

These are template policy values, not independently exercised event tests. They explain why a
static signature snapshot cannot reproduce all vendor selection/rendering behavior.

**Insertion calls (SOURCE).** The loaded vendor code wraps
`Office.context.mailbox.item.addFileAttachmentFromBase64Async(..., {isInline: true}, ...)` and
`Office.context.mailbox.item.body.setSignatureAsync(html, {coercionType: ...}, ...)`; the insertion
path processes HTML image assets before invoking the signature setter. Microsoft's
[`Body.setSignatureAsync` documentation](https://learn.microsoft.com/en-us/javascript/api/outlook/office.body?view=outlook-js-preview#outlook-office-body-setsignatureasync-member(1))
describes this as a compose API that adds or replaces a signature in the item body. It is not a
mailbox signature-settings read API. The capture shows the resulting upload/save contracts, but
does not include an Office.js callback/event trace proving each source function executed.

**Settings distinction.** `SaveExtensionSettings.request.Settings` is serialized JSON. Its
`ofaw.user.signatureFileReference` value requires two further JSON decodes to obtain
`{sourceId, itemIdentifier:{type, name, driveId, driveItemId, webUrl}}`. This is a selected template
pointer, not the rendered signature HTML. It may offer a discovery route, but reading it with the
connector's own token has not been tested.

The two `UserConfigurations` PATCH requests carry `<UserConfiguration><Info .../><Data><e .../>`
XML in `DictionaryData`. The `18-ExtensionSettings` entry decodes to Sales add-in preference keys:
`SPTenantId`, `SPUserId`, `isEURegionKey`, `region`, `SPEnvironmentType`, `SPDefaultLocation`.
Responses also contain `18-OLPrefsVersion`. Neither write stores signature HTML; they must not be
counted as native signature creation/edit/deletion. Likewise, `SaveExtensionCustomProperties`
stores `SalesProductivityExternalContacts`, unrelated to the signature template reference.

**Inline upload contract.** The request metadata resides in URL-encoded JSON in
`X-OWA-UrlPostData`; the request body holds image bytes. The envelope types are
`CreateAttachmentJsonRequest:#Exchange` → `CreateAttachmentRequest:#Exchange`. Relevant fields are
`ParentItemId:{Id,ChangeKey}`, `Attachments:[{__type:"FileAttachment:#Exchange", Content:"",
ContentType:"image/png", IsInline:true, Name, Size, ...}]`, `IncludeContentIdInResponse:true`,
`SliceNumber:0`, `TotalSlices:1`, and cancellation/IRM flags. Each response attachment includes
`AttachmentId`, `ContentId`, `LastModifiedTime`, `IsInline`; the returned content IDs match the
four references in the subsequent saved HTML. No independent app-owned upload replay was done.

**Extraction and read-back trap.** The HAR's upload bodies were decoded lossily as UTF-8, so they
cannot supply exact original PNG bytes on their own. Complete binary bodies were recoverable from
the package and SharePoint downloads. Each of the four uploaded assets matched a unique recovered
PNG by size (10,836, 1,751, 1,000 and 997 bytes), and decoding that PNG as UTF-8 with replacement
exactly reproduced the HAR upload text. This is strong correlation, not a hash comparison against
the unavailable original upload bytes.

The final `GetItem.NormalizedBody` instead contains four `data:image/gif;base64,...` sources, each
decoding to a 42-byte placeholder. Those are not the signature PNGs. Preserve submitted CID HTML
and retrieve actual attachments; do not treat this display-oriented normalized body as an asset
export. The private local extraction contains only the saved signature block (1,741 characters),
four recovered PNGs, the original template and reduced metadata; CID references were replaced with
local image paths. It is ignored by Git. No credentials, profile JSON or full message was exported.

**Authentication limits.** A successful vendor-resource token response reports scope names
`Files.Read.All`, `Sites.Read.All`, `User.Read`, but its audience is the vendor resource, not Graph.
This does not demonstrate those scopes for the connector's clients or prove that the token was
used for the captured Graph requests. A separate Graph `Files.ReadWrite.AppFolder` request fails
with `invalid_grant` / `AADSTS65001`; the other template and image reads still succeeded in this
browser session. These facts are not grounds to copy vendor tokens or add broad permissions.

**Follow-up:** §4.7 supersedes the native retrieval gap: app-owned list/default/content reads now
work. App-owned Graph reads also retrieved the first capture's Web-created draft signature and four
inline PNGs (10,836, 1,751, 1,000 and 997 bytes), with every body CID backed by an attachment.
Corporate template-settings access and policy execution remain untested. W9's public-rule blockers
are unchanged: this HAR contains no inbox-rule operations. Calendar authentication is unaddressed.

### 4.7 Native roaming-signature discovery and standalone reads (2026-10-07)

**Discovery (SOURCE).** The second owner-supplied HAR contains 343 entries, including native Outlook
settings and compose JavaScript. `owa.93987.m.5c7464db.js` defines `signaturehtml`, `signaturetxt`,
`roaming_signature_list`, `roaming_new_signature` and `roaming_reply_signature`. Its save code uses
`RoamingSetting` entries scoped to the account, HTML/text secondary keys `htm`/`txt`, and
`parentSetting: "roaming_signature_list"`. The list is written as comma-joined names (`BlobArray`),
while the two defaults are separate `String` values. Worker source uses the `x-islargesetting`
header for large signature settings. These are code observations; the HAR does not contain a full
native create/edit/delete request lifecycle.

Compose source in `owa.MailComposeActions.m.2ceb7db6.js` selects `defaultSignatureName` for new mail
and `defaultReplySignatureName` for replies/forwards, then looks up the selected entry in the roaming
signature map. The older `UserOptions.SignatureHtml`/`SignatureText` path is separate. A standalone
`GetOwaUserConfiguration` call succeeds with our OWS sign-in but returns null legacy signature
contents; that does not mean the modern roaming-signature list is empty. One attempted
`GetUserConfiguration` JSON envelope returned HTTP 400; no general lack of configuration-read
access is inferred from that malformed/unsupported variant. Microsoft's documented Graph
[`mailboxSettings` properties](https://learn.microsoft.com/en-us/graph/api/resources/mailboxsettings?view=graph-rest-1.0)
do not expose this native signature list/default/content model.

**Graph-first decision (2026-10-07).** The owner prefers Graph because its contracts are documented.
Microsoft's current [`PostponeRoamingSignaturesUntilLater` documentation](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-organizationconfig?view=exchange-ps#-postponeroamingsignaturesuntillater)
explicitly says it has no plans to support roaming-signature management in Graph and recommends
the Office.js signature API/event hooks for vendors. The documented Graph beta
[`userSettings`](https://learn.microsoft.com/en-us/graph/api/resources/usersettings?view=graph-rest-beta)
model also does not expose native signature list/default/content properties. These are documentation
findings, not proof that every possible undocumented Graph route fails. Use OWA only for this native
configuration gap; keep Graph for message/draft and attachment reads. Revisit if Graph gains a
documented equivalent rather than treating OWA as the preferred general backend.

**Proven reads (STANDALONE).** All requests below used the connector's existing encrypted, app-owned
`write` profile (One Outlook Web → Outlook resource), no browser credentials. Reads were sequential
and did not change mail or signature settings. No extra scopes, vendor authentication or user file
upload was required. Responses and private snapshots remain outside Git.

| Purpose | Proven request | Result |
|---|---|---|
| List and default pointers | `GET /ows/v1/OutlookCloudSettings/settings/?settingname=roaming_signature_list,roaming_new_signature,roaming_reply_signature` | HTTP 200, three settings. List `type: BlobArray`; default settings `type: String`. Every returned `scope` matched the bound account. |
| Signature contents | Same endpoint, URL-encoded `settingname={name}`, header `x-islargesetting: true` | HTTP 200, per-format entries identified by `secondaryKey`. Default returned `htm`, `rtf`, `txt`; a newly created image signature returned `htm`, `txt`. |
| Freshness verification | Repeat the list/default GET after retrieving contents | Settings values and timestamps unchanged in the tested read sequence. |

Headers used were `x-outlook-client: owa` and `Accept: application/json`, plus the large-setting flag
for content reads. No claim is made that all are mandatory. Native entries include `name`, `value`,
`type`, `source`, `scope`, `metadata`, `parentSetting`, `Timestamp`, `itemClass`, `id`, timestamps and
`secondaryKey`; large responses also contain `nameBase64Encoded`. For this probe the returned names
matched the list entries directly. Do not infer an extra decoding step from that flag without
testing its semantics. Request names must be URL-encoded rather than interpolated into query text.

**Current-default retrieval.** The first read returned two list names. Both the new-message and
reply/forward pointers selected the same readable entry, with 1,629 characters of HTML, 152 of text
and an RTF value. That HTML contains no image tags. A fresh list/default read after content retrieval
returned exactly the same values and `Timestamp` fields. A review-only HTML snapshot was saved
privately; application composition must perform fresh reads, not use that file as a cache.

**Empty list reference.** The other initial list name returned HTTP 200 with `[]`, both on an
individual content query and a combined name query. The owner confirmed the native settings UI
showed one signature. This supports a leftover list reference, but does not establish all deletion
or tombstone semantics. A reliable list operation must resolve contents and report missing entries;
a raw name is not proof that a retrievable signature exists. The selected default was readable.

**New signature with image, observed live.** The owner then created a new signature and added an
image in Outlook. A fresh list read contained the existing and new signatures and no longer contained
the previous empty reference. Both entries were readable. Both default pointers still selected the
existing signature: adding a signature did not automatically change the defaults in this test.
The new entry returned 82,274 characters of HTML, 32 of text and one `<img>` with an embedded
`data:image/png;base64,...` source. Decoding it produced a 61,483-byte PNG with a valid PNG header.
Its complete HTML and image were recovered privately through our authentication, without a new HAR.
No native signature send/recipient rendering or automatic application insertion was tested here.

**Unsuccessful variants.** A content-name URL path (`/settings/{name}`) returned HTTP 404; use the
proven query route. A guessed `parentsetting=roaming_signature_list` query returned HTTP 403 and
does not prove inability to read children: exact-name large-setting reads work. No guessed write
request was sent during this read-only discovery phase, and no unsuccessful authentication/client-scope
pair was retried. The subsequent authorized lifecycle writes are recorded in §4.8.

**Owner-selected W8 design.** Read the current native configuration automatically, expose reliable
listing/content retrieval, and resolve the correct default per compose operation. Check account
scope and completeness; respect an explicitly empty default. If a configured default is missing,
the read fails, or its revision changes during retrieval, report it rather than silently using an
old snapshot or another signature. HTML image data must become actual inline assets/CIDs during
composition. Signature selection is done before saving the draft; send still sends that saved draft
unchanged. API discovery is proved; public read tools and automatic integration are not implemented.

**Separate corporate add-in.** These are the native Outlook defaults. The officeatwork add-in can
subsequently replace/render a signature using sender, recipient, language and policy context (§4.6).
Reading native defaults does not reproduce that additional logic. Graph draft extraction is now
proved, but is a snapshot rather than an always-current corporate-default resolver.

### 4.8 Native signature CRUD and default lifecycle (2026-10-07)

**Scope (STANDALONE).** The owner explicitly requested the remaining live tests. All operations used
the connector's existing encrypted, bound-account `write` profile against Outlook Cloud Settings,
with no browser credentials or new scopes. Only uniquely named synthetic signatures and temporary
default selections were changed. No real signature contents were overwritten; no messages were
created, edited or sent. Writes were sent once, with fresh reads to verify their results and cleanup.
Recovery state was encrypted and never committed.

**Proven write contracts:**

| Action | Request | Result |
|---|---|---|
| Create and update HTML/text | `PATCH /ows/v1/OutlookCloudSettings/settings/account`, JSON array of per-format settings; `x-islargesetting: true` | HTTP 200. New contents readable immediately; created signature automatically appeared in the name list. |
| Set/clear a default | Same PATCH endpoint, `String` settings; `x-islargesetting: false` | HTTP 200. New-message and reply defaults changed independently; empty new-message default accepted. |
| Delete all fixture formats | `DELETE /ows/v1/OutlookCloudSettings/settings/account`, JSON body `{name}`; `Content-Type: application/json`, `x-islargesetting: true` | HTTP 200. Content read returned `[]`; fixture was also removed from the name list. |

Content patches contain `itemClass: "RoamingSetting"`, `name`, account `scope`, `secondaryKey: "htm"`
or `"txt"`, `type: "Blob"`, `value`, `parentSetting: "roaming_signature_list"`,
`metadata: "encoding:utf-8"`, a current UTC .NET-ticks `timestamp`, and `value@is.Large: true`.
Default patches contain the same item class/scope, `name` and `secondaryKey` equal to
`roaming_new_signature` or `roaming_reply_signature`, `type: "String"`, and `value`. Headers also
included `x-outlook-client: owa`, `Accept: application/json` and `x-overridetimestamp: false`.
This proves the tested header/payload combination, not that every field is independently mandatory.
The captured native source and this [original HTTP implementation](https://gist.github.com/EionRobb/21dd89b05417b247bface6373fb380fb)
provided wire-format leads; our own live calls establish capability for this account.

**Content lifecycle.** A dedicated fixture with a space-containing name was created with synthetic
HTML and text containing accents and Japanese characters. HTML read-back was normalized: Outlook
added wrappers/CRLFs and changed `id="Signature"` to `id="x_Signature"`. Visible text and the text
format matched exactly. An edit changed the text and added a base64 PNG; read-back retained the
expected visible/text content and exactly matching decoded PNG bytes. A later edit removed the
image and retained the updated text. Deletion removed both formats and its name-list reference.
No append-at-end order was assumed; the original signatures' relative order was preserved.

**Default lifecycle.** With the fixture's image-bearing HTML selected as the new-message default,
fresh default resolution retrieved that current HTML. The reply default initially stayed unchanged;
it was then separately set to the fixture. Clearing only the new-message default produced an
explicit empty string while the reply default still selected the fixture. Both original values
were restored before ordinary fixture deletion. These calls establish server configuration behavior,
not automatic signature insertion by the connector, which is unimplemented.

**Deleting the selected default.** A separate temporary fixture was selected only for new messages
and deleted. Contents returned `[]` and the name was absent from the list, but
`roaming_new_signature` still pointed to the deleted name. The reply default stayed unchanged.
Cleanup restored the original new-message default. A consumer must verify that selected content
exists and report a missing configured default, rather than silently falling back to another name
or a cached signature. The backend DELETE alone does not maintain the default pointer.

**Probe corrections and cleanup.** Two initial setup passes stopped on overly strict harness
expectations (byte-identical HTML and insertion at the end of the list); both fixtures were cleaned
up. A DELETE with no JSON content type returned HTTP 415, and a query-only name with that header
returned HTTP 400. The corrected JSON-body DELETE succeeded. Rejected requests were not automatically
retried. Four temporary signatures in total were created and removed, including the complete
lifecycle and selected-default deletion fixture. Final reads confirmed the original two signatures'
full records were unchanged (hashed comparison), all test contents were absent, and original
signature-list/default values were restored. Default/list revision timestamps can advance when
values are changed and restored; no claim of restoring those metadata revisions is made.

**Limits.** The tests prove native create/read/update/delete, embedded PNG add/remove, list
registration/removal, independent default changes and no-default behavior through our authentication.
They do not prove rename, editing RTF, concurrent-write races, every image format, public MCP write
tools, or integration/recipient rendering of automatic signatures in a draft. Graph-first routing
remains unchanged: these native operations fill the documented roaming-signature gap in Graph.

## 5. Substrate search (`/searchservice/api/v2/query`), parked

- **Works STANDALONE** at `https://outlook.office.com/searchservice/api/v2/query` with the `search` token (§2) and `X-AnchorMailbox: Oid:<oid>@<tid>`.
- Minimal request: `{Cvid, LogicalId, Scenario:{Name:"owa.react"}, TimeZone:"UTC", TextDecorations:"Off", EntityRequests:[{EntityType:"Conversation", ContentSources:["Exchange"], Query:{QueryString, DisplayQueryString}, From, Size, Sort:[], PropertySet:"Optimized"}]}`. **Reuse the same `Cvid` / `LogicalId` across pages.** A new session per page returns 400.
- Response: `EntitySets[].ResultSets[]` with `Total`, `TotalWithoutCollapsing`, `MoreResultsAvailable`, and `Results[].Source` holding `ConversationId`, `ImmutableId`, `ParentFolderId`, `From`, `UniqueSenders`, `MessageCount`, `Preview`, `HasAttachments`, ….
- The raw `ConversationId` uses a different encoding from Graph's, but `ImmutableId` is directly GET-able in Graph.
- Its only advantage over Graph `$search` is server-side conversation collapsing.

## 6. Rejected alternatives

- **Browser-token tools** (`outlook-cli`, `owa-piggy`) offer useful route vocabulary, but their authentication model is unacceptable. [outlook-cli](https://github.com/yusufaltunbicak/outlook-cli) · [owa-piggy](https://github.com/damsleth/owa-piggy)
- **Outlook REST v2** is decommissioned as a public API. [Comparison](https://learn.microsoft.com/en-us/outlook/rest/compare-graph)
- **`GetAccessTokenForResource` via OWS** is a historical token-minting path. Do not use it. [Black Hat 2019](https://i.blackhat.com/USA-19/Thursday/us-19-Jaiswal-Preventing-Authentication-Bypass-A-Tale-Of-Two-Researchers.pdf)
- **EWS:** Exchange Online disablement runs October 2026 → April 2027. [EWS retirement](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online)

## 7. Identifiers worth knowing

- The Exchange resource app id is `00000002-0000-0ff1-ce00-000000000000`. It is not a client id.
- `ab0455a0-8d03-46b9-b18b-df2f57b9e44c` appears in Outlook Web's consumer-auth config. It is unverified and must not be used.
- Deep links (BROWSER): `/mail/inbox/id/<id>`, `/owa/?viewmodel=ReadMessageItem&ItemID=<id>`. These could be "open in Outlook" links in the UI.

## 8. Mailbox side effects of research

- One synthetic self-send to `lucas.rhode@landisgyr.com` (2026-10-01).
- 2026-10-02, on four user-named Inbox messages: read and flag toggles, a category set and then cleared, a conversation read toggle, a move round trip, and two soft deletes. Two messages remain in Deleted Items. Sent copies were untouched, and nothing was hard-deleted.
- 2026-10-04 (W9 replay): one throwaway inbox rule (conditions that match no real mail) created, edited, renamed, disabled, enabled, moved last and first, and removed. The rule list ended as it started.
- 2026-10-07 (connector validation): four dedicated messages were sent, each with a received
  self-copy; personal recipients were included as authorized. The sent and received copies remain
  as evidence. 101 unsent bulk reply drafts and four signature probes were moved to Deleted Items
  (`done: 105` in total), all locations independently verified; no test drafts remain in Drafts.
  Bulk and partial-failure read states were restored before cleanup. One throwaway inbox rule was
  created and removed; the eight original rules ended with identical ids, order, priorities,
  enabled states and revisions. Nothing was permanently deleted.
- 2026-10-07 (native signatures): four synthetic signatures created and deleted. New-message and
  reply defaults were temporarily switched, the new-message default was cleared, and deletion of
  a selected fixture left a dangling pointer that was repaired. Final list/default values matched
  the original baseline, both original signature records were unchanged, and all temporary contents
  were absent. No message writes or sends occurred in these signature tests.
- No browser cookie, token or canary was ever used by a probe. The capture review decoded token *claims* (audience, client and scope names) only.

## Sources

- [MSAL Python](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens) · [Entra error codes](https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes)
- [Graph mail overview](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0) · [List messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0) · [Get message / MIME](https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0) · [Attachments](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0)
- [Immutable IDs](https://learn.microsoft.com/en-us/graph/outlook-immutable-id) · [Folder delta](https://learn.microsoft.com/en-us/graph/api/mailfolder-delta?view=graph-rest-1.0) · [Message delta](https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0)
- [`$search`](https://learn.microsoft.com/en-us/graph/search-query-parameter) · [Microsoft Search for messages](https://learn.microsoft.com/en-us/graph/search-concept-messages) · [JSON batching](https://learn.microsoft.com/en-us/graph/json-batching)
