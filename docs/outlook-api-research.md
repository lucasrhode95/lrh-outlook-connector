# Outlook API research

**Evidence dates:** 2026-10-01 (first standalone probes) · 2026-10-02 (browser capture review, probe suite)  
**Product scope:** [Requirements v4](outlook-requirements-v4.md) · **How it is built:** [Architecture](architecture.md) · **Probes:** [`research/`](../research/README.md)

This is the single record of what Microsoft's APIs allow and how they behave for this tenant (Landis+Gyr, one user mailbox). It keeps evidence and conclusions only. How to re-run the probes and take captures is in `research/README.md`.

| Label | Meaning |
|---|---|
| **STANDALONE** | Called by our own process with app-owned tokens. No browser credentials. |
| **BROWSER** | Seen in an Outlook Web capture (response side only). |
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
| Attachments: list, item (`$select` must not include `@odata.type`), `$value`, `$select=microsoft.graph.fileAttachment/contentId` | 200 |
| `/me/mailFolders/delta`, `/me/mailFolders/{id}/messages/delta` (`odata.maxpagesize`) | 200, `deltaLink`, no-change replay returns 0 |
| `/me/translateExchangeIds` | 200. Converts the regular ids `$search` returns (`AQMk…` and `AAMk…`) into immutable ids (live 2026-10-03, read sign-in); the results equal the ids listings return (3 of 3, live 2026-10-04). An input that is already immutable fails the whole call (HTTP 400 `InvalidArgument`, expected `EntryId`). Not needed for OWS (§4.2). |

**Sign-in (live 2026-10-03):** `auth write` asked for its own device code right after `auth read`: One Outlook Web does not reuse Outlook Mobile's sign-in, so two sign-ins are needed. The read token carries 18 Graph scopes (mail: `Mail.Read`, `Mail.Read.Shared`); the write token 74 Outlook scopes (mail: `Mail.ReadWrite(.All/.Shared)`, `Mail.Send(.Shared)`).

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

### 3.3 Search (`search.py`)

Three pages of 25 per backend.

| Query | Graph `$search` | Graph `/search/query` | Substrate (§5) |
|---|---|---|---|
| `relatório be` | 75 msgs → 52 conversations, more available | total 238 | 75 conversations, total 340 |
| `subject:relatório` (complete set) | 63 msgs → **50 conversations** | total 65 | **50 conversations**, total 65 |

- All three backends fold accents (`relatório` ≡ `relatorio`).
- After resolving Substrate's `ImmutableId`s through Graph, `subject:relatório` returned the **identical 50 conversations** from both. For `relatório be`, all 52 `$search` conversations appear in Substrate's top 75.
- **Conclusion:** adopt Graph `$search`, and group hits by `conversationId` locally. `/search/query` is not used: its total comes from another engine and ignores the folder rules.

### 3.4 Conversations and reply headers (`threads.py`)

Sample: 25 recent conversations, 192 messages.

| Measure | Result |
|---|---|
| Conversation filter | Returns every folder: 21 of 25 conversations span 2–4 folders |
| + `$orderby` | **400 `InefficientFilter`** → sort locally |
| Size | median 4, max 40 messages |
| Headers | `Message-ID` 137, `Thread-Index` 136, `In-Reply-To`/`References` 113. **55 messages have no headers at all** (drafts and the user's own messages). |
| Reply tree | 6 of 25 conversations branch (18 branch points). 9 parents are not in the mailbox. |
| `uniqueBody` / `body` | median length ratio 0.10 |

**Conclusion:** `get_thread` = conversation filter + local sort. Branch detection is feasible for received mail. The user's own messages need a fallback.

### 3.5 Attachments (`attachments.py`)

40 messages, 176 attachments.

| Measure | Result |
|---|---|
| Types | `fileAttachment` 168, `itemAttachment` 8, `referenceAttachment` 0 |
| Inline | 125 of 168 file attachments. All are referenced as `cid:` in the full body; only 47 in `uniqueBody`. |
| `$value` | Works for file attachments, and for item attachments (returns MIME → `.eml`) |
| Sizes | median 11 KB, max 27.5 MB |

**Conclusion:** export non-inline attachments by default. Include inline images only when the rendered body references them.

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

Two captures (2026-10-04). In the first, the owner captured Outlook on the web's rules page (`Settings → Mail → Rules`) while creating, editing, disabling and deleting a throwaway rule ("ZZ connector test rule": From `nobody@example.invalid`, Subject includes `zz-connector-test`, Move to a folder, Mark as read, Stop processing more rules), then reordering two rules and back. In the second, two more throwaway rules: Sent to, Subject includes, Move to Archive; disabled, re-enabled, conditions cleared, moved to a user folder, renamed, deleted. Shapes below are from that capture with every value redacted; the capture itself (it held tokens) was deleted after analysis. Not yet replayed by the connector.

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
- **Folders:** in requests, `MoveToFolder.RawIdentity` is a base64 folder id starting `AQMk` (116–120 characters; Archive and a user folder), the same family as the regular ids `$search` returns (H12); whether it equals the Graph folder id that `list_folders` returns, as is or in its OWS form (§7), is not yet checked. In answers, `MoveToFolder.RawIdentity` is a different form: `<organisation path>/<mailbox GUID>:\<folder name>` (136–140 characters, the folder's own name only, e.g. `:\Archive`), so reading a rule's target folder means matching by name (ambiguous when two folders share a name) or by `DisplayName`.
- **Other conditions and flags (second capture):** `SentTo` has the same `PeopleIdentity` shape as `From` (read back with `Address` and `AddressOrigin` added). `NewInboxRule` from the UI also carried `DisplayAlert: "Default"` or `PlaySound: "Default"` depending on the options shown; neither is needed.
- **A rule has 90 fields** (conditions, `ExceptIf…` exceptions, actions, `Description` texts, `InError`, `SupportedByTask`, `RuleProvider`). The owner's 8 rules use only `MoveToFolder` (8), `SentTo` (6), `SubjectContainsWords` (5), `From` (1), `SubjectOrBodyContainsWords` (1), all with `StopProcessingRules`, all enabled, none in error.
- Outlook also calls `GetMailboxByIdentity` after each change and `/ows/v1.0/OutlookOptions/MailForwardingNotification` on the rules page; neither is needed to manage rules.

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
- No browser cookie, token or canary was ever used by a probe. The capture review decoded token *claims* (audience, client and scope names) only.

## Sources

- [MSAL Python](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens) · [Entra error codes](https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes)
- [Graph mail overview](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0) · [List messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0) · [Get message / MIME](https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0) · [Attachments](https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0)
- [Immutable IDs](https://learn.microsoft.com/en-us/graph/outlook-immutable-id) · [Folder delta](https://learn.microsoft.com/en-us/graph/api/mailfolder-delta?view=graph-rest-1.0) · [Message delta](https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0)
- [`$search`](https://learn.microsoft.com/en-us/graph/search-query-parameter) · [Microsoft Search for messages](https://learn.microsoft.com/en-us/graph/search-concept-messages) · [JSON batching](https://learn.microsoft.com/en-us/graph/json-batching)
