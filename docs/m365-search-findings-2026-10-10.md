# Microsoft 365 search investigation findings — 2026-10-10

This is the cross-product index for the investigation. It keeps each product's UI, routes, result navigation, observations, and open questions together, then links shared Graph mechanisms across products. Detailed request/response notes remain in the canonical research files linked below. Findings describe one signed-in test profile and point-in-time captures; permission outcomes and result counts are not universal product limits.

## Outlook

**Graph link:** the Outlook connector's mail route and additional Graph entity probes are documented under [Microsoft Graph](#microsoft-graph-cross-product-api-product). Outlook top-bar Search remains a separate Substrate-backed workflow.

### Search surface and workflow

- The Outlook Web top-bar search was compared with the connector's mail search. For one subject query, both returned the same 50 conversations; for one two-term query, all 52 Graph `$search` results appeared in the first 75 Substrate results. This is a small comparison, not a general parity claim.
- The connector's mail path is Microsoft Graph message `$search`, not Graph Search API. Results are grouped/ranked by conversation and expose matching messages. A hit's message ID identifies the excerpt; `get_message` retrieves that message and `get_conversation` retrieves its full conversation oldest-first.
- The OWA source-bundle review found Search v2 request-wrapper, suggestion, and initialization builders, but did not tie a concrete route builder or result-click handler to the OWA source. Separate browser captures recorded the shared Substrate Search v2 route below. We did not capture the network request caused by clicking an OWA result. The supplied OWA bundle is not evidence about Teams Web.
- OWA mail search and the connector's Graph `$search` are distinct mechanisms: the first uses Outlook's Substrate-backed UI path; the connector uses Graph mail search. Graph `POST /search/query` was probed separately and is not the connector's mail-search implementation.

### Outlook-associated endpoints and evidence

| Endpoint or route family | Purpose and observed shape | Evidence / limit |
| --- | --- | --- |
| `POST /searchservice/api/v2/query` (Microsoft Substrate Search v2) | Global Outlook/Teams Web search. Bodies included `EntityRequests` and, in some flows, `AnswerEntityRequests`. | Captured browser route; UI-category mapping, full pagination semantics, and click-through behavior remain incomplete. See [Outlook research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10). |
| Microsoft Graph `GET /v1.0/me/messages?$search=...` | Connector mailbox search. Conversation grouping/ranking is applied by the connector; matching message IDs remain available for hydration. | Implemented and exercised; `$search` semantics and immutable-ID translation are documented in [Outlook API research](outlook-api-research.md). |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["message"]` | Separate Graph Search API experiment. | HTTP 200 for a synthetic mail query; not selected for connector search because its result/folder semantics do not match the current behavior. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["event"]` | Calendar-event search experiment. | HTTP 403 under the tested profile, which lacked `Calendars.Read`; this is account-specific evidence, not a general restriction. `/me/calendarView` was also inaccessible under that profile. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["driveItem"]` | Shared file-name and content search; see [SharePoint and OneDrive](#sharepoint-and-onedrive). | Synthetic uploaded-file query returned hits. Site/folder-restricted query returned zero in that sample; this does not prove the scope is unsupported. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["person"]` | People-search experiment. | HTTP 200 and two hits in the tested profile. |
| `POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail` | Outlook Web Service (OWS) action endpoint. Actions observed elsewhere in the connector include folder/item/conversation reads and changes. | `ExecuteSearch` variants returned HTTP 400; OWS was not established as a search mechanism. Full OWS action records are in [Outlook API research](outlook-api-research.md). |
| Microsoft Graph `GET /v1.0/me/calendarView?...` | Calendar-view route considered for event retrieval. | Not accessible with the tested profile; this is separate from Search API event search. |

### Outlook result selection and context

The connector can identify the exact matching mail by message ID and its conversation by conversation ID. It can then retrieve just that message or the full conversation. We did not establish that clicking a result in OWA uses the same API or hydration sequence. Mail search does not cover calendar events, SharePoint files, Teams messages, or meeting transcripts by itself.

Further detail: [Outlook API research](outlook-api-research.md), [investigation plan](m365-search-investigation-plan.md), and [Outlook roadmap](outlook-roadmap.md).

## Teams

**Graph link:** Graph message-history and transcript APIs are listed under [Microsoft Graph](#microsoft-graph-cross-product-api-product). Teams global Search v2 and its MCPS recap/transcript routes are separate Teams services.

### Search and meeting surfaces

- The global Teams search bar exposed **All, Messages, Files, People, Channels, Group chats, and Meetings**. No Transcript tab was visible; an overflow menu was present but its labels were not captured.
- A captured global-search route was `POST https://teams.microsoft.com/search/api/v2/query`. Transcript queries included `EntityType=Transcript` and `ContentSources=Teams`. Three requests returned HTTP 200; one interpretable response had zero results and no further pages, while two response bodies could not be interpreted. This does not establish that transcript search is unsupported.
- The browser capture set also recorded Teams global-search traffic to the shared Substrate route `POST /searchservice/api/v2/query`, whose bodies used `EntityRequests` and sometimes `AnswerEntityRequests`. A separate Meet-app builder posts to `/search/api/v2/query` with lowercase `entityRequests`, date bounds, proximity sort, `from=0`, and `size=30`; it has not been tied to the global Meetings filter.
- A later capture showed File, People, Message, and Event categories. The separate Meet-app Event query was not proven to back the global-search Events category. Teams' global Substrate search is distinct from Graph Search API and direct SharePoint search.
- A common query with the Meetings filter returned 25 rows in the observed UI. A calendar-event detail/expanded view exposed agenda, participants, and an attachment.
- The Recaps surface loaded transcript tabs for both known recap pages. In-pane transcript search returned matches in one recap. The underlying transcript-viewer search request was not mapped. A download control was disabled for one recap and indicated host permission was needed; this is distinct from viewing/searching the transcript.

### Teams and meeting endpoint inventory

| Endpoint or route family | Purpose and observed shape | Evidence / limit |
| --- | --- | --- |
| `POST /searchservice/api/v2/query` (Microsoft Substrate Search v2) | Teams global search in captured flows; body fields included `EntityRequests` and sometimes `AnswerEntityRequests`. | Same route family recorded for Outlook global search above; category-to-request mapping and standalone app-owned authentication remain unproven. See [Outlook research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10). |
| `POST https://teams.microsoft.com/search/api/v2/query` | Teams global Search v2; request fields included `EntityType` and `ContentSources`, including transcript/Teams values. | Captured route and bounded response observations; complete schemas and all response bodies are not available. See [Teams API research](../../lrh-teams-connector/docs/teams-api-research.md). |
| Meet-app `POST /search/api/v2/query` | Separate Event search builder; lowercase `entityRequests`, date bounds, proximity sort, `from=0`, `size=30`; parses `EntitySets[0].ResultSets[0].Results[].Source`. | Teams client source observation, not proven to back the global Meetings tab. |
| `GET https://teams.microsoft.com/api/mcps/{region}/user/meetings/contents?startDateTime=...&endDateTime=...&includeTypes=...&top=...` | Lists meeting contents. Observed `includeTypes` values: `MeetingChat`, `MeetingStage`, `Notes`, `Recording`, `TranscriptV2`. | Returned `value` plus `paginationInfo.hasNextPage`; advancing this endpoint's page was not verified. |
| `GET https://teams.microsoft.com/api/mcps/{region}/contents/?threadId=...` | Meeting resource lookup. Captured query variants also used `recapCallId`, `iCalUid`, `organizerId`, `isSeriesMeeting`, and `organizerEmail`. | Resource records included `addedBy`, `additionDateTime`, `id`, `mcoResourceId`, `metadata`, `subType`, and `type`. Metadata keys observed are recorded in the Teams research file. |
| `POST https://teams.microsoft.com/api/mcps/{region}/meetingcatchup/multilingual/aicontent` | Recap/AI-content retrieval. | Captured body keys: `EntityRequests`, `Flights`, `Scenario`, `TargetLanguage`; response included `EntitySets[].MeetingInsights`. |
| Recap builders `getMeetingRecapCatchupItems` and `getMeetingRecapCallSummary` | Substrate recap catch-up/summary request families; `QueryParameters` carried entity type/id/context and optional time bounds, top count, and meeting-thread filter. | Source names and fields observed; the exact caller/pagination path for `PaginationSessionId` remains unproven. These are distinct from calendar Event search and MCPS AI content. |
| `GET https://teams.microsoft.com/_api_cached/v2.1/drives/{drive}/items/{item}/cdnmedia/transcripts?...` | Cached transcript-family request. Query key names included `cTag`, `P1`–`P4`, `isformatjson`, `transcriptkey`, `applyhighlights`, `applymediaedits`, and `kid`. | Captured route shape only; parameter meanings, continuation, and full response behavior remain uncertain. |
| `/drives/{drive}/items/{item}/media/transcripts/{transcript}/content` | Transcript-content route assembled by the Teams client bundle. A chat-history embedded URI used two additional path segments. | Static source observation, not a captured request; method, auth, and response remain unverified. Do not treat as an endpoint contract yet. |
| `/v1/objects/{object}/views/transcript`, `/v1/objects/{object}/views/transcript/status`, and `/MeetingIntelligence/TranscriptService/storage/{storage}/scope/{scope}/` | Legacy AMS and Meeting Intelligence transcript route families; the storage/scope builder adds `isChannel=true` for channel meetings. | Teams bundle source observations, not confirmed requests in the captured workflow. |
| `https://teams.microsoft.com/l/message/{chatId}/{messageId}?context={...}` | Documented Teams deep link to a chat message. | Microsoft-documented navigation form. User-observed click-to-hit behavior is consistent with message-ID anchoring; the actual client click handler and follow-up network request were not captured. |
| `https://teams.microsoft.com/l/message/{channelId}/{messageId}?tenantId=...&groupId=...&parentMessageId=...&teamName=...&channelName=...&createdTime=...` | Documented Teams deep link to a channel message. | Same limit: URL contract is documented; internal surrounding-window loading sequence is not captured. |

### Clicking a Teams message hit

The observed product behavior jumps to the selected message in its conversation. A message-ID deep link explains how the client can address that exact message without exposing a page ordinal. Microsoft documents `chatMessage.webUrl` as a link that opens the message in Teams; Graph also exposes the message ID plus chat or channel identity. This supports exact-hit navigation by identity, but it does **not** show which internal API Teams Web calls to load the surrounding window. That click handler/request remains an open capture item.

The captured Teams Search v2 route and the message deep link serve different steps: Search v2 discovers results; the deep link identifies a destination message. They should not be collapsed into one endpoint. The in-pane transcript search is a third workflow whose request route remains unknown.

Further detail: [Teams API research](../../lrh-teams-connector/docs/teams-api-research.md) and [Teams roadmap](../../lrh-teams-connector/docs/teams-connector-roadmap.md).

## SharePoint and OneDrive

**Graph link:** the file search here uses Graph Search API `driveItem`; see [Microsoft Graph](#microsoft-graph-cross-product-api-product) for the shared request family and its differences from mail and Teams search.

SharePoint/OneDrive file search was investigated as a file source, separately from Outlook mail and Teams global search. The observed route was Microsoft Graph Search API `POST /v1.0/search/query` with `entityTypes: ["driveItem"]`, using a synthetic filename query and a content-only query. The user confirmed that the synthetic file had uploaded before the search. Both queries returned hits, including results associated with multiple sites and OneDrive. A site/folder-restricted query returned no hits in that sample; more tests are needed before attributing that to indexing, permissions, or filter semantics.

The SharePoint/OneDrive web UI also used a private browser-search route: `POST /search/api/v2/query` with `EntityRequests[].EntityType=File`, query text, `From`, `Size`, content sources, sort, and fields. Search-box suggestions used `GET /search/api/v1/suggestions`. The response envelope included `ApiVersion`, `SearchTerms`, `EntitySets`, and result sets. This route and its suggestion path are distinct from Graph `driveItem` search; browser-to-Graph result identity mapping, continuation, and complete scope coverage remain open.

The file result click/open request was not captured, so we have not verified how the browser result maps to a Graph drive/item identity or whether opening it uses another SharePoint route. For agent-directed expansion, retain the Graph drive/item identity and `webUrl` when the Graph result provides them.

The same Graph Search API transport is used in the Outlook investigation for `message`, `event`, and `person`, but the `driveItem` entity is the file-search path. Teams' captured `search/api/v2/query` is a separate search service. Teams may display files in its Files results category, but the captured evidence does not establish that category's implementation is the same as the tested SharePoint/OneDrive Graph query.

Search route limits and request notes are in [Outlook API research](outlook-api-research.md), [§3.8](outlook-api-research.md#38-microsoft-365-search-api-coverage-review-2026-10-10) and [§3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10). Search API `driveItem` results are not evidence that every SharePoint library, OneDrive item, attachment, or Teams file is indexed and accessible.

## Microsoft Graph (cross-product API product)

| Shared mechanism | Products / data | How the workflows relate |
| --- | --- | --- |
| `POST /v1.0/search/query` | SharePoint/OneDrive `driveItem`; Outlook-side experiments for `message`, `event`, and `person`; Microsoft documents other entity types including `chatMessage`. | One Graph Search API envelope routes by `entityTypes`, but permissions, coverage, and result shapes vary by entity. It is not the same service as Teams Search v2 and is not the Outlook connector's current mail-search route. |
| `GET /v1.0/me/messages?$search=...` | Outlook mailbox messages. | The connector's current server-side mail search. For context, use the hit's message ID or its conversation ID with the connector's message/conversation retrieval operations. |
| `POST /v1.0/search/query` with `entityTypes: ["chatMessage"]` | Graph chat-message search. | Denied for the tested profile; account-specific, not a general Graph restriction. Separate from Teams Search v2. |
| `GET /v1.0/chats/{chatId}/messages` and `/v1.0/teams/{teamId}/channels/{channelId}/messages` | Teams chat/channel message history. | Context retrieval, distinct from search discovery and Teams' direct message link. Graph paging is opaque; list responses support up to 50 messages/page and do not provide a stable page number for a search hit. |
| `GET /v1.0/me/onlineMeetings/{meetingId}/transcripts` and `GET /v1.0/me/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content` | Transcript listing and per-meeting transcript content. | Documented Graph routes, separate from full-text discovery, Teams Search v2, and MCPS recap routes. Retrieval requires documented transcript permissions and applicable tenant settings; not live-verified in this investigation. |

### Context expansion implications

For Teams message hits, preserve message ID, chat/channel identity, and timestamp. Prefer the Teams deep link when the agent/user wants to open the UI at the exact hit. For programmatic context, fetch by chat/channel identity and use opaque paging or a bounded date strategy; an exact page number was not found. Let the agent choose whether to hydrate a hit, fetch more surrounding messages, or stop. The connector should return provenance and useful identifiers rather than guessing which hits are promising.

Outlook mail already has a direct message-ID and conversation-ID expansion path. SharePoint/OneDrive hits should retain drive/item identity and a web URL when available so the selected file can be opened or fetched by its own route. These are related context workflows but use different resource identities and APIs.

## Measurement and provenance

The connector measurements and their small-sample limits are in [the dated investigation log](m365-search-investigation-log-2026-10-10.md) and the product roadmaps. In brief: one Teams chat/day exact search required one remote history page and took about one second; a five-page all-catalog search covered five pages across 966 catalog IDs in 8.7 seconds and remained incomplete. Local FTS helped rare-term matching in the small benchmark and slowed common/all-terms queries while increasing database size, so this evidence does not justify switching the search design alone. Outlook mail `$search` timings were three observations (3.409 seconds cold, 0.760 and 0.784 seconds warm), too few for a broad performance claim.

## Endpoint comparison: attachments and OneDrive/SharePoint files (2026-10-11)

For Microsoft Graph, email, Teams messages, OneDrive files, and company SharePoint files all use the same Search API URL, `POST https://graph.microsoft.com/v1.0/search/query`; the request chooses a different `entityTypes` value. This is one transport, not one result type or one permissions set. In particular, `chatMessage` cannot be combined with other entity types and must be searched in its own request.

| Content | Graph Search API entity | What it searches | Status in this investigation |
| --- | --- | --- | --- |
| Email and supported attachments | `message` | Message body plus supported attachment content; results are message hits. | Subject/body known-hit requests succeeded. Attachment-only content and which attachment matched were not live-tested. |
| Teams messages and supported attachments | `chatMessage` | Message body and attachments; supported attachment file types match SharePoint Online search. Results identify the message, not a documented attachment-specific hit. | Microsoft documents the feature, but the tested profile's request was denied. This is account-specific; permission and known-hit tests remain. |
| OneDrive and company SharePoint files | `driveItem` | Files/folders within content the signed-in user can read, across OneDrive and SharePoint. | Known filename/content file probes succeeded, and an unrestricted query returned results from multiple SharePoint sites and OneDrive. Site/folder filters, exhaustive recall, and the minimum permission set remain unverified. |

Two routes should not be confused with that shared Search API. The Outlook connector currently uses `GET /v1.0/me/messages?$search=...`; unfielded terms target sender, subject, and body, while KQL `attachment:` targets attachment **names**. Full-text attachment search is documented on Search API `message`. For files, Graph also offers drive-scoped `GET /me/drive/root/search(q='...')` and broader accessible-item `GET /me/drive/search(q='...')`; those are separate GET routes. Teams and SharePoint web clients also use private Substrate/SharePoint routes recorded above, which are not Graph endpoints.

The file-search route is established well enough to start a bounded, read-only adapter. Remaining confirmation is about production scope and result use: app-token delegated scopes/consent, paging beyond the first page, hydrating one selected result by its returned drive/item identity, and distinct known filename/content probes for personal OneDrive and company SharePoint. Keep results user-readable only, preserve drive/item identity and `webUrl`, and let the agent decide whether to fetch a selected file.

Primary references: [Outlook message search](https://learn.microsoft.com/en-us/graph/search-concept-messages), [Teams message search](https://learn.microsoft.com/en-us/graph/search-concept-chat-messages), [OneDrive and SharePoint search](https://learn.microsoft.com/en-us/graph/search-concept-files), [Graph Search API permissions and page limits](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0), and [Graph `$search` mail properties](https://learn.microsoft.com/en-us/graph/search-query-parameter).

## Primary references

- [Outlook API research](outlook-api-research.md)
- [Outlook investigation plan](m365-search-investigation-plan.md)
- [Outlook roadmap](outlook-roadmap.md)
- [Teams API research](../../lrh-teams-connector/docs/teams-api-research.md)
- [Teams roadmap](../../lrh-teams-connector/docs/teams-connector-roadmap.md)
- [Microsoft Teams deep links](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams)
- [Microsoft Graph chat-message search](https://learn.microsoft.com/en-us/graph/search-concept-chat-messages)
- [Get a chatMessage](https://learn.microsoft.com/en-us/graph/api/chatmessage-get?view=graph-rest-1.0)
- [List chat messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0)
- [Search API overview](https://learn.microsoft.com/en-us/graph/search-concept-overview)
- [Search for driveItems](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0)
- [List online meeting transcripts](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-transcripts?view=graph-rest-1.0)
- [Get online meeting transcript content](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-get-transcript?view=graph-rest-1.0)
