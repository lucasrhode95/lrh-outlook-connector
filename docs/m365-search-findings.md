# Microsoft 365 search findings

**Updated:** 2026-10-12. This is the cross-product index of search surfaces, routes, result navigation, observations, and remaining evidence gaps. Detailed request/response notes remain in the product research files linked below. Findings describe one signed-in test profile and point-in-time captures; permission outcomes and result counts are not universal product limits.

**Product ownership:** Outlook owns email/attachments and Exchange calendar events, including
agenda, organizer, invitees, recurrence, event attachments, and the join link. Teams owns a related
Teams meeting's chat, recordings, transcripts, recap, and meeting notes. Verified references connect
these resources; a virtual event on another platform does not become a Teams meeting. The separate,
read-only SharePoint connector owns general file discovery/retrieval across OneDrive and SharePoint. Storage
overlap does not change ownership of email, message, or meeting relationships.

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
| `POST /outlookgatewayb2/graphql?cv=...&n=...` | Outlook's internal GraphQL service. During the Teams Recaps capture, operations fetched teammate-calendar groups, app tabs/entitlements, working hours, user auth settings, and a scheduling-service meeting response. | Ten HTTP 200 calls; response data roots are documented in Outlook research §3.9. This is a meeting-support path, not Microsoft Graph or a global search provider. |
| Outlook people/persona routes: `/api/v1/livepersonacard/configuration`, `/api/v1/peopleInsights/owareadingpane`, `/api/v2/personaphoto`, `/api/v2/personacards/preparePersona` | Profile/persona configuration, photo, and reading-pane support calls appeared in the Outlook capture set. | Respective outcomes: 3×200, 3×200, 23×200 plus 3×404, and 3×200. These were not tied to one suggestion/result activation and are not search submissions; see Outlook research §3.9. |
| Microsoft Graph `GET /v1.0/me/messages?$search=...` | Connector mailbox search. Conversation grouping/ranking is applied by the connector; matching message IDs remain available for hydration. | Implemented and exercised; `$search` semantics and immutable-ID translation are documented in [Outlook API research](outlook-api-research.md). |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["message"]` | Separate Graph Search API experiment. | HTTP 200 for a synthetic mail query; not selected for connector search because its result/folder semantics do not match the current behavior. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["event"]` | Calendar-event search experiment. | HTTP 403 under the tested profile, which lacked `Calendars.Read`; this is account-specific evidence, not a general restriction. `/me/calendarView` was also inaccessible under that profile. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["driveItem"]` | Shared file-name and content search; see [SharePoint and OneDrive](#sharepoint-and-onedrive). | Filename/content known hits, selected-item hydration/download, and distinct pages were verified for personal OneDrive and a shared library. A site/folder-restricted query returned zero in an earlier sample; broad recall and minimum permissions remain unverified. |
| Microsoft Graph `POST /v1.0/search/query` with `entityTypes: ["person"]` | People-search experiment. | HTTP 200 and two hits in the tested profile. |
| `POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail` | Outlook Web Service (OWS) action endpoint. Actions observed elsewhere in the connector include folder/item/conversation reads and changes. | `ExecuteSearch` variants returned HTTP 400; OWS was not established as a search mechanism. Full OWS action records are in [Outlook API research](outlook-api-research.md). |
| Microsoft Graph `GET /v1.0/me/calendarView?...` | Calendar-view route considered for event retrieval. | Not accessible with the tested profile; this is separate from Search API event search. |

### Outlook result selection and context

The connector can identify the exact matching mail by message ID and its conversation by conversation ID. It can then retrieve just that message or the full conversation. We did not establish that clicking a result in OWA uses the same API or hydration sequence. Mail search does not cover calendar events, SharePoint files, Teams messages, or meeting transcripts by itself.

### Standalone calendar and Teams-message replay

[Outlook research §3.14](outlook-api-research.md#314-owa-search-v2-calendar-and-teams-message-replay-2026-10-10)
records a later operation-level result that supersedes the calendar-only Graph blocker above.
App-owned One Outlook Web credentials silently acquired the Search resource token and successfully
called `POST https://outlook.office.com/searchservice/api/v2/query` for Event/Exchange and
Message/Teams. Event offsets advanced to distinct results; OWS `GetItem` hydrated a selected event's
body, attendees, and attachment. Search attachment flags and recurring-event totals were unreliable
in the observed sample. The Outlook connector now exposes bounded event discovery and event hydration;
synthetic adapter coverage exists, while live invocation through the connector and fresh-account
consent remain unverified. See the [Outlook roadmap](outlook-roadmap.md).

The Teams-message replay returned paginated hits with a verified known-chat identity for one target;
a second target's identity bridge remains unresolved. The Teams connector now implements bounded
online Message/Teams discovery on this shared Outlook-hosted Search v2 route. A live call through the
connector and direct hydration of returned hit IDs remain unverified. Transcript discovery,
retrieval, and full-text search remain separate unresolved capabilities.

Further detail: [Outlook API research](outlook-api-research.md) and [Outlook roadmap](outlook-roadmap.md).

## Teams

**Graph link:** Graph message-history and transcript APIs are listed under [Microsoft Graph](#microsoft-graph-cross-product-api-product). Teams global Search v2 and its MCPS recap/transcript routes are separate Teams services.

### Search and meeting surfaces

- The global Teams search bar exposed **All, Messages, Files, People, Channels, Group chats, and Meetings**. No Transcript tab was visible; an overflow menu was present but its labels were not captured.
- Teams global search submitted to `POST https://teams.microsoft.com/searchservice/api/v2/query`, whose bodies used `EntityRequests` and sometimes `AnswerEntityRequests`. A separate Teams route, `POST /search/api/v2/query`, carried three `Transcript` queries with `ContentSources=Teams`; one structured response had zero results and two response bodies could not be interpreted. This does not establish that transcript search is unsupported.
- The separate Meet-app Event builder also posts to `/search/api/v2/query` with lowercase `entityRequests`, date bounds, proximity sort, `from=0`, and `size=30`; source code has not tied that builder to the global Meetings filter. The route family therefore has distinct payload/workflow variants.
- A later capture showed File, People, Message, and Event categories. The separate Meet-app Event query was not proven to back the global-search Events category. Teams' global Substrate search is distinct from Graph Search API and direct SharePoint search.
- One parseable mixed response contained eight Message, three File, and five Event hits. Message `Source` fields included `InternetMessageId`, conversation/client-thread identities, and sent/received/created timestamps; File `Source` fields included SharePoint site/list/item pointers; Event `Source` included event ID and start/end times. A Teams Web Message-result normalizer uses `InternetMessageId` for the UI result id and builds `primaryAction.navigationId` from reply-chain identity, `ClientThreadId`, and `DateTimeSent`; the click handler uses those values to address the hit. The capture confirms those core input fields, but `ParentMessageId` and `ThreadSubType` used in reply-chain derivation were not top-level Source keys, and the exact component caller is not fully traced. This identifies and dates a captured message hit, but not a history-page number. The identifier has not been cross-checked against Graph/Chat Service history IDs, and the history-fetch wire request remains uncaptured.
- A common query with the Meetings filter returned 25 rows in the observed UI. A calendar-event detail/expanded view exposed agenda, participants, and an attachment.
- The Recaps surface loaded transcript tabs for both known recap pages. In-pane transcript search returned matches in one recap. The underlying transcript-viewer search request was not mapped. A download control was disabled for one recap and indicated host permission was needed; this is distinct from viewing/searching the transcript.
- The Recaps capture also links meeting-resource storage metadata to follow-up routes: for three captured artifacts, MCPS `driveId`/`driveItemId` values match cached ODSP media requests and Graph DriveItem `/preview` requests. This establishes ID handoff for those recap artifacts, but does not map global Search v2 file hits to Graph IDs. The same workflow fetched video visual frames from SharePoint media routes; these are playback hydration, not search results.

### Teams and meeting endpoint inventory

The Teams connector now exposes bounded online Message/Teams discovery through Search v2 and the
shared `m365.search` profile. Request mapping and continuation have synthetic coverage; a live MCP
call and direct hydration of a returned hit ID remain unverified. Recap and transcript capabilities
remain separate open work. See the [Teams search roadmap](../../lrh-teams-connector/docs/teams-connector-roadmap.md#p22--search)
and [adapter research](../../lrh-teams-connector/docs/teams-api-research.md#teams-online-message-search-adapter-2026-10-12).

| Endpoint or route family | Purpose and observed shape | Evidence / limit |
| --- | --- | --- |
| `POST https://teams.microsoft.com/searchservice/api/v2/query` | Teams global Search v2; request fields included `EntityRequests` and sometimes `AnswerEntityRequests`, with `File`, `People`, `Message`, and `Event` categories. | Captured submitted global searches. Results include `EntitySets` and/or `AnswerEntitySets`; no later page offset was captured. |
| `POST https://teams.microsoft.com/search/api/v2/query` | A distinct Teams route used for three `Transcript` queries with `ContentSources=Teams`. | All returned HTTP 200; one structured response had zero hits, two bodies were not interpretable. This does not establish unsupported transcript search. |
| `POST /search/api/v1/recommendations` | Teams Recaps recommendation flow requested `MeetingCatchUp`; interpreted responses included `EntitySets[].MeetingInsights`. | Five HTTP 200 calls in the Recaps capture; no continuation request using the observed recap pagination identifier was captured. This is not the global result-query route. |
| `POST /search/api/v1/suggestions` and `GET /search/api/v1/suggestions` | Search-box suggestions vary by surface and request shape. | Teams global search used POST; SharePoint used GET; Outlook used POST. These are suggestion routes, not submitted result searches. |
| `POST /search/api/v1/suggestions?scenario=powerbar` | Captured Teams Powerbar People suggestion workflow. | Nine HTTP 200 POST entries; one interpretable response returned one People suggestion. This differs from Middle Tier People Picker routes and from global Search v2 results; full result completeness/paging is unknown. |
| Meet-app `POST /search/api/v2/query` | Separate Event search builder; lowercase `entityRequests`, date bounds, proximity sort, `from=0`, `size=30`; parses `EntitySets[0].ResultSets[0].Results[].Source`. | Teams client source observation, not proven to back the global Meetings tab. |
| `GET https://teams.microsoft.com/api/mcps/{region}/user/meetings/contents?startDateTime=...&endDateTime=...&includeTypes=...&top=...` | Lists meeting contents. Observed `includeTypes` values: `MeetingChat`, `MeetingStage`, `Notes`, `Recording`, `TranscriptV2`. | Four HTTP 200 responses; returned `value` plus `paginationInfo.hasNextPage`; advancing this endpoint's page was not verified. |
| `GET https://teams.microsoft.com/api/mcps/{region}/contents/?threadId=...` | Meeting resource lookup. Thirteen captured HTTP 200 calls used `threadId` with `recapCallId`, `iCalUid`, `organizerId`, `organizerEmail`, and/or `isSeriesMeeting`. | Resource records included `addedBy`, `additionDateTime`, `id`, `mcoResourceId`, `metadata`, `subType`, `type`, and `webUrl`. Metadata keys observed are recorded in the Teams research file. |
| `POST https://graph.microsoft.com/v1.0/drives/{driveId}/items/{itemId}/preview` | DriveItem preview route used during Recaps. | Four HTTP 200 POSTs plus two successful OPTIONS preflights across three drive/item pairs; all three pairs match MCPS resource metadata. This is preview/context hydration, not Search API, and no body/response contract was established. |
| `GET /_api/v2.1/sites/{siteId}/drive/items/{itemId}/media/visualSegmentFrames/{frameId}/content` | SharePoint media frame retrieval during Recaps. | Twenty HTTP 200 JPEG GETs and twenty OPTIONS preflights; no query keys. Video display hydration, not global file/transcript search. |
| `POST https://teams.microsoft.com/api/mcps/{region}/meetingcatchup/multilingual/aicontent` | Recap/AI-content retrieval. | Captured body keys: `EntityRequests`, `Flights`, `Scenario`, `TargetLanguage`; response included `EntitySets[].MeetingInsights`. |
| `GET /api/mcps/{region}/collab/readcollabobject/V2/{opaquePath}` | Recaps collaboration-object hydration. | Five HTTP 200 responses; response included resource metadata fields. Opaque path composition is not established. |
| `GET /api/apps/part/{region}/beta/chats/{chatId}/tabs` | Recap-associated chat tab discovery. | One HTTP 200 response with a `value` envelope; not a message-history request. |
| `GET /api/chatsvc/{region}/v1/threads/{threadId}/consumptionhorizons` | Thread-level metadata fetched during recap navigation. | Two HTTP 200 calls; no captured response field establishes the message page/window for a search hit. |
| `GET /api/teamsevents/{region}/beta/events/{eventId}` and `GET /api/mt/part/{region}/v2.0/me/calendars/events/{eventId}` | Event and calendar-event detail hydration. | One Teams Events call and two Middle Tier calendar calls returned HTTP 200; response schemas and attendance meaning are not established. |
| `GET /api/csa/{region}/api/v1/teams/users/me/pinnedChannels` | Teams shell channel-pin state. | Two HTTP 200 calls in the global-search capture; not a Search v2 result request or chat-list GraphQL response. |
| `POST /ups/{region}/v1/presence/getpresence` and `/ups/global/v1/presence/getpresence` | Teams presence support calls. | In the Recaps capture, two regional requests returned HTTP 200 and two global-alias requests returned HTTP 404; this is capture-specific service discovery, not a general restriction. |
| `POST /ups/{region}/v1/pubsub/subscriptions/{subscriptionId}` | Teams shell subscription activity. | Three HTTP 200 calls total: two in the Recaps capture and one in the global-search capture; payload semantics are unresolved. |
| `PUT /api/mt/part/{region}/v1/schedulingService/meetings` | Adjacent meeting-scheduling update route. | Five HTTP 200 calls observed in Recaps with a `SchedulingServiceUpdateUrls` body field; the write route was not replayed. |
| Outlook `POST /outlookgatewayb2/graphql?cv=...&n=...` | The same Recaps browser session made private OWA GraphQL calls for calendar availability, meeting response, and tab/entitlement data. | See Outlook section above. This cross-product call does not establish that Outlook GraphQL performs Teams Search or transcript retrieval. |
| Recap builders `getMeetingRecapCatchupItems` and `getMeetingRecapCallSummary` | Substrate recap catch-up/summary request families; `QueryParameters` carried entity type/id/context and optional time bounds, top count, and meeting-thread filter. | Source names and fields observed; the exact caller/pagination path for `PaginationSessionId` remains unproven. These are distinct from calendar Event search and MCPS AI content. |
| `GET https://teams.microsoft.com/_api_cached/v2.1/drives/{drive}/items/{item}/cdnmedia/transcripts?...` | Cached transcript-family request. Query key names included `cTag`, `P1`–`P4`, `isformatjson`, `transcriptkey`, `applyhighlights`, `applymediaedits`, and `kid`. | Captured route shape only; parameter meanings, continuation, and full response behavior remain uncertain. |
| `POST /api/mt/part/{region}/beta/users/fetch` and `fetchShortProfile` | Middle Tier profile resolution; both took a JSON-array body and returned an object with a `value` array. | One HTTP 200 call per route in the Teams search capture. The `fetch` response included profile field names; these are lookup helpers, not free-form People Picker searches. |
| `POST /api/mt/part/{region}/beta/users/fetchfeaturesettings` | Batch feature-setting lookup by `Mri`; returns an array with per-entry `status`. | Three HTTP 200 calls; values and status meaning were not retained or interpreted. |
| `GET /api/mt/part/{region}/beta/users/{userId}/profilepicturev2/{mri}` and `GET /api/mt/part/{region}/beta/users/{userId}/mergedProfilePicturev2` | Profile-image hydration with display/size/etag query-key variants. | Successful image responses (JPEG/PNG); these hydrate People results and are not search calls. |
| `/drives/{drive}/items/{item}/media/transcripts/{transcript}/content` | Transcript-content route assembled by the Teams client bundle. A chat-history embedded URI used two additional path segments. | Static source observation, not a captured request; method, auth, and response remain unverified. Do not treat as an endpoint contract yet. |
| `/v1/objects/{object}/views/transcript`, `/v1/objects/{object}/views/transcript/status`, and `/MeetingIntelligence/TranscriptService/storage/{storage}/scope/{scope}/` | Legacy AMS and Meeting Intelligence transcript route families; the storage/scope builder adds `isChannel=true` for channel meetings. | Teams bundle source observations, not confirmed requests in the captured workflow. |
| `https://teams.microsoft.com/l/message/{chatId}/{messageId}?context={...}` | Documented Teams deep link to a chat message. | Microsoft-documented navigation form. User-observed click-to-hit behavior is consistent with message-ID anchoring; the actual client click handler and follow-up network request were not captured. |
| `https://teams.microsoft.com/l/message/{channelId}/{messageId}?tenantId=...&groupId=...&parentMessageId=...&teamName=...&channelName=...&createdTime=...` | Documented Teams deep link to a channel message. | Same limit: URL contract is documented; internal surrounding-window loading sequence is not captured. |

### Clicking a Teams message hit

The observed product behavior jumps to the selected message in its conversation. The 2026-10-10 Teams bundle gives a concrete client-side explanation: the search result activation path parses the result's primary navigation reference and calls `chat.chatById` with `replyChainId`, `messageId`, `dateTimeSent`, a generated `searchSessionId`, and highlight text; the channel path calls `channel.channelMessagesById` with channel/team identity, `replyChainId`, `deepLinkId`/message identity, and highlight options. A separate deep-link handler uses `focusOnLoad=DeeplinkedMessage`. The client can therefore address the hit by identity and timestamp without a page number.

No message-history HTTP call or event-level click trace appears in the supplied HARs. The exact request chain and the conversion of `dateTimeSent` to the existing Chat Service `startTime` parameter remain unverified. Treat the source-level navigation fields as established, and the timestamp-to-wire-query mapping as an inference. See [the Teams capture/source audit](../../lrh-teams-connector/docs/teams-api-research.md#search-result-activation-and-exact-message-navigation-source-2026-10-11-audit).

The captured Teams Search v2 route and the message deep link serve different steps: Search v2 discovers results; the deep link identifies a destination message. They should not be collapsed into one endpoint. The in-pane transcript search is a third workflow whose request route remains unknown.

Further detail: [Teams API research](../../lrh-teams-connector/docs/teams-api-research.md) and [Teams roadmap](../../lrh-teams-connector/docs/teams-connector-roadmap.md).

## SharePoint and OneDrive

**Graph link:** the file search here uses Graph Search API `driveItem`; see [Microsoft Graph](#microsoft-graph-cross-product-api-product) for the shared request family and its differences from mail and Teams search.

SharePoint/OneDrive file search was investigated separately from Outlook mail and Teams global search. The observed Graph route was `POST /v1.0/search/query` with `entityTypes: ["driveItem"]`; synthetic filename-only and content-only queries returned hits. The read-only SharePoint connector now uses this route, preserves returned drive/item identity, hydrates selected metadata, and supports bounded downloads. Live checks found synthetic known items in personal OneDrive and a shared SharePoint library, hydrated their metadata, downloaded matching bytes, and returned distinct pages for explicit offsets. The account's permissions exceed the connector's requested Graph scopes, so minimum permissions and exhaustive recall remain unverified; see the [SharePoint roadmap](../../lrh-sharepoint-connector/docs/roadmap.md).

The web routes now resolve by surface. The OneDrive Files flow submitted Substrate `POST /search/api/v2/query` for `File` and also used the SharePoint REST proxy `POST /_api/sp.searchresultitemcollection.getsearchresultsasstream`; SharePoint list-stream calls loaded list content. The company SharePoint search page submitted `POST /searchservice/api/v2/query` for `File` and `People`, with `GET /search/api/v1/suggestions` for suggestions and separate init/configuration calls. The response envelope carried Search result sets, but these private routes are distinct from Graph `driveItem` search. See [Outlook research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10) for counts and route details.

The SharePoint search capture also contains 24 HTTP 200 `GET /api/v1/documents/searchpreview` requests returning JPEG images. In both relevant HARs, all 24 preview requests' `uniqueId`, `siteId`, `webId`, and `listId` values matched those four fields on a captured Search result `Source` record. This proves item-level association between the preview and a search result, but not that a click caused each request (the page may preload previews), not file-content opening, and not a mapping to Graph drive/item IDs. For agent-directed expansion, retain the direct Graph drive/item identity and `webUrl` when a Graph result provides them.

The SharePoint client source adds two browser-side paging contracts: a scoped File request builder uses `From=0`, `Size=5` and a 10-minute client cache keyed by search text; a Lists Home provider requests `Size=24`, advances `From` by 24 while `MoreResultsAvailable` is true, and parses the first result set. The source callers were not matched one-to-one to HAR requests; in particular, the captured company SharePoint result submissions used `/searchservice/api/v2/query`, while these builders target `/search/api/v2/query`. Only offset-zero requests appeared in the capture set, so the next-page path is source-established, not capture-verified. The result parser retains SharePoint navigation fields such as `AccessUrl`, `SPWebUrl`, `SpoUniqueId`, `SiteId`, `WebId`, `ListId` and `ParentLink`; this still does not map the browser item to Graph drive/item identifiers. The preview source builds an image request with a `ClientScenario` header and returns a blob URL, not document content. See [Outlook research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10).

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
| `POST /v1.0/drives/{driveId}/items/{itemId}/preview` | OneDrive/SharePoint DriveItem preview. | Four HTTP 200 POSTs plus two successful OPTIONS preflights across three items whose drive/item identifiers were also returned in MCPS meeting-resource metadata. This is an item-level preview route, not Graph Search; file-search result ID equivalence remains unproven. |

### Context expansion implications

For Teams message hits, preserve message ID, chat/channel identity, and timestamp. Prefer the Teams deep link when the agent/user wants to open the UI at the exact hit. For programmatic context, fetch by chat/channel identity and use opaque paging or a bounded date strategy; an exact page number was not found. Let the agent choose whether to hydrate a hit, fetch more surrounding messages, or stop. The connector should return provenance and useful identifiers rather than guessing which hits are promising.

Outlook mail already has a direct message-ID and conversation-ID expansion path. SharePoint/OneDrive hits should retain drive/item identity and a web URL when available so the selected file can be opened or fetched by its own route. These are related context workflows but use different resource identities and APIs.

## Measurement and provenance

Measurements below are point-in-time observations, not broad latency claims. Before online Teams search was implemented, a one-chat/one-day exact search fetched one history page in about one second. A five-page all-catalog search covered five pages across 966 catalog IDs in 8.7 seconds and remained incomplete. Four concurrent one-page history reads took 3.552 seconds versus 4.298 seconds sequentially, about 17% lower; each result was incomplete at the one-page limit, so this small sample does not establish a safe concurrency setting.

A synthetic SQLite FTS5 comparison across 14,720 and 100,000 messages found selective rare-substring searches improved from 15.36 to 3.82 ms and 44.06 to 4.64 ms, respectively. Common-substring searches regressed from 3.97 to 29.28 ms and 4.02 to 153.09 ms; common all-terms searches regressed from 4.01 to 31.80 ms and 3.87 to 185.98 ms. The FTS index increased database size by 30.5% and 34.1% and took 77 ms and 717 ms to build. Only three query shapes were compared, so this does not justify replacing the current matcher. An empty temporary-scope delete measured roughly 1–3 microseconds per search, within noise. Detailed methods and limits are in [Teams API research](../../lrh-teams-connector/docs/teams-api-research.md#connector-code-path-audit-and-measurements-2026-10-10).

Three Outlook one-hit `search_messages` calls took 3.409 seconds cold, then 0.760 and 0.784 seconds warm. They returned one conversation/message each and were incomplete at the one-hit limit. These end-to-end calls do not isolate Graph latency, and no same-query comparison against Graph Search API was made. See [Outlook API research §3.16](outlook-api-research.md#316-outlook-mail-search-timing-and-route-comparison-2026-10-10).

## Endpoint comparison: attachments and OneDrive/SharePoint files (2026-10-11)

For Microsoft Graph, email, Teams messages, OneDrive files, and company SharePoint files all use the same Search API URL, `POST https://graph.microsoft.com/v1.0/search/query`; the request chooses a different `entityTypes` value. This is one transport, not one result type or one permissions set. In particular, `chatMessage` cannot be combined with other entity types and must be searched in its own request.

| Content | Graph Search API entity | What it searches | Status in this investigation |
| --- | --- | --- | --- |
| Email and supported attachments | `message` | Message body plus supported attachment content; results are message hits. | Subject/body known-hit requests succeeded. Attachment-only content and which attachment matched were not live-tested. Outlook's current connector uses Graph mail `$search`; attachment-name filtering is separate. |
| Teams messages and supported attachments | `chatMessage` | Message body and attachments; supported attachment file types match SharePoint Online search. Results identify the message, not a documented attachment-specific hit. | Graph `chatMessage` search was denied for the tested profile. The Teams connector instead implements the separately captured Search v2 Message/Teams route; live invocation through that connector remains unverified. |
| OneDrive and company SharePoint files | `driveItem` | Files/folders within content the signed-in user can read, across OneDrive and SharePoint. | Implemented in the read-only SharePoint connector. Filename/content known hits, selected-item hydration/download, and two explicit pages succeeded for the tested locations. Exhaustive recall and minimum permissions remain unverified. |

Two routes should not be confused with that shared Search API. The Outlook connector currently uses `GET /v1.0/me/messages?$search=...`; unfielded terms target sender, subject, and body, while KQL `attachment:` targets attachment **names**. Full-text attachment search is documented on Search API `message`. For files, Graph also offers drive-scoped `GET /me/drive/root/search(q='...')` and broader accessible-item `GET /me/drive/search(q='...')`; those are separate GET routes. Teams and SharePoint web clients also use private Substrate/SharePoint routes recorded above, which are not Graph endpoints.

These results distinguish shared transport from product ownership: Graph Search API uses one URL with a different `entityTypes` value, but each entity has different permissions, coverage, and result shape. Outlook owns email attachments, Teams owns message attachments, and the SharePoint connector owns general OneDrive/SharePoint file search and retrieval. Its selected-file path retains drive/item identity and a web URL when available; the agent decides whether to hydrate or download a result.

Primary references: [Outlook message search](https://learn.microsoft.com/en-us/graph/search-concept-messages), [Teams message search](https://learn.microsoft.com/en-us/graph/search-concept-chat-messages), [OneDrive and SharePoint search](https://learn.microsoft.com/en-us/graph/search-concept-files), [Graph Search API permissions and page limits](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0), and [Graph `$search` mail properties](https://learn.microsoft.com/en-us/graph/search-query-parameter).

## Primary references

- [Outlook API research](outlook-api-research.md)
- [Outlook roadmap](outlook-roadmap.md)
- [SharePoint roadmap](../../lrh-sharepoint-connector/docs/roadmap.md)
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
