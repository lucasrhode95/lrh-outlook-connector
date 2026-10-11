# Microsoft 365 Search Investigation Log

**Product-separated endpoint and workflow ledger:** [Microsoft 365 search findings](m365-search-findings-2026-10-10.md). It has distinct Outlook, Teams, SharePoint/OneDrive, and Microsoft Graph sections, with links between the products where they share an API or context mechanism.

**Date:** 2026-10-10
**Scope:** Search behavior and improvement opportunities for the Outlook and Teams connectors. This is a portable, anonymized evidence log. No real message content, account identifiers, tenant URLs, or captured source files are included.

## User requirements

- Investigate Microsoft 365 search across Outlook and Teams, including calendar events, meeting recaps, and transcripts.
- Prefer remote search that returns candidate items and enough metadata for an agent to retrieve context selectively, rather than downloading an entire corpus into the model.
- Let the agent decide which hits merit expansion, whether to request more pages, and how much surrounding context to retrieve. Avoid semantic “promising hit” heuristics in connector code.
- As a secondary objective, find code that can be simplified and queries that can be made faster or more reliable. Measure claims where feasible.
- Treat existing implementations as comparison points, not as design constraints.

## Executive findings

- Teams Web can navigate to an exact message without a history-page number. Microsoft documents message deep links keyed by chat/channel identity and message ID. The search result can identify the message; the client resolves that anchor and loads the UI around it.
- Microsoft Graph Teams message search is a viable remote candidate-discovery route. It returns message identifiers and excerpts, but it does not return every message property, a global result count, or a stable page number in the conversation history.
- For a first-principles topic-search workflow, remote search should return one page of hits and provenance. The agent should explicitly request a full hit, surrounding messages, thread replies, another result page, or a refined query.
- Teams transcripts and meeting recaps require separate discovery/retrieval paths. Ordinary `chatMessage` search does not establish coverage of those records.
- Current Teams `search_messages` behavior is fetch-and-archive followed by local matching over the requested chat/date scope. It does not use a verified server-side message-search route. Broad searches can spend significant time on remote pages and remain incomplete at the call’s page budget; this is not necessarily a full-history download on every request.
- Outlook already uses remote Graph search for mail. One measured call was slower cold and faster on repeated calls, but the sample is too small to generalize.
- A persistent full-message cache is not required for remote-first search. Whether to remove the Teams cache should be decided after mapping every other feature that depends on message history or roster/catalog data.

## Teams findings

### Exact-hit identity and navigation

Microsoft Search API results for `chatMessage` include a message resource with fields such as `id`, `chatId`, `createdDateTime`, sender, and a result summary; channel hits can carry `channelIdentity`. Graph also exposes a direct get-message operation using the chat/channel identity plus message ID.

Microsoft documents Teams message deep links. A chat link is addressed by chat ID and message ID. A channel link is addressed by channel ID and message ID and can include the parent message and creation time. This is the documented mechanism that explains exact navigation from a search hit without calculating a history page ordinal.

Relevant primary sources:

- [Microsoft Search API for Teams messages](https://learn.microsoft.com/en-us/graph/search-concept-chat-messages)
- [Deep links to Teams](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams)
- [Get a chatMessage](https://learn.microsoft.com/en-us/graph/api/chatmessage-get?view=graph-rest-1.0)

### History paging and context limits

Graph chat-history listing returns pages of up to 50 messages, ordered descending by `createdDateTime` or `lastModifiedDateTime`; ascending order is not supported. Paging uses an opaque `@odata.nextLink`, not a globally meaningful page number. `createdDateTime` filtering supports `lt` but not a general two-sided range; `lastModifiedDateTime` supports `gt` and `lt` when the matching order-by is used. The latter is modification-time based and should not be treated as original chronological order.

Therefore, a search hit’s message ID and timestamp are strong anchors. They do not reveal a stable “page N” in the conversation. A context tool can accept the exact hit and caller-chosen direction/count, then use message lookup plus deterministic paging or time filters. The tool should report any scan/page limit or incomplete context. Efficiency for older hits and very active chats still needs a live measurement.

Primary source: [List messages in a chat](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0)

### Search API limitations and adjacent corpora

Graph message search covers signed-in-user Teams message bodies and supported attachments. Search results omit some `chatMessage` properties, so the hit can be hydrated through the Teams API. The documented `total` is the current page count, sorting is not customizable, and `chatMessage` search cannot be combined with other entity types. Preserve `moreResultsAvailable` and page state; do not describe a result page as an exhaustive match count.

Meeting transcript listing is a separate Graph API and is scoped to an online meeting. Transcript access can depend on delegated permission and tenant policy. Meeting AI insights/recaps are a separate surface with additional licensing/permission conditions. Search across these sources needs separate, explicitly reported result sets and cursors.

- [List transcripts for an online meeting](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-transcripts?view=graph-rest-1.0)
- [Teams meeting AI insights](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/meeting-insights)

### Current connector measurements and behavior

Measurements are from one connected account and should not be treated as universal:

- One fresh, one-chat/one-day exact search fetched one remote page in about one second. Repeating it fetched the page again rather than serving a completed history from a durable search index.
- A broad all-catalog search with a five-page budget covered 966 catalog chat IDs, fetched five pages, took about 8.7 seconds, and remained incomplete with a continuation cursor.
- The local index status showed 14,723 stored messages, 966 catalog rows, and no complete message histories. A separate visible chat listing had 734 chats and 232 empty chats excluded; the scope difference needs validation before using these figures to claim missed or redundant work.
- Search matching is local exact, substring, or all-terms after Unicode NFKC normalization and case folding; accents remain significant. The current search tool requires a chat scope or all-catalog scope and an inclusive date range. Its cursor continues preparation for that same search and expires after 30 minutes or server restart.

### Synthetic FTS experiment

An experimental local FTS5 comparison used synthetic messages. At approximately 14.7k rows, a rare-substring case was about 4x faster, while common-substring/all-terms cases were about 7.4–7.9x slower; the database grew about 30.5%. At 100k rows, rare-substring search was about 9.5x faster and common queries about 38–48x slower, with database growth around 34.1%. Only three synthetic query samples were compared. This does not justify replacing the current matcher without representative workload tests.

## Outlook findings

- The connector’s mail search is server-side Graph search; it already avoids requiring a full mailbox-content download for each query.
- A measured search call took about 3.409 seconds cold and 0.760/0.784 seconds on two subsequent calls, with one hit. These are end-to-end timings for one profile and tiny sample; they do not isolate Graph latency or establish general performance.
- A search-route probe stopped before any HTTP request because the selected profile lacked a complete Graph/Outlook/Search client-and-scope configuration pair. This is configuration setup failure, not evidence that Microsoft denied the route.
- Outlook’s current architecture retains local SQLite state for other purposes, including folder/catalog freshness and shared auth/cache coordination. “Remote search” does not automatically imply “no local state.”

## Recommended investigation direction

1. Compare Graph message search with the Teams Web search/open flow using known, synthetic or otherwise explicitly authorized cases. Record query, hit identity, timestamp, source type, result cursor, and latency/counts, without retaining private message bodies or raw captures.
2. Trace the Teams Web click-through request after opening a search hit. Determine whether the route uses a direct message ID, time anchor, scroll token, page cursor, or a combination to hydrate the surrounding window. The public deep-link contract confirms message-addressed navigation, but the internal Web JavaScript and request sequence have not yet been verified.
3. Prototype a remote-search result contract that returns compact excerpts plus source IDs/provenance and search paging state. Let the agent request exact-hit hydration and caller-sized context only when it chooses.
4. Measure old-hit and high-volume-chat context retrieval, including how to move in both directions when Graph only supplies descending pages and opaque next links.
5. Measure freshness/index delay, recall/coverage across chats, channels, transcripts, and recaps, auth/consent requirements, request count, response bytes, and latency before selecting a backend.
6. Audit Teams cache consumers before removing the full-message cache. Keep this separate from the question of whether remote-first search needs that cache; it does not need a full-history cache by itself.
7. Reassess simpler/reliable queries and code only after tracing current call paths and collecting comparable measurements. Do not infer that FTS or any particular API route is faster from the small samples above.

## Evidence gaps and status

- **Confirmed from Microsoft documentation:** message identity fields, direct Teams message links, exact-message Graph lookup, Graph chat-history paging limits/order/date-filter operators, and Graph Search API response limitations.
- **Measured for one profile:** current Teams fetch/local-match behavior and the small search/FTS timing samples above; Outlook end-to-end mail-search samples above.
- **Not yet verified:** the actual Teams Web JavaScript function and network request used when clicking a hit; whether its internal API offers a direct anchor/window request or stable directional continuation; live Graph-versus-Web search recall and latency on common query shapes; transcript/recap keyword-search feasibility in this tenant.
- **Changes:** this log records research only. No connector product code was changed and no repository checks were run for this documentation-only record.
