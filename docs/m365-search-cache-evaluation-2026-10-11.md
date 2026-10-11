# Teams remote search and cache evaluation (2026-10-11)

This dated addendum records live connector behavior and public Microsoft Graph documentation reviewed during the Microsoft 365 search investigation. Live measurements describe one connected account and are not service-wide performance guarantees. No message content, identifiers, or account names are retained.

## Current Teams connector behavior

The Teams connector fetches pages for the requested chat/date scope, archives those messages in SQLite, then matches locally. Search reports `server_side_search_used=false`. A page budget is per call; the cursor resumes the same search's preparation and then local match paging. A fresh query creates new preparation work rather than reusing the old cursor's completed preparation.

Live measurement: the same exact one-chat, one-day search fetched one remote page on each of two independent runs (about one second per run). An all-catalog search with a five-page cap requested 966 chat IDs, fetched five pages in 8.7 seconds, remained incomplete, and returned a cursor. It did not download the entire message database in that call, but a fresh global search can repeat remote page reads across the requested scope over multiple cursor calls.

Index status reported 14,723 indexed messages and no complete message histories. It reported 966 catalog rows; a refreshed visible chat list reported 734 chats and 232 empty chats excluded. The search scope therefore appears to include entries excluded from the visible non-empty chat list. Confirm whether visiting those empty entries causes avoidable remote calls before changing scope behavior.

## Remote search candidate: Microsoft Search API

Microsoft documents `POST /v1.0/search/query` with `entityTypes: ["chatMessage"]` as server-side Teams message search. It searches message bodies and supported attachment content in the signed-in user's Teams message storage, returns ranked/date-ordered hits, and can be paged with `from` and `size`. The documented delegated permission for chat messages is `Chat.Read`; the entity-type permission table also lists `ChannelMessage.Read.All` for channel messages.

This is a strong replacement candidate for per-chat message fetching plus local text matching. Search results expose message and chat IDs and a summary, but do not include every `chatMessage` property, so hydrate returned hits selectively through the Teams message API. The documented limitations include no custom sort, no exact total beyond the current page for Teams hits, and no combining `chatMessage` with other entity types. Validate supported query semantics and result coverage against the connector's current exact/substring/all-terms behavior before switching.

## Transcripts and meeting recaps

The message-search entity does not cover transcript entities. Graph has separate transcript APIs: list transcripts for an online meeting and retrieve transcript content. The delegated permission is `OnlineMeetingTranscript.Read.All`; tenant policy can disable Graph transcript access. The documented per-meeting list applies to scheduled online meetings and excludes live events, so it is not a global full-text transcript-search endpoint.

For Copilot meeting recaps, the Meeting AI Insights API returns structured meeting notes, action items, and participant mentions. It requires `OnlineMeetingAiInsight.Read.All` and a Microsoft 365 Copilot license, is available only after a meeting, may take up to four hours to appear, and has meeting-type/retention limitations. It retrieves insight objects rather than searching their text remotely.

## Cache decision

Prefer remote message search and then remove the durable message-history cache once a live Graph request succeeds and parity, paging, hydration, and query semantics are measured. Treat transcript and recap search separately: without a server-side search route, matching their content still requires downloading content per query or maintaining a local index.

Removing the entire local store also affects participant search, which currently uses locally indexed rosters, and persisted catalog/index coverage. Replace those dependencies with verified remote routes or intentionally remove those features. The Outlook connector is a precedent for avoiding a durable mail-content cache, while still retaining minimal local application state.

## Sources

- Microsoft Search API for Teams messages: https://learn.microsoft.com/en-us/graph/search-concept-chat-messages
- Microsoft Search API entity types and permissions: https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0
- List Teams meeting transcripts: https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-transcripts?view=graph-rest-1.0
- Teams Meeting AI Insights API: https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/meeting-insights
