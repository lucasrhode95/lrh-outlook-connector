# Microsoft 365 search investigation

Last updated: **2026-10-11**

## Objective

Determine which Microsoft 365 searches can run online, their scope and limits, and how the
capabilities fit the two separate connectors: `lrh-outlook-connector` and
`lrh-teams-connector`. This is an investigation and handoff; product API changes and any
repository merge remain separate decisions.

A secondary objective is to review the existing search, list, and get paths in both connectors for
evidence-backed opportunities to reduce code, redundant requests, over-fetch, latency, and paging or
retry fragility. Keep this review separate from Graph migration: recommend reuse only where
authentication, permissions, semantics, and coverage are demonstrated. Research findings do not
authorize product-code changes.

Required search scope:

All source groups below are in scope for the requested Microsoft 365 search capability. If a
source cannot be reached through supported connector authentication, record it as a capability
gap; do not silently treat it as optional.

- Search company SharePoint, the signed-in user's OneDrive, and accessible shared files.
- Search email and Teams files/attachments while retaining parent-message context.
- Search calendar and Teams meetings, details, participants, and attachments.
- Discover meetings with transcripts, retrieve accessible transcript text, and test whether
  transcript text can be searched online across meetings.
- Reassess Teams message/catalog behavior and identify Graph migrations only where usable
  authentication and behavioral parity are shown.

## Evidence labels

- **DOCUMENTED** — Microsoft or another primary API contract describes the capability.
- **CAPTURED** — a browser HAR contains a request and response associated with an observed UI
  flow. A browser success does not prove the route is usable by a standalone connector.
- **SOURCE** — downloaded client source contains a builder or parser. Source alone does not
  prove the associated operation ran in the capture or can run outside the browser.
- **STANDALONE** — a bounded read-only request succeeded through the connector's app-owned
  authentication flow.
- **USER-PROVIDED** — a prior UI walkthrough or named resource identifies a useful test target;
  this is a lead, not standalone endpoint or permission proof.

For each experiment, record: question, evidence, bounded test, observed result, interpretation,
remaining uncertainty, and next action. Keep account-specific grants and denials separate from
general API support.

## Safety and bounds

- Keep HARs, resource archives, downloaded source, and all temporary scanners outside both
  repositories. Do not execute captured JavaScript or commit raw/anonymized captures, indexes,
  manifests, or capture-reading machinery.
- Use only app-owned authentication and existing encrypted token/cache boundaries. Never reuse
  captured cookies, tokens, canaries, or signed URLs.
- Keep requests read-only. Create test messages, files, meetings or transcripts only when the user
  explicitly authorizes the destination and write. Use generic synthetic values, never retain their
  marker strings in the evidence record, and wait for OneDrive upload before searching a synced
  test file. Do not change permissions or mark items read.
- Default to at most 25 results per page, three pages per provider/query, and five hydrated hits
  per comparison. Record any justified bound expansion before running it.
- Do not retain query values, result names, message/file contents, real identifiers, or personal
  account/organization details. Report only sanitized route shapes, field names, counts, timing,
  status, and supported conclusions.
- A missing result is not proof of unsupported behavior. A transcript retrieval route is not proof
  of transcript search. A parent-message match is not proof that a particular attachment matched.

## Work packages

### A. SharePoint and OneDrive files (Outlook X11)

Compare company-wide accessible, personal, shared/library, and site/folder scopes. Map search
request fields, refiners, ranking, snippets, result references, paging, partial-result signals,
and the path from a result to Graph metadata/download. Prove title/name and content-only matches
with suitable existing content plus a negative control. Do not equate SharePoint item ids with
Graph drive/item ids without a demonstrated mapping.

### B. Email and Teams attachments

For each source, determine whether a search hit is a file, parent message, or both; whether it
identifies the matching attachment; and whether a selected hit can be hydrated without downloading
every candidate. Keep cloud files, embedded files, reference attachments, and inline/hosted images
distinct. Compare captured Files filters with Graph message and drive-item search.

### C. Calendar events and meetings (Outlook X10 and Teams)

Trace the Teams Meetings filter, recap search, calendar event, and expanded details separately.
Establish which results are calendar events, occurrences, online meetings, chats, or recap items.
Record how agenda, invited attendees, actual attendance, join links, and attached files are obtained.
Build only evidence-backed links among event, meeting occurrence, online meeting, chat, recap, and
transcript; leave missing links unresolved. Recheck calendar access through the existing app-owned
profile before considering any new consent or sign-in route.

**USER-PROVIDED targets:** an earlier walkthrough identified a calendar event whose expanded view
showed an agenda, participants, and an attachment, plus two transcript examples reached through
different UI paths. Use those existing examples to test the event-to-meeting/chat/recap linkage;
keep their names out of portable research notes. Invited attendees and actual attendance remain
separate fields.

### D. Meeting transcripts

Evaluate three separate capabilities: discovery of meetings with transcripts, retrieval/paging of
an accessible transcript, and online keyword search across transcript text. Map continuation,
segment ordering, timestamps/speakers, and view-versus-download permissions from source and capture.
Compare per-meeting Graph transcript retrieval with any captured Teams routes. Treat bulk
organizer-only and per-meeting delegated routes as different capabilities.

**USER-PROVIDED permission/paging lead:** one recap-app transcript displayed additional content as
the user scrolled, while its separate download action asked the user to request permission from
the meeting host. A second transcript was accessed from the meeting chat. Investigate the read/view
path, continuation, and download permission separately; do not use transcript-content requests to
bypass the observed download restriction.

### E. Teams search and Graph migration

Trace global search suggestions, submitted searches, filters, result hydration, and transcript
recap flows. Compare the Teams Files/Meetings/Message categories with Graph `driveItem`, `message`,
`chatMessage`, and `event` operations under the existing app-owned authentication. Keep broad online
discovery distinct from the durable local message archive and scoped content search. Produce an
operation-by-operation recommendation: ready, conditional, no current benefit, or blocked.

### F. Connector code reduction and query quality (both repos)

Trace the in-scope public search/list/get tools through service functions to their remote adapters in
`lrh-outlook-connector` and `lrh-teams-connector`. For each path, note request count, repeated
hydration, over-fetch, paging and retry behavior, cache use, and fields already available in earlier
responses. Compare calls only where endpoint behavior and authentication are evidenced. Identify
safe opportunities to remove redundant work, reduce response volume, improve continuation handling,
or use fewer and more reliable requests. Keep Graph migration as a separate decision from code
simplification.

For each candidate, record the current call path, relevant files, supporting evidence, expected
benefit, compatibility or permission risks, confidence, and a synthetic or sanitized way to measure
latency and request-count changes. Classify each as demonstrated, promising but needs measurement,
or not recommended. Do not implement a code change during this research phase.

## Checkpoint (2026-10-10)

- **DOCUMENTED:** Outlook research §3.8 reviews Graph Search for files, messages, Teams messages,
  calendar events, and transcript retrieval, including permission and result-limit boundaries.
- **CAPTURED + SOURCE:** the first sanitized HAR pass and targeted Teams client-source trace are
  documented in Outlook research §3.9 and Teams API research. They distinguish suggestions from
  submitted search, SharePoint from Substrate search, a separate Meet-app Event builder, recap
  `MeetingInsights`, and several transcript-retrieval families.
- **CAPTURED paging:** a second sanitized pass found SharePoint and Teams file responses with
  `MoreResultsAvailable=true`, plus a 25-hit Teams Event answer response. Every explicit search
  offset remained zero, so no advancing page request has been observed. The transcript capture
  likewise does not prove paging. Details and response caveats are in both dated research records.
- **SOURCE: Outlook Web:** a read-only audit of the user-provided source ZIP mapped the Search v2
  body wrapper and separate suggestion/init builders. The bundle was read in place; no extracted
  source or index was retained. Its page-state handling does not establish an advancing request.
- **STANDALONE:** earlier Graph file search is recorded in Outlook research §3.7. A bounded
  Graph message-search request through the encrypted app-owned token cache returned HTTP 200 but
  zero hits; this confirms route acceptance and response shape only. Outlook research §3.7 also
  records earlier mail keyword, subject and body searches with hits, but omits their query values
  and result details; those results do not settle controlled content-only matching or attachment
  attribution. Teams roadmap records that Graph `driveItem` search worked in the tested profile
  while Graph `chatMessage` search was denied. These outcomes are account-specific, not universal
  restrictions.
- **STANDALONE, controlled mail:** on 2026-10-10, separate subject-only and body-only terms for
  one authorized synthetic message each returned HTTP 200 and two hits from Graph `/search/query`
  (`message`, `from=0`, `size=5`, `total=2`, `moreResultsAvailable=false). The Outlook product
  search returned one grouped conversation per term at `limit=1`, with incomplete coverage. No
  content, query values or identifiers were retained. This validates subject/body matching for
  the tested mailbox only.
- **STANDALONE, Teams local search:** a bounded four-minute self-chat scope fetched one message;
  substring search returned one match with complete requested-scope coverage. The connector
  reports `server_side_search_used=false`; this validates local archive search after retrieval,
  not online Teams search. See the dated Teams API research note.
- **USER-PROVIDED meeting targets + STANDALONE chat-history check:** the earlier walkthrough named
  two transcript-bearing meeting chats and one calendar event with agenda, participants, and an
  attachment. The current Teams catalog matched both named chats; bounded `get_messages` calls
  returned complete remote histories in one page each (18 and 9 items). Seven items in each chat
  contained a transcript/recap-related text signal. Names, message bodies, and identifiers were
  not retained. Envelope metadata in each history included `RichText/Media_CallRecording` and
  `RichText/Media_CallTranscript`; five message items contained an `amsTranscript` XML reference
  and one contained an ODSP transcript-path signal. Counts may overlap and are not distinct
  transcript objects. No URI was followed. This is message-history reference evidence only; it
  does not establish transcript artifact access or online transcript keyword search.
- **CONFIRMED, narrow file known hits:** after the user confirmed upload of one authorized generic
  file to the shared folder, Graph `driveItem` searches returned HTTP 200 with one filename-only
  hit and one content-only hit. No result was hydrated. The effective permission basis was not
  isolated, so this does not establish a generally sufficient permission set or broad recall; see
  Outlook research §3.12.
- **Required calendar/meeting sources, still open:** Graph `event` Search is a required path, but
  the tested Outlook token was denied because it lacks `Calendars.Read`; the documented search is
  limited to the signed-in user's primary calendar. The configured Graph profile requests only
  `Mail.Read`, and `auth graph --force` repeats that configured scope list. Delegated
  `Calendars.Read` is user-consentable in Microsoft's permission table, but preauthorization for
  this Microsoft first-party client/scope pair is untested. A future bounded check would add that
  scope, force sign-in, inspect the granted scope claim, and repeat event search. Delegated
  transcript access separately requires admin-consented `OnlineMeetingTranscript.Read.All` and
  the tenant transcript API setting. Teams recap `MeetingInsights` and transcript retrieval remain
  separate captured/source-mapped paths without standalone authentication proof.
- **Not established yet:** broad file recall across sources and folders, minimum permissions for
  shared-file search, exact attachment attribution, complete result linkage/hydration, fresh-process
  private-search reproduction, exact global Files source defaults, event-to-online-meeting/chat/
  recap identity linkage, recap continuation, transcript discovery and paging through supported
  authentication, or online full-text transcript search. Graph transcript content retrieval and
  transcript text search are separate acceptance checks.

### Initial connector code-path audit (2026-10-10)

This section combines source inspection with bounded synthetic/live timing checks. It does not
establish a product-level latency improvement.

- **Outlook — keep the count lookup under the current contract:** `MailboxService.search` calls
  `conversation_sizes` for the unique conversations on a result page after grouping hits. The
  default search response exposes `message_count` and `message_count_at_least`, and failures are
  intentionally ignored. `GraphMailReader.conversation_folders` coalesces those lookups into one
  Graph `$batch` HTTP request with one bounded message-listing subrequest per conversation, then
  the service applies folder reachability, exclusions and duplicate-copy rules. Removing this path
  would remove a documented response field; no current consumer-opt-out exists. The count batch's
  isolated latency was not measured.
- **Outlook — keep unless stable-ID requirements change:** `GraphMail.search` translates returned
  IDs in one bulk `translateExchangeIds` call per page because Graph search ignores the immutable-ID
  preference. This supports stable IDs across moves. Removing it could break identity semantics.
- **Outlook — current search timing (2026-10-10):** three sequential one-hit calls to the signed-in
  `search_messages` tool returned one conversation and one message each, with `message_count`
  present on all three. Wall time was 3,409 ms on the first call, then 760 ms and 784 ms; coverage
  was incomplete at the one-hit limit. These timings include the connector/tool round trip and
  warm-up effects; they do not isolate Graph or the count batch. A direct Graph bake-off rerun was
  blocked before any HTTP request because the research probe requires configured graph, outlook,
  and search client/scope pairs, which this research profile does not define. This is configuration,
  not an endpoint denial.
- **Outlook — no demonstrated faster replacement route:** direct Graph `/search/query` has prior
  controlled subject/body known-hit evidence, while the product search wrapper has incomplete
  coverage. The timed calls did not compare identical queries across both routes. Do not replace the
  current route until an app-owned, same-query parity/timing comparison is available. Existing
  batched counts, immutable-ID translation, paging and transient retry behavior support the current
  contract and stable result identity.
- **Teams — promising, needs benchmark and throttle checks:** all-chat search prepares remote
  history sequentially under a shared page budget before local matching. Bounded concurrency might
  lower latency, but it can change throttling, coverage accounting and page-budget behavior.
- **Teams — promising, needs synthetic benchmark:** SQLite message search uses leading-wildcard
  `LIKE` matching over normalized stored text within the selected scope/date range. At larger
  archives this may scan many rows. Consider an FTS index only after benchmarking synthetic data
  and checking equivalence for normalization, literal matching, accents, substring semantics,
  ordering and keyset pagination.
- **Teams — retain current continuation safeguards:** remote history/catalog code validates
  continuation and detects repeated links. No evidence supports simplifying those checks; they
  protect completeness and prevent loops.

**Teams measured results:** on synthetic archives of 14,720 and 100,000 messages across 32 chats,
SQLite FTS5 trigram search preserved the first-page IDs and continuation flag for the three tested
queries. A selective rare substring improved from 15.36 ms to 3.82 ms at 14,720 rows and from
44.06 ms to 4.64 ms at 100,000 rows. Common substring search regressed from 3.97 ms to 29.28 ms
and from 4.02 ms to 153.09 ms; common all-terms search regressed from 4.01 ms to 31.80 ms and
from 3.87 ms to 185.98 ms. The FTS index added 30.5% and 34.1% to database size and took 77 ms
and 717 ms to build. Three representative queries are not enough to establish equivalence for
short terms, punctuation, Unicode edge cases, paging, edits/deletes or FTS write maintenance; do
not replace `LIKE` generally.

On four bounded live Teams history reads (one remote page per chat, page size one), sequential
calls took 4,298 ms total. Four concurrent calls took 3,552 ms wall time, about 17% lower, with
four pages and no retries/errors; individual calls queued up to 3,552 ms and each response remained
incomplete because only one page was requested. This is modest evidence for bounded overlap, not a
full all-chat search benchmark; keep sequential preparation unless a broader probe shows worthwhile
gains without changing the global page budget or coverage semantics.

The scope-table microbenchmark measured removing the empty `DELETE FROM scope` at roughly 1–3
microseconds per search for scopes of 1–100 chats, within noise. `_tx` opens a fresh SQLite
connection for every operation, so the temp table is new and empty on each search; the delete is
redundant and safe to remove when product-code changes are authorized, but its speed impact is
negligible. No product code has changed. Remaining work: test FTS edge semantics and write cost,
measure end-to-end multi-chat preparation over larger bounded scopes, and rerun Outlook's two Graph
search routes with matching controlled terms through configured app-owned authentication.

## Required handoff

Maintain dated evidence in the research files and concise status/next steps in the roadmaps of
both repositories. Deliver a sanitized capability matrix, route/field contracts, bounded probe
results, Graph migration recommendations for Teams, and explicit unresolved evidence gaps. Do not
mark any search product feature implemented until its retrieval, limits, permissions, and coverage
are demonstrated through supported connector authentication.

Current evidence:

- Outlook: [API research §§3.8–3.11](outlook-api-research.md#38-microsoft-365-search-api-coverage-review-2026-10-10)
  and [roadmap X10/X11](outlook-roadmap.md).
- Teams: `lrh-teams-connector/docs/teams-api-research.md` and
  `lrh-teams-connector/docs/teams-connector-roadmap.md`.

## Evidence update (2026-10-11 capture/source reconciliation)

The October 10 browser HARs and source bundles were re-audited against both products' research
records. The audit corrected route attribution by surface: OneDrive Files submitted Substrate
`/search/api/v2/query` plus a SharePoint REST search proxy; company SharePoint submitted
`/searchservice/api/v2/query`; Outlook and Teams global search also used `/searchservice/api/v2/query`;
Teams Transcript queries used a separate `/search/api/v2/query` route. The common Graph Search API
`POST /v1.0/search/query` remains a distinct connector-facing route family, with `entityTypes`
choosing message/event/chatMessage/driveItem search and different permission/result contracts.

The capture review added SharePoint's image-only `GET /api/v1/documents/searchpreview`, OneDrive's
list-stream calls, SharePoint suggestion/init/configuration calls, and Teams recap/resource
hydration routes to the dated product research. It also corrected the People/profile evidence:
Teams Web did call `fetch`, `fetchShortProfile`, and `fetchfeaturesettings`, returning profile
objects or per-entry status envelopes; free-form People Picker and working-set routes remain
source-only. Teams bundle source now explains exact-hit navigation by message identity and
timestamp, but no message-history request or click-triggered HTTP sequence was present in the HAR.
One parseable mixed Search v2 response also included Message `Source` fields for
`InternetMessageId`, conversation/client-thread identities, and sent/received/created timestamps.
The Teams Web normalizer uses `InternetMessageId` for the UI result id and composes its exact-hit
navigation object from reply-chain identity, `ClientThreadId`, and `DateTimeSent`. That identifies
and dates the hit without assigning a history-page number. The normalizer also uses `ThreadSubType`
and `ParentMessageId` for reply-chain derivation; they were not top-level captured Source fields,
so their extraction is unresolved. The identifier is not yet cross-checked against Graph/Chat
Service history IDs, and the history-fetch wire request is absent. Mapping
`dateTimeSent` to Chat Service `startTime` remains inferential.

The route inventory omits static assets and telemetry as non-result endpoints. It records
bootstrap/configuration and image-preview paths when they explain a search surface or result
hydration. No captured browser route is being promoted to a portable connector contract. No live
probe or product-code change was made during this reconciliation.

Remaining open checks include standalone authentication and permissions, true continuation for
each browser provider, mapping SharePoint Search identities to Graph drive/item IDs, checking the
Teams Search `InternetMessageId`/thread identities against Graph or Chat Service history, the
post-click Teams history wire request, and the successful in-pane transcript-search request. See
the dated endpoint/click-through ledger at [M365 search findings](m365-search-findings-2026-10-10.md)
and the canonical [Outlook](outlook-api-research.md) and
[Teams](../../lrh-teams-connector/docs/teams-api-research.md) research files.

## Evidence update (2026-10-10)

The Teams Meeting Recaps web app opened both user-provided known recaps with loaded Transcript tabs. Its in-page search returned matches for one target; the successful search request route is not mapped. The global Teams Search page exposed visible filters All, Messages, Files, People, Channels, Group chats and Meetings; Meetings showed 25 rows for a common query. No Transcript filter was visible, while a two-item overflow menu could not be enumerated. The app disabled transcript download for one target, and no separate transcript file was saved in the research workspace. Separately, captured global Search v2 Transcript calls returned HTTP 200, with one structured zero-result response and two uninterpretable bodies. Meeting-list, per-meeting resource, recap-content and ODSP transcript retrieval shapes are recorded in `lrh-teams-connector/docs/teams-api-research.md`.

A targeted scan of the supplied OWA source archive found an internal Substrate Search scope-catalog entry and a Teams MCPS service-base configuration, but no concrete search, recap or transcript route builders. Static strings do not establish that a signed-in token requested or received that scope. Keep app-owned authentication and the in-page search request path as open investigation items.

## Evidence update (2026-10-10)

A bounded read-only revisit of the two user-provided transcript-bearing meeting chats found five
`amsTranscript` marker occurrences in each complete message history. Each history also contained
one embedded `URIObject` with the sanitized route shape
`/drives/{drive}/items/{item}/{segment}/{segment}/media/transcripts/{transcript}/content`;
the observed `URIObject` attribute names were `format_version`, `type`, `uri`, `url_thumbnail`,
and `version`. This is a message-embedded reference, not an observed outgoing request. The two
intervening path segments are not mapped. No URI value was retained or followed, so this provides
no evidence about HTTP method, authentication, response shape, or access beyond the previously
observed recap UI. Compare it with the Teams client builder path and a future sanitized browser
network capture; do not infer that the reference can be fetched by the standalone connector.
