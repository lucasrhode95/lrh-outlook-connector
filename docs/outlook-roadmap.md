# Outlook Connector Roadmap

Open decisions, implementation, fixes and validation only. Completed behavior belongs in the
requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-13**.

## Decisions

- **One message-listing tool.** Merging `list_messages` and `search_messages` needs one response shape. A list returns flat messages in date order; search returns conversation groups in rank order. Options: one tool whose result carries a `mode` field, or one flat shape for both (search would lose conversation grouping and `message_count`). Routing is simple: an empty query lists, and a query searches. Folder, date and scope apply to both. Renaming the tools would break existing agent configurations.
- **`search_calendar_events` date bounds.** Date terms in the query returned no result set ([research §3.20](outlook-api-research.md#320-review-gap-live-checks-2026-10-13)), so `since` and `until` only filter the pages fetched. Keep the bounds with the documented limit, or remove them so window questions go to `list_calendar_events`.
- **Conversation read budget.** `get_conversation` reads at most 1,000 messages of a conversation, and `notes` reports the cap. Graph pages conversation listings ([research §3.22](outlook-api-research.md#322-conversation-listing-paging-check-2026-10-13)), so the cap is the connector's own budget. Keep it, or raise it?
- **Default conversation scope (O4).** `get_conversation` leaves out Deleted Items and Junk unless `scope.deleted_items=true`. Confirm or change.

## Enhancements

- **Passive HTML is a deny-list.** It relies on Python's HTML parser, so markup that tokenizes differently in a browser could pass. The one save test ([research §3.21](outlook-api-research.md#321-unfinished-markup-on-save-2026-10-13)) showed Outlook dropping an event handler, but that is one account and one client. Candidates: an allow-list, or wider sanitization evidence checked against browser parsing.
- **Event to Teams meeting chat.** Establish a reference through a supported Teams-side lookup before offering an automatic handoff. Search v2 event fields expose no Teams chat id ([research §3.17](outlook-api-research.md#317-live-outlook-connector-mcp-calendar-check-2026-10-12)). Transcript and recap discovery stay in the [Teams roadmap](https://github.com/lucasrhode95/lrh-teams-connector/blob/main/docs/teams-connector-roadmap.md).
- **Calendar profile.** Calendar windows reuse the shared `teams.graph` profile, which carries the calendar scope. Decide whether a purpose-named calendar profile in `lrh-m365-auth` should replace it, or whether `outlook.graph` should request `Calendars.Read` ([research §3.13](outlook-api-research.md#313-calendar-events-meeting-recaps-and-transcripts-are-required-search-sources-2026-10-10), [§3.18](outlook-api-research.md#318-calendar-range-listing-and-description-search-through-the-teams-client-graph-profile-2026-10-12)).
- **Graph-first event hydration.** `get_calendar_event` reads through OWS only. Graph hydration needs its own check.
- **Graph event search.** It returned 12 hits for a known description term ([research §3.18](outlook-api-research.md#318-calendar-range-listing-and-description-search-through-the-teams-client-graph-profile-2026-10-12)), but its date and body-only semantics are unverified, so Search v2 stays the topic search.
- **Other calendars.** Only the default calendar is listed; shared calendars and other people's calendars are not.
- **Event attachment download.** Event attachments are listed but not downloadable.
- **Branch-aware conversations (parked, O5).** Build the reply tree from RFC 5322 `Message-ID` / `In-Reply-To` / `References` headers, with Exchange `ConversationId` and `Thread-Index` as supporting evidence. Detect branches, including forwards, and offer a merged chronological view where each message is labeled with its branch. Research the quality of the existing headers first (R4). Until then a conversation is exactly Exchange's conversation.

## Validation

- **Fresh-account consent.** The `m365.search` and `teams.graph` calendar grants were verified on one account only. Check them on a new account before relying on them.
- **Mail search route.** Graph Search API has not been compared with the current Graph `$search` route on identical queries ([research §3.16](outlook-api-research.md#316-outlook-mail-search-timing-and-route-comparison-2026-10-10)). Compare terms, limits, hydration, request counts and latency before considering a change.
