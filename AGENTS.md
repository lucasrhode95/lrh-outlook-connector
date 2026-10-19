# Repository agent instructions

`CLAUDE.md` (Claude Code) and `AGENTS.md` (Codex and other agents) must be identical. Change both together.

The first part, **Shared rules**, is the same in every `lrh-*` repository: change it in all of them at
once. The second part holds this repository's own rules.

# Shared rules

## Process and persistence

This is a local, single-user tool unless this repository documents otherwise. Keep processes scoped to
their documented entry point or session. Do not add schedulers, sync loops, watchers, warm-up tasks,
remote listeners, or multi-user features unless asked.

- Do not assume in-memory state survives across calls or processes. Return what a later call needs or
  read it from the server.
- Serve cached data only while it meets the documented freshness policy. If fresh data is required, wait
  for refresh instead of serving stale data.
- Shared local files (auth state, caches, indexes, output folders) must tolerate concurrent processes.

## Input boundaries and code structure

Validate untrusted input once at the entry point that receives it, and let internal routines trust the
validated contract. Document entry-point validation and internal assumptions in docstrings. When adding a
caller to a routine that assumes sanitized input, make sure that caller provides it.

Order modules top-down: public entry points first, then called functions, then private helpers and
formatting. Keep calls followable with named functions; prefer explicit branches over string-based
dispatch. Import the names used. Keep test-only seams in tests; constructor injection that is part of the
design stays. Modules that do the same job in several repositories (logging, local files, domain errors)
keep the same names and behavior; change them in every repository together.

## Checks, tests, and documentation

Run every check documented for this repository before each commit and fix failures. Never skip, disable,
or weaken a check to get a green result. Add or update tests for behavior changes and use synthetic
fixtures, never real captures or private user content. Update affected documentation in the same change,
and date new live research findings.

Documentation keeps each file to its own scope: requirements (what it must do, decisions, ownership) and
architecture (how it is built) are numbered; the research record holds only current findings that
Microsoft does not document or documents poorly; the roadmap holds open work only, split into
implementation, decisions and live verification. Keep documentation direct and summarized. Text files use
LF line endings (`.gitattributes`).

## Private data, external effects, and stored formats

Never log or commit credentials, tokens, session cookies, private user content, or real identifiers. Use
this repository's documented secure storage for authentication state. Logs carry operational timings,
statuses and counts only. Do not perform external writes without the user's explicit request. Retry a
write automatically, a bounded number of times and honoring `Retry-After`, only when the failure proves it
was not applied: the request was never sent, or the service answered 429 or 503. Never resend a write whose
outcome is uncertain (sent without a complete answer, or answered 504); report it so it can be checked.
When a multi-step write fails, undo the steps it already completed where that is safe, and report what
could not be undone. Resuming a partial upload from the ranges the service reports missing is allowed.

Treat persisted data as the current format. Do not add backward-compatibility shims, old-format readers,
automatic migrations, schema-version frameworks, or fallback behavior for previous application versions
unless asked. Keep safeguards that protect current user data and account ownership.

## Research scope

Research probes make explicitly authorized requests to configured endpoints and report responses. They do
not consume browser captures. HARs, Resource Saver archives and downloaded source bundles are temporary
inputs for ad hoc exploratory work outside the repository. Document useful URLs, request/response shapes,
dated findings and limitations. Do not retain raw or anonymized captures, per-file/member manifests, source
indexes, or capture-reading, regeneration and redaction machinery in the repository. Probes must work
without captures.

## Portability

Keep source, documentation, comments and fixtures free of real personal names, account addresses,
customer/project names and organization identifiers. Use fictional examples and synthetic content.
Deployment-specific URLs, clients, scopes and any expected-account checks belong in named configuration
settings. Microsoft public endpoints and first-party client IDs may be baseline settings; keep regional
service discovery and account-ownership checks.

The research record is an anonymized, dated starting point for forks and future policy changes. State
which client, resource, operation and evidence support a finding, without identifying the account or
organization. One environment's grants, denials or empty responses are not universal restrictions;
reassess capabilities before choosing a backend for a deployment. Flag genuinely organization-exclusive
endpoints or fixed-account dependencies for the user's decision before changing them. Standalone research
scripts require an explicit private config file rather than embedded fallback settings.

# Outlook connector rules

## The project

A local connector for one user's Exchange Online mailbox: an MCP server for agents and a small local web UI. Core mail-data reads go through Microsoft Graph. Outlook Web Service (OWS) handles mail changes and inbox-rule access; Outlook Cloud Settings handles native signature settings.

- Documentation index: `docs/README.md`
- What it must do: `docs/requirements.md`
- How it is built: `docs/architecture.md`
- What Microsoft allows (evidence): `docs/api-research.md`
- Work register and status: `docs/roadmap.md`

Layers: `surfaces/` (MCP, web) → `service/` (every domain decision) → `remote/` (Graph, OWS; the only code that knows wire formats and ids) and `store/` (SQLite). The service depends on the ports in `remote/ports.py`, never on a concrete adapter.

## Process and cache

The MCP server lives for one agent session (stdio); the web UI stays on `127.0.0.1` until closed or idle; `auth` runs for one sign-in. Weeks can pass between runs, and several processes may use the same local store and token cache. The folder cache is fresh for 10 minutes; otherwise the call waits for refresh. Do not use stale-while-revalidate or background refresh that only helps a later process.

## Checks

Run all five before every commit:

```bash
uv run ruff check .
uv run ruff format --check .   # use `uv run ruff format .` to fix
uv run pyright
uv run pytest
node --test tests/surfaces/message_reader.test.cjs
```

Tests use the fake mailbox in `tests/fakes/graph_fake.py`; use synthetic data only, never real captures.

## Research toolkit

The `research/` folder is a portable investigation toolkit. Its standalone helpers must work before the app
does. Preserve same-account safeguards while deriving routing and self-send identity from sign-in. Run the
research offline checks when changing probe behavior.

## Outlook safety and documentation

- Never send mail or change the real mailbox without the user's explicit request. Live tests use self-sends and messages the user names.
- Never log or commit mail content, addresses, tokens, or real identifiers. Tokens stay in the encrypted cache.
- When behavior changes, update the affected README, architecture, requirements, and roadmap docs. Record live findings in the roadmap and research docs with their date.
