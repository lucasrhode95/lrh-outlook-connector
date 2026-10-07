# Repository agent instructions

`CLAUDE.md` (Claude Code) and `AGENTS.md` (Codex and other agents) must be identical. Change both together.

## Shared development rules

### Process and persistence

This is a local, single-user connector unless this repository documents otherwise. Keep processes scoped to their documented entry point or session. Do not add schedulers, sync loops, watchers, warm-up tasks, remote listeners, or multi-user features unless asked.

- Do not assume in-memory state survives across calls or processes. Return what a later call needs or read it from the server.
- Serve cached data only while it meets the documented freshness policy. If fresh data is required, wait for refresh instead of serving stale data.
- Shared local files (auth state, caches, indexes, output folders) must tolerate concurrent processes.

### Input boundaries and code structure

Validate untrusted input once at the entry point that receives it, and let internal routines trust the validated contract. Document entry-point validation and internal assumptions in docstrings. When adding a caller to a routine that assumes sanitized input, make sure that caller provides it.

Order modules top-down: public entry points first, then called functions, then private helpers and formatting. Keep calls followable with named functions; prefer explicit branches over string-based dispatch. Import the names used. Keep test-only seams in tests; constructor injection that is part of the design stays.

### Checks, tests, and documentation

Run every check documented for this repository before each commit and fix failures. Never skip, disable, or weaken a check to get a green result. Add or update tests for behavior changes and use synthetic fixtures, never real captures or private user content. Update affected documentation in the same change, and date new live research findings.

### Private data, external effects, and stored formats

Never log or commit credentials, tokens, session cookies, private user content, or real identifiers. Use this repository's documented secure storage for authentication state. Do not perform external writes without the user's explicit request. Make each external write once; do not retry automatically, and report uncertain outcomes.

Treat persisted data as the current format. Do not add backward-compatibility shims, old-format readers, automatic migrations, schema-version frameworks, or fallback behavior for previous application versions unless asked. Keep safeguards that protect current user data and account ownership.

## The project

A local connector for one user's Exchange Online mailbox: an MCP server for agents and a small local web UI. Reads go through Microsoft Graph; writes (drafts, send, mailbox changes) go through Outlook Web (OWS).

- What it must do: `docs/outlook-requirements-v4.md`
- How it is built: `docs/architecture.md`
- What Microsoft allows (evidence): `docs/outlook-api-research.md`
- Work register and status: `docs/outlook-roadmap.md`

Layers: `surfaces/` (MCP, web) → `service/` (every domain decision) → `remote/` (Graph, OWS; the only code that knows wire formats and ids) and `store/` (SQLite). The service depends on the ports in `remote/ports.py`, never on a concrete adapter.

## Outlook-specific process and cache rules

The MCP server lives for one agent session (stdio); the web UI stays on `127.0.0.1` until closed or idle; `auth` runs for one sign-in. Weeks can pass between runs, and several processes may use the same local store and token cache. The folder cache is fresh for 10 minutes; otherwise the call waits for refresh. Do not use stale-while-revalidate or background refresh that only helps a later process.

## Checks

Run all four before every commit:

```bash
uv run ruff check .
uv run ruff format --check .   # use `uv run ruff format .` to fix
uv run pyright
uv run pytest
```

Tests use the fake mailbox in `tests/fakes/graph_fake.py`; use synthetic data only, never real captures.

## Portability and research

Keep source, documentation, comments and fixtures free of real personal names, account
addresses and organization identifiers. Use fictional examples. Deployment-specific URLs,
clients, scopes and optional expected-account checks belong in named configuration settings,
not scattered literals. Microsoft public endpoints and first-party client IDs may be defaults.

The research folder is a portable investigation toolkit and an anonymized, dated evidence
record. Its standalone helpers must work before the app does. Never treat one environment's
grants or denials as universal restrictions; allow users to configure and reassess them.
Preserve same-account safeguards while deriving routing and self-send identity from sign-in.
Record any genuinely organization-exclusive endpoint or account dependency for the user's
decision before changing it. Run the research offline checks when changing probe behavior.

## Outlook safety and documentation

- Never send mail or change the real mailbox without the user's explicit request. Live tests use self-sends and messages the user names.
- Never log or commit mail content, addresses, tokens, or real identifiers. Tokens stay in the encrypted cache.
- When behavior changes, update the affected README, architecture, requirements, and roadmap docs. Record live findings in the roadmap and research docs with their date.
