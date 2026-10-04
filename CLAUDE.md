# Repository agent instructions

`CLAUDE.md` (Claude Code) and `AGENTS.md` (Codex and other agents) are identical. Change both together.

## The project

A local connector for one user's Exchange Online mailbox: an MCP server for agents and a small local web UI. Reads go through Microsoft Graph; writes (drafts, send, mailbox changes) go through Outlook Web (OWS).

- What it must do: `docs/outlook-requirements-v4.md`
- How it is built: `docs/architecture.md`
- What Microsoft allows (evidence): `docs/outlook-api-research.md`
- Work register and status: `docs/outlook-roadmap.md`

Layers: `surfaces/` (MCP, web) → `service/` (every domain decision) → `remote/` (Graph, OWS; the only code that knows wire formats and ids) and `store/` (SQLite). The service depends on the ports in `remote/ports.py`, never on a concrete adapter.

## Process lifetime: short-lived, local, single user

This is not a hosted or long-running server. Every entry point is a short-lived local process for one user: the MCP server lives for one agent session (stdio), the web UI until it is closed or idle (127.0.0.1 only), `auth` for one sign-in. Weeks can pass between runs, and several processes may run at once against the same local store and token cache. Design for that:

- Assume nothing in memory survives between calls of different sessions; anything a later call needs travels in the result (cursors are self-contained) or comes from the server.
- Never serve a stale local cache to the call that needs it. A cache may be weeks old when a process starts; it is used only while fresh (the folder cache: 10 minutes), otherwise the call waits for a refresh. No stale-while-revalidate, no background refresh that only helps a later process.
- No background work that outlives the call: no schedulers, sync loops, watchers or warm-up tasks.
- Shared local files (store, token cache, output folders) must tolerate concurrent processes: short transactions, cross-process locks, exclusive file creation.
- Do not add multi-user, remote-access or always-on concerns (auth for other users, network listeners beyond localhost, process supervision) unless the user asks.

## Validate once

Check an input once, at the entry point that receives it (the service method a surface calls, or the
one that reads it from the server), and let the routines it calls trust it. Do not re-check what a caller
already guarantees. Write the contract in the docstring: an entry point says what it validates
("Entry point: ..."), and an internal routine says what it assumes ("Assumes (not re-checked here): ...").
When you add a caller to a routine that assumes sanitized input, make sure the new caller provides it.

## Checks: run all four before every commit, and fix what they report

```bash
uv run ruff check .
uv run ruff format --check .   # `uv run ruff format .` to fix
uv run pyright
uv run pytest
```

Do not commit with a failing check, and never skip, disable or weaken a test to get green. Add or update tests with every behaviour change; they run against the fake mailbox in `tests/fakes/graph_fake.py` (synthetic data only, never real captures).

## Compatibility policy

- Treat this repository as an actively developed application with a fresh current format.
- Do not add backward-compatibility shims, old-format readers, automatic data migrations, schema-version frameworks, or fallback behavior for previous application versions unless the user explicitly asks for it.
- When changing a current format, update its producers, consumers, fixtures, and tests together. Do not preserve obsolete formats “just in case.”
- Keep account-ownership safeguards and current-format validation; these protect present user data and are not format-compatibility layers.

## Safety rules for mail

- Never send mail, or change the real mailbox, without the user's explicit request. Live tests use self-sends and messages the user names.
- Writes are sent once and never retried automatically; an unclear outcome is reported as unknown.
- Never log or commit mail content, addresses, tokens or real identifiers. Tokens stay in the encrypted cache.

## Docs

When behaviour changes, update the affected docs in the same change (README, architecture, requirements, roadmap). Record live findings in the roadmap and research docs with their date.
