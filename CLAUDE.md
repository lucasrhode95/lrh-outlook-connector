# Repository agent instructions

`CLAUDE.md` (Claude Code) and `AGENTS.md` (Codex and other agents) are identical. Change both together.

## The project

A local connector for one user's Exchange Online mailbox: an MCP server for agents and a small local web UI. Reads go through Microsoft Graph; writes (drafts, send, mailbox changes) go through Outlook Web (OWS).

- What it must do: `docs/outlook-requirements-v4.md`
- How it is built: `docs/architecture.md`
- What Microsoft allows (evidence): `docs/outlook-api-research.md`
- Work register and status: `docs/outlook-roadmap.md`

Layers: `surfaces/` (MCP, web) → `service/` (every domain decision) → `remote/` (Graph, OWS; the only code that knows wire formats and ids) and `store/` (SQLite). The service depends on the ports in `remote/ports.py`, never on a concrete adapter.

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
