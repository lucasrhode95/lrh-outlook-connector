# Research probes

Small, **stdlib-only** scripts that ask Microsoft what this tenant actually allows and how the
APIs actually behave. They are independent of the `outlook_connector` package on purpose:

- They must work on a new tenant, or after a Microsoft policy change, before the app does.
- They test Microsoft, not our code, so the app's tests and tooling skip this folder.

Findings go to [`docs/outlook-api-research.md`](../docs/outlook-api-research.md) (sanitized).
How to act on them is in [`docs/architecture.md` §6](../docs/architecture.md) (capability routing and portability).

## What belongs in the findings

- Link to official Microsoft documentation for public API routes, schemas, parameters,
  permissions, limits and lifecycle guidance. Do not copy that reference material here.
- Keep dated client/tenant authentication results, unexpected behavior, integration checks
  that settle a design question, measured costs, failed variants and unresolved coverage.
  A documented route returning 200 is useful when it proves our existing sign-in can use it;
  it is not a reason to reproduce the API walkthrough or every returned field.
- For private OWS/Substrate APIs, retain the discovered request and response contracts,
  authentication audience, evidence source, tested variants and failure behavior. Separate
  browser/source observations from successful standalone calls.
- State the sample and what it does not prove. An empty search is not proof that the feature
  is unsupported; a few hits are not proof of exhaustive recall. Keep implementation choices
  in the roadmap and architecture rather than expanding the research record into a plan.

## Running

From the repo root, with any Python ≥ 3.11. On Windows, set `PYTHONUTF8=1` so non-ASCII query
text prints correctly.

```bash
python research/probes/auth.py read        # sign in: Outlook Mobile -> Graph Mail.Read
python research/probes/auth.py write       # sign in: One Outlook Web -> outlook.office.com
python research/probes/auth.py search      # usually silent, via the write client's refresh token
python research/probes/auth.py --status
```

`auth.py` prints a device-code link and code. Sign in in any browser. Tokens are stored as
**plaintext** in `.local/probe-tokens.json` (git-ignored). Delete that file when you are done.
The probes refuse to use a token that belongs to a different account than the one configured
in `common.py`.

## Order for a new tenant (or a re-check)

| Step | Probe | Answers | Changes the mailbox? |
|---|---|---|---|
| 1 | `graph_scopes.py` | Which Graph mail scopes each first-party client can obtain (`.default` + explicit). This decides the Graph/OWS split. | No (token requests only) |
| 2 | `catalog.py` | Mailbox size, folder tree, metadata walk and delta timings | No |
| 3 | `conversations.py` | Conversation retrieval across folders, `$orderby` support, reply-header quality | No |
| 4 | `search.py "query" …` | Graph `$search` vs `/search/query` vs Substrate: recall, totals, paging, id alignment | No |
| 5 | `attachments.py` | Attachment types, inline/`cid:` usage, `$value` per type | No |
| 6 | `delta_moves.py --snapshot`, then a change, then `--check` | What delta reports for moves, deletes and read toggles | No (the change is made by you or by step 7) |
| 7 | `ows_mutations.py --case OPS "subject" …` | OWS write contracts: read, flag, categories, conversation read, move, soft delete | **Yes**, only with `--i-authorize-mutations`, on named Inbox messages |

Steps 2–5 need the `read` profile. Step 4 also needs `search`, and step 7 needs `write`.

## Output and safety rules

- Output is statuses, counts, timings, shapes, error codes and salted short hashes. No tokens,
  bodies, subjects, addresses or raw ids are printed. The exceptions are query strings and case
  labels the operator supplied.
- Save runs under `.local/probe-results/` (git-ignored) if you want to keep them:
  `python research/probes/catalog.py > .local/probe-results/catalog.json`.
- The client/scope pairs Microsoft denied with `AADSTS65002` are listed in `common.DENIED`
  and are never requested again. Update the list per tenant.
- Writes: no send, no hard delete, explicit subject targets only, Inbox copies only, no
  retries after an ambiguous result. A prior authorization does not carry over to new targets.

## Capturing Outlook Web (reverse engineering)

Probes test what we already know how to call. To discover *new* routes or request shapes, capture
what Outlook Web itself does while you perform one deliberate action:

1. Open a fresh browser profile/session and sign in to Outlook Web.
2. Perform **one** controlled action with a unique marker. For example, search for
   `MARKER-1234 something`, or flag a named test message.
3. Export the traffic and loaded sources with either:
   - a resource-export extension such as **Save All Resources** (Chrome), which saves the loaded
     JavaScript bundles and responses as a folder or zip, or
   - DevTools → Network → **Export HAR** (with content), which also includes request bodies.
     Response-only exports omit them; the 2026-10-02 capture lacked request bodies, so the
     search request shape had to be reconstructed from the bundles.
4. Save the export under `web-app-download/` (git-ignored) and inspect it **in place**. Search
   the JavaScript for the action or route name, and match responses by the marker.
5. Record only route names, request/response **shapes**, key names and counts in
   `docs/outlook-api-research.md`. Delete the export when done.

**Captures are credentials.** They contain live access and refresh tokens, cookies, canaries,
mailbox content and addresses. Never commit them, share them, or turn them into test fixtures.
Use them only to learn request contracts, which are then re-implemented with our own app-owned
tokens. Browser credentials must never be reused (requirements v4 §4).
