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

## Configure the environment

The configuration file is required; the code supplies no endpoint, client, scope or
account defaults. Copy research/probe-config.example.json to .local/probe-config.json
and configure your environment before running a probe. Keep every listed setting;
missing or unknown settings cause an error before authentication or network access.
The example targets Microsoft's public cloud and Microsoft-owned clients as starting
points, not promised grants:

- tenant: organizations, a tenant ID, or a verified tenant domain.
- expected_user: optional sign-in email check; empty means no fixed email restriction.
- profiles: the client ID and resource scope for each of graph, outlook, and search.
  Registered public clients can be substituted where your administrators permit them.
- login_url, graph_resource, graph_url, ows_url, substrate_urls: endpoint and resource
  settings. URLs must use HTTPS; requests are restricted to their configured hosts.
  Changing a hostname alone does not establish compatibility with another Microsoft cloud.
- denied_pairs: client/scope pairs to skip in this environment, with "*" for all scopes
  of a client. Empty in the example: historical denials do not block a new investigation.

Use a separate token/config/results folder for each account or environment by setting
OUTLOOK_PROBE_HOME to a private directory; the default is .local/. Never commit that
directory or a real account configuration. Changing accounts requires a fresh directory
or deleting the old probe token file. All cached profiles must identify the same tenant
and account, even when expected_user is empty.

OWS mailbox routing is derived from the token's email, or its tenant/object identity.
The self-send body helper requires a token and derives its recipient from that token.
Missing identity is reported rather than filled with an example account.

## Running

From the repo root, with any Python ≥ 3.11. On Windows, set `PYTHONUTF8=1` so non-ASCII query
text prints correctly.

```bash
python research/probes/auth.py graph        # sign in: Outlook Mobile -> Graph Mail.Read
python research/probes/auth.py outlook       # sign in: One Outlook Web -> outlook.office.com
python research/probes/auth.py search      # usually silent, via the Outlook client's refresh token
python research/probes/auth.py --status
```

`auth.py` prints a device-code link and code. Sign in in any browser. Tokens are stored as
**plaintext** in `.local/probe-tokens.json` (git-ignored). Delete that file when you are done.
The probes check the optional expected account and reject mixed-account caches.
These helpers are independent of the production app's encrypted MSAL cache.

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

First authenticate each client needed for the run. graph_scopes.py needs a cached
refresh token for each configured client; if a baseline scope fails, try an administrator-
approved baseline such as that resource's .default in the local profile configuration.
A missing refresh token, skipped pair, transport failure, or expired refresh token does
not establish that a capability is unsupported. Inspect the OAuth error and distinguish
a scope denial from a sign-in problem. Only a successful API probe establishes usable behavior.

Steps 2–5 need graph. Step 4 also needs search, and step 7 needs outlook.
Route capabilities according to your results: documented Graph where usable, OWS for proven
gaps, and unavailable or unresolved where neither is established. Update the findings with
date, configuration context (without private identifiers), evidence and remaining uncertainty.

## Offline checks

The probe helpers have synthetic, network-free checks separate from the application suite:

    python -m unittest discover -s research/probes -p test_probes.py
    python -m compileall -q research/probes

## Output and safety rules

- Output is statuses, counts, timings, shapes, error codes and salted short hashes. No tokens,
  bodies, subjects, addresses or raw ids are printed. The exceptions are query strings and case
  labels the operator supplied.
- Save runs under `.local/probe-results/` (git-ignored) if you want to keep them:
  `python research/probes/catalog.py > .local/probe-results/catalog.json`.
- Historical AADSTS65002 results belong to the dated research record. Add confirmed
  denials to your local denied_pairs if you want later runs to skip them. Remove entries
  deliberately when reassessing changed policy; the probes never retry denials automatically.
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
4. Save the export in a private temporary folder **outside this repository** and inspect it **in place**. Search
   the JavaScript for the action or route name, and match responses by the marker.
5. Record only route names, request/response **shapes**, key names and counts in
   `docs/outlook-api-research.md`. Delete the export when done.

**Captures are credentials.** They contain live access and refresh tokens, cookies, canaries,
mailbox content and addresses. Never commit them, share them, or turn them into test fixtures.
Use them only to learn request contracts, which are then re-implemented with our own app-owned
tokens. Browser credentials must never be reused (requirements v4 §4).

Probe profiles are named `graph`, `outlook`, and `search`. Keep local probe configuration keys consistent with these names and sign in with `auth.py graph` or `auth.py outlook` when a matching probe token is missing.
