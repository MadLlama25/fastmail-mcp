# Fastmail MCP — vision draft

**Status:** proposal, awaiting Jared review
**Author:** Jarvis (autonomy run `autonomy-fastmail-mcp-20260503T021424Z`)
**Date:** 2026-05-03

This is the first vision-level proposal for `fastmail-mcp`. Per
`autonomy.yaml`, the first autonomy run on this project is
propose-mode only — no implementation. The goal here is to give Jared
a coherent target to push on, accept, or redirect.

---

## 1. Premise

`fastmail-mcp` is a TypeScript MCP server (1.9.4) that fronts
Fastmail's JMAP API for email/contacts and CalDAV for calendars. It
exposes 43 tools across listing, sending, threading, attachments,
bulk ops, contacts, and calendar events. It is the live, primary
inbox surface for Jared's Jarvis assistant — meaning regressions
break daily workflows immediately.

Today it is healthy: tests across auth, JMAP, CalDAV, URL validation,
and coercion (~3.7k LOC of test against ~3.7k LOC of source); a
recent (5279615) gap-closure pass; v1.9.x security hardening; and a
DXT package for Claude Desktop. The codebase is **not in trouble.**
The question this proposal exists to answer is: *what does "great"
look like in 6 / 18 months, given that "working" is already true?*

## 2. End-goal (3-year horizon)

`fastmail-mcp` becomes the **canonical reference implementation** for
"how an LLM should talk to a personal-grade JMAP server," not just
"a Fastmail wrapper." Specifically:

- A. **LLM-shaped, not API-shaped.** Tools return digest views by
     default, with opt-in projection for full payloads. The 43-tool
     surface today reflects the JMAP API's shape; the vision surface
     reflects what an agent actually needs to do (triage, draft,
     thread, schedule).
- B. **Safe by default for a live inbox.** Every destructive op has a
     dry-run / preview path. Every bulk op caps at a sensible default
     and surfaces the cap. Identity / from-address mistakes are caught
     before the JMAP submission, not after.
- C. **Stateful where it pays off, stateless everywhere else.** A
     tiny in-memory cache for mailbox-id ↔ name resolution and
     identity lookup (the two things every other tool re-fetches).
     No persistent state on disk; restart = clean.
- D. **JMAP-calendar ready.** The CalDAV fallback is a transitional
     wart. When the IETF JMAP-calendars draft becomes an RFC and
     Fastmail ships it, the swap should be a single client-injection
     change, not a rewrite.

## 3. Six-month priorities (in rough order)

These are concrete, scoped enough to fit single PRs, and don't
require architectural decisions Jared hasn't made yet.

### P1 — Response-size discipline (resolves upstream issue #40)
- Default email-list and search responses to a digest projection
  (id, threadId, from, subject, sentAt, hasAttachment, isUnread).
  Add `fields: "full" | "digest" | string[]` opt-in.
- Cap thread / search result counts in the response envelope; emit a
  `truncated: true` + `nextPage` cursor when over.
- Acceptance: a default `list_emails(limit=20)` against a busy inbox
  fits in <8KB JSON. Full payloads still reachable via `get_email`.

### P2 — Update-op field-preservation invariant (resolves upstream issue #56)
- `update_calendar_event` and `update_contact` currently risk
  silently overwriting fields with empty strings. Establish a
  project-wide rule: *omitted = preserve, explicit `null` = clear.*
  Empty string is treated as "preserve" unless the field's domain
  treats `""` as a meaningful clear (none currently do).
- Acceptance: round-trip test "update with only one field changed
  leaves all other fields equal" passes for every update tool.

### P3 — Identity + mailbox resolution cache
- Single in-memory map, populated lazily, invalidated on
  `list_mailboxes` / `list_identities` calls. Used by `send_email`,
  `move_email`, `add_labels`, etc., to accept either an id or a
  human name (`"Inbox"`, `"Archive"`, an alias address).
- Acceptance: `move_email(emailId, targetMailbox: "Archive")` works
  without the caller knowing the mailbox id.

### P4 — Confirm-token pattern for destructive non-bulk ops
- `delete_email`, `delete_contact`, `delete_calendar_event`,
  `send_email` to a brand-new recipient: each returns a
  `confirmToken` + structured preview on first call; second call
  with the token executes.
- Opt-out: env flag `FASTMAIL_MCP_NO_CONFIRM=true` for non-interactive
  pipelines.
- Acceptance: an LLM that calls `delete_email` once gets a preview,
  not a deletion.

### P5 — Transactional logging hooks
- Optional structured log file (`FASTMAIL_MCP_LOG=/path`) capturing
  tool, params (with PII fields redacted to length+hash), result
  size, latency. No bodies, no addresses, no tokens.
- Acceptance: a week of inbox usage produces a log Jared can grep
  for "what did Jarvis just do at 03:14?" without disclosing email
  bodies.

### P6 — Mock-JMAP test harness
- Today's tests mock at the HTTP layer ad-hoc. Extract a reusable
  `MockJmapServer` (and `MockCalDav`) so adding a new tool requires
  ~10 lines of fixture, not 200. Capture+replay one real JMAP
  session (sanitized) as a golden integration test.
- Acceptance: adding a hypothetical new tool is faster end-to-end
  than today.

## 4. Non-goals (explicit)

- **Not a generic JMAP client.** Fastmail-specific quirks (CalDAV
  fallback, `urn:ietf:params:jmap:mail` capability assumptions) stay.
  If a generic JMAP MCP is wanted, it's a fork, not a rename.
- **Not a workflow engine.** No "rules", no "filters", no scheduling.
  Those live in Jarvis or in Fastmail's own sieve, not here.
- **Not multi-account.** One account per process. If multi-account
  is ever wanted, it's a `FASTMAIL_MCP_PROFILE` env switch + parallel
  processes, not in-server multiplexing.
- **No persistent state on disk.** Caches are in-memory only.

## 5. Premortem — how this vision could fail

- **F1: Tool-renaming breaks Jarvis prompts.** Any rename of an
  existing tool invalidates Jarvis's tool-use history and any saved
  prompts. Mitigation: P1–P5 add fields and behaviors but rename
  nothing in 6mo. Renames, if ever, ship behind a `v2/` namespace
  with both surfaces live for a deprecation window.
- **F2: Digest-default surprises power users.** A user who relied on
  full payloads gets digests and confused. Mitigation: digest is
  default only for *list* tools, never for `get_email` /
  `get_thread`; `fields: "full"` is documented in tool description.
- **F3: Confirm-token UX makes Jarvis chatty.** Two-call delete adds
  latency on routine archives. Mitigation: P4 applies only to
  `delete`/first-time-recipient. Bulk ops already have dry-run.
- **F4: Cache desync after web-UI changes.** User renames a mailbox
  in the Fastmail web UI; cache is stale. Mitigation: cache TTL is
  short (60s) and any 404/forbidden from a cached id triggers
  invalidation + retry once.
- **F5: Upstream divergence.** This repo is `MadLlama25/fastmail-mcp`
  upstream; Jared is a contributor, not the maintainer. Pursuing
  P1–P6 unilaterally risks fork-drift. Mitigation: each priority
  gets a GitHub issue first; implementation only after upstream
  signal of acceptance, OR Jared's explicit "fork it" call.

## 6. Acceptance criteria for "vision adopted"

This proposal is "adopted" when:

1. Jared has read it, marked sections he disagrees with, and written
   responses to the open questions in §7.
2. The §3 priority list is reordered or pruned to match Jared's
   actual constraints.
3. Each retained P-item has a tracking issue (GitHub upstream or
   local backlog) before any implementation begins.

This proposal is "rejected" cleanly if Jared writes "park it" or
similar — the autonomy log records that and the branch is left for
deletion.

## 7. Open questions for Jared (load-bearing)

These are forks where I genuinely cannot pick without Jared's
preference. Each one is queued to Discord per autonomy rule 4.

1. **Upstream vs. fork.** Do you want this work merged upstream into
   `MadLlama25/fastmail-mcp` (slow, requires Jeremy's review), or are
   you comfortable maintaining a Jarvis-flavored fork?
2. **Backwards-compat envelope.** Are you willing to flip the default
   for *list* tools to digest projection (P1), accepting that any
   external caller of this MCP — if there are any beyond Jarvis —
   sees a different shape? The alternative is opt-in only, which
   means LLM responses keep blowing up by default.
3. **Confirm-token scope.** Should P4 cover `send_email` to
   *any* new recipient (more friction, more safety), or only to
   external addresses outside the contacts list (less friction)?
4. **Multi-account "never" or "later"?** Comfortable putting
   multi-account on the explicit non-goal list, or want to leave it
   open?

## 8. Alternatives considered

- **"Do nothing, the project is fine."** Reasonable read; the
  codebase is healthy. Counter: issue #40 (huge JSON) and the
  silent-overwrite issue #56 are real LLM-ergonomic problems that
  won't fix themselves, and the 43-tool surface will keep growing
  unless there's a guiding shape.
- **"Rewrite as a generic JMAP MCP."** Bigger scope, no clear demand,
  Fastmail-specific code (CalDAV fallback, identity wildcards) is
  the actual value. Rejected as a non-goal in §4.
- **"Stateful with a SQLite cache for full email bodies."** Tempting
  for offline triage and search, but takes the project from
  "translator" to "client + store" — different maintenance burden,
  different security posture (now we're holding email content on
  disk). Rejected; in-memory id-resolution cache only.

---

## Appendix — current state (2026-05-03)

- v1.9.4, Node ≥20, `@modelcontextprotocol/sdk` ^1.29.0, `tsdav` ^2.1.8
- Source: `src/index.ts` (2133 lines, the MCP shell + tool handlers),
  `jmap-client.ts` (1526), `caldav-client.ts` (478),
  `contacts-calendar.ts` (455), plus auth/url-validation/coerce.
- Tests: ~3.7k LOC across 6 test files; HTTP-layer mocks, no
  end-to-end harness.
- Recent direction (last 10 commits): security hardening (1.9.x),
  Dependabot integration, four small surface-gap fixes, drop Node 18.
- Open upstream issues at time of writing: #40 (huge JSON), #56
  (silent overwrite), #62 (follow-up on #57 usage in Claude).

This appendix will go stale fast; treat the body of the proposal as
the durable artifact.
