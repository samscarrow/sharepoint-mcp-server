# Implementation Backlog

Prioritized improvements to the bay-view-graph MCP server. Open work first; shipped
work logged at the bottom.

---

## Open

### 1. SharePoint URL resolution — consolidation (LOW — functionally already works)

Investigated: all four file tools already resolve viewer/sharing/path URLs.
`get_file_content`/`download_file` go through `resolveContentEndpoint` and
`share_file`/`create_share_link` through `resolveDriveItemEndpoint`, both routing any
`http` URL through Graph `/shares/{encoded}/driveItem` — so `Doc.aspx`, `/:w:/r/…`,
and `IQC…` links already work, and the only real bug (extension detection) is fixed.
Remaining is a *pure DRY refactor* merging the two resolvers — but they differ in auth
(content path uses the app token for tenant; item path uses delegated), so merging is
risk > reward on four working tools. **Defer unless the duplication actually bites.**

### 2. Upload check-in for check-out-required libraries (LOW — not yet hit)

If a library enforces "require check-out", a simple `PUT /content` can leave the new
version as an unpublished draft. Auto check-in after `upload_file` when that state is
detected. Not triggered by current libraries (worship library publishes fine).

### 3. New capabilities (LARGER — net-new reach, optional)

From the assistant-capabilities review: scheduling (`find_meeting_times`,
`get_schedule`), people/presence (`search_people`, `get_presence`), tasks (Microsoft
To Do / Planner — scopes already provisioned), and Teams chat/channels. Bigger build;
additive rather than fixing a broken path.

---

## Shipped

### Drafts flagged as unsent in email results
Unsent drafts were returned inline with sent/received mail by `search_emails`
(hit live 2026-07-07: a keyword search for "Dorrien" surfaced an unsent "Arrival
time" draft as the top hit, with a populated `to`+`date` and nothing marking it a
draft — an agent could conclude the mail already went out and silently drop the
task). Fix: added `isDraft` to `$select` and the returned object on
`search_emails`, `list_emails`, and `get_email`. `search_emails` also appends a
note when any result is a draft (`N result(s) are unsent drafts … NOT proof a
message was sent`), same class as the existing "capped search ≠ proof of absence"
flagging. Tool descriptions updated to advertise `isDraft`. Parent-folder surfacing
skipped — `parentFolderId` is opaque and `isDraft` is the actionable signal.

### search_emails deterministic mode oldest-first bug
Filter mode (`from`/`to`/`after`/`before`) had no `$orderby`, so Graph returned
filtered `/me/messages` oldest-first; a capped page returned the OLDEST matches and
the in-page client sort couldn't surface newer mail on later pages — a recent reply
from a known sender (e.g. `from:anthonyp133@gmail.com`, 45 hits) looked absent.
Fixed with `$orderby=receivedDateTime desc`, made legal by leading the filter with a
`receivedDateTime ge …` clause (1900 floor when no `after`), per Graph's
`$filter`/`$orderby` ordering rule for messages. Newest-first, exact `total`
preserved, ordered pagination. Verified live.

### create_draft — save an unsent draft to tweak and send
New `create_draft` tool: builds a message with fully populated recipients
(`to`/`cc`/`bcc`) and subject (both required; body optional) and POSTs it to
`/me/messages`, which saves it in the Drafts folder rather than sending. Mirrors
`send_email` (signature auto-append, recipient coercion) but returns the draft
`id` and `webLink` so it can be opened in Outlook, edited, and sent manually.

### Email retrieval overhaul — COMPLETE (PRs #1, #2)
Root cause: `search_emails` was Graph `$search` only (ranked, capped, preview-only, no
`conversationId`/count/filters), so a live session failed to find a known email.
Retrieval by **thread** and by **sender/date** are now both first-class — no keyword
guessing, and capped-search results are flagged as not proof of absence.
- `get_thread` — full conversation by `messageId`/`conversationId`, chronological.
- `conversationId` exposed on `get_email`, `list_emails`, `search_emails`.
- `search_emails` deterministic mode (`from`/`to`/`after`/`before`/`folder`) via
  `$filter`, with exact `total` (`$count` + `ConsistencyLevel: eventual`); keyword
  mode unchanged with `hasMore` + not-proof-of-absence note. Verified live.
- `list_emails` `$orderby`+filter 400 (`InefficientFilter`) fixed — omit `$orderby`
  when a `$filter` is set, sort client-side.
- Not done (low value): subject `$filter` (Graph mail `contains(subject,…)` unreliable
  — use the `query` KQL `subject:…` path); `fullBody` flag (get_thread returns bodies).

### SharePoint viewer-link extraction (PR #1)
`pathForExtension` reads the `file=`/`sourcedoc` query param, so `Doc.aspx` viewer
links extract to text instead of returning UTF-8-mangled bytes. (Consolidation still
open — see Open #1.)

### .docx text-extraction attribute leak (PR #1)
Paragraph-tag attributes (`w:rsid…`) no longer bleed into extracted text.

### docx-edit structural modes
`scripts/docx-edit.mjs` gained `--prune-empty-runs` (drop content-less `<w:r>`
shells), `--set-margin top=…,bottom=…` with `--margin-section all|continuous` (the
mid-page-gap fix), and `--insert-after "<anchor>" "<text>"` (clone a paragraph's
style — pPr + first run's rPr — drop its section break, insert after the anchor).
Verified on a fixture; bundle regenerated. Replaces the hand-rolled JSZip surgery
used on the Skylar contract.

### File round-trip tooling (earlier session)
`upload_file localFilePath` + MIME inference; `download_file` (raw bytes to disk);
`scripts/docx-edit.mjs` (cross-run, formatting-preserving .docx edits) +
self-contained `docx-edit.bundle.mjs`.
