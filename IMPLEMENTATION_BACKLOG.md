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
