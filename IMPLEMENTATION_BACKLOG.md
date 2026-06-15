# Implementation Backlog

Prioritized improvements to the bay-view-graph MCP server. Open work first; shipped
work logged at the bottom.

---

## Open

### 1. SharePoint URL resolution — consolidation (MEDIUM)

Core silent-corruption bug is fixed (`pathForExtension` reads the `file=`/`sourcedoc`
query param so `Doc.aspx?…&file=X.docx` viewer links extract correctly — see Shipped).
Remaining: a shared `resolveFileIdentity(url)` helper that also normalizes
viewer/sharing-token URLs (`/:w:/r/…`, `IQC…`) to a single driveItem for
`resolveContentEndpoint`, consolidating the logic across `get_file_content`,
`download_file`, `share_file`, `create_share_link` so pasting *any* SharePoint link
works uniformly (not just for text extraction).

### 2. docx-edit structural modes (MEDIUM)

Add to `scripts/docx-edit.mjs`: `--prune-empty-runs` and `--set-section-margin`, and
consider `--insert-after "<anchor>" "<paragraph>"` for true paragraph insertion. All
three had to be hand-rolled during a live contract edit.

### 3. Upload check-in for check-out-required libraries (LOW — not yet hit)

If a library enforces "require check-out", a simple `PUT /content` can leave the new
version as an unpublished draft. Auto check-in after `upload_file` when that state is
detected. Not triggered by current libraries (worship library publishes fine).

### 4. New capabilities (LARGER — net-new reach, optional)

From the assistant-capabilities review: scheduling (`find_meeting_times`,
`get_schedule`), people/presence (`search_people`, `get_presence`), tasks (Microsoft
To Do / Planner — scopes already provisioned), and Teams chat/channels. Bigger build;
additive rather than fixing a broken path.

---

## Shipped

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

### File round-trip tooling (earlier session)
`upload_file localFilePath` + MIME inference; `download_file` (raw bytes to disk);
`scripts/docx-edit.mjs` (cross-run, formatting-preserving .docx edits) +
self-contained `docx-edit.bundle.mjs`.
