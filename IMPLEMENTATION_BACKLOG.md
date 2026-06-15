# Implementation Backlog

Prioritized improvements to the bay-view-graph MCP server. Newest analysis on top.

---

## 1. Email retrieval overhaul (HIGH — root cause of real retrieval failures)

**Problem.** `search_emails` is the only discovery tool and it is Microsoft Graph
`$search`: relevance-ranked, top-N capped, body-preview-only, and it exposes no
`conversationId`, no result count, and no sender/date filter
(`src/index.ts` ~L2574). Consequences observed in a live session (failing to find
a known title-request email):

- The only discovery strategy is **guessing keywords**. If the target text sits in
  a quoted/down-thread reply or uses different wording, it never ranks into the
  capped top-N and is silently missed.
- **No thread retrieval.** You can read one message (`get_email`) but cannot pull a
  whole conversation in one step — so "open the thread and read all of it" is
  impossible, and a known thread gets walked past.
- **No truncation signal.** Results don't say how many matched, inviting the
  absence-of-evidence trap ("10 results" read as "only 10 exist").
- **Preview-only bodies** force N re-fetches to judge relevance.

**Changes:**

- [x] **`get_thread`** (highest leverage): given a `messageId` (or `conversationId`),
  return every message in the conversation, full bodies, chronological. *Shipped* —
  resolves conversationId from messageId, filters `/me/messages` by conversationId
  across all folders, sorts client-side (Graph rejects `$filter`+`$orderby` on
  different properties).
- [x] **Expose `conversationId`** in `get_email`, `list_emails`, and `search_emails`
  output so any found message can pivot to its thread. *Shipped.*
- [x] **Truncation visible (partial)**: `search_emails` now returns `hasMore` + a
  note that absence isn't proof; `get_thread` flags `hasMore`. *Shipped.*
- [x] **`list_emails` $orderby bug (CONFIRMED live):** the handler always appended
  `$orderby=receivedDateTime desc`, which Graph rejects with `InefficientFilter`
  whenever a `from/…` (recipient/sender) filter is present — so the deterministic
  "all of X's mail" path 400'd. *Shipped* — `handleListEmails` now omits `$orderby`
  when a `$filter` is set and sorts the page client-side by `receivedDateTime desc`
  (same pattern get_thread uses); keeps server-side sort on the unfiltered path.
- [x] **Deterministic filters** on `search_emails`: `from`, `to`, `after`, `before`,
  `folder`. When present, uses `$filter` (complete, sorted client-side) instead of
  `$search` (ranked, capped); modes can't combine, filter wins. *Shipped & verified
  live* (`from`+date range returns exact match with `mode:"filter"`). Subject still
  routes through the `query` KQL path (`subject:…`) — `contains(subject,…)` isn't
  reliably supported in Graph mail `$filter`.
- [x] **`@odata.count`** on the filter path (`$count=true` + `ConsistencyLevel:
  eventual`) → exact `total` in the response. *Shipped & verified.* (Not available on
  the `$search` path — Graph disallows `$count` with `$search` on messages; that path
  keeps `hasMore` + the not-proof-of-absence note.) Fuller bodies / `fullBody` flag
  not done — low value given get_thread returns full bodies.

Behavioral lesson the tools should make easy: retrieve by **thread/sender**, not by
guessed content; never assert non-existence from a ranked, capped search.

---

## 2. Robust SharePoint URL resolution (MEDIUM-HIGH — fixes a silent-corruption bug)

**Status (partial — core silent-corruption fixed):** `pathForExtension` now extracts
the real filename from the `file=` (fallback `sourcedoc`) query param when it carries
an extension, falling back to the URL pathname — so `Doc.aspx?…&file=X.docx` viewer
links are detected as `.docx`/`.xlsx`/… and parsed to text instead of returning
mangled bytes. Remaining (smaller follow-up): a shared `resolveFileIdentity(url)` that
also normalizes viewer/sharing-token URLs to a single driveItem for
`resolveContentEndpoint` and consolidates the logic across get_file_content,
download_file, share_file, create_share_link.

`get_file_content` / `download_file` detect file type from the URL **path**, but a
SharePoint `Doc.aspx?…&file=X.docx` viewer link carries the real filename in the
`file=` query param — so `.docx` isn't detected and it falls through to the
raw-bytes branch, returning UTF-8-mangled binary instead of extracted text. Sharing
tokens (`/:w:/r/…`, `IQC…`) are also in play.

**Change:** a `resolveFileIdentity(url)` helper that (1) extracts the real filename
from `file=` (fallback `sourcedoc`) for extension/MIME detection, (2) normalizes
viewer / sharing / path URLs to one driveItem, (3) returns the content endpoint and
correct extension. Have `pathForExtension` and `resolveContentEndpoint` consume it.
Makes pasting *any* SharePoint link "just work" across get_file_content,
download_file, share_file, create_share_link.

---

## 3. docx-edit structural modes (MEDIUM — carried from session)

Add to `scripts/docx-edit.mjs`: `--prune-empty-runs` and `--set-section-margin`
(and consider `--insert-after "<anchor>" "<paragraph>"` for true paragraph
insertion). Both had to be hand-rolled during a live contract edit.

---

## 4. Upload check-in for check-out-required libraries (LOW — not yet hit)

If a document library enforces "require check-out", a simple `PUT /content` can
leave the new version as an unpublished draft. Add an auto check-in after
`upload_file` when that state is detected. Not triggered by current libraries
(worship library publishes fine), so low priority.

---

## 5. New capabilities (LARGER — net-new reach, optional)

From the assistant-capabilities review: scheduling intelligence
(`find_meeting_times`, `get_schedule`), people/presence (`search_people`,
`get_presence`), tasks (Microsoft To Do / Planner — scopes already provisioned),
and Teams chat/channels. Bigger build; additive rather than fixing a broken path.
