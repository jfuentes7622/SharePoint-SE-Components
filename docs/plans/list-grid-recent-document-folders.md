# List/Grid recent-document folder browsing: execution plan

Status: implementation authorized by the maintainer's November 4, 2026 instruction
to implement this execution plan. Local implementation and validation are complete;
target-farm acceptance and human-reviewed release remain pending.
Prepared: November 4, 2026.

## Approved follow-up: folder browsing independent of pruning

The maintainer selected **Always browse folders inside List/Grid; checkbox controls
branch pruning only**. This supersedes the original opt-in/legacy-library split below:

- Detect document libraries by authoritative list BaseType; discovery errors are explicit.
- Always use in-control folder buttons, breadcrumbs, document filtering and five-minute
  refresh for libraries. Folder buttons never open document-link dialogs.
- With **Matching document folders** off, retrieve every permission-trimmed folder
  page independently of document filters, including empty and nonmatching branches.
  With it on, retain only accessible ancestors of matching documents.
- Preserve document link opening settings and honor **Show link to item**.
- Apply existing selection, deferred edit/refresh and current-folder upload protections
  in both checkbox states. Checkbox changes refresh pruning without clearing form binding.
- Standard lists retain their existing query, grouping and link behavior.
- Validate both checkbox states, initial discovery failures/retries, complete folder
  paging, stale source discovery and Grid draft/upload safety before packaging
  ListControl 1.0.0.120 and GridControl 1.0.0.98.

Follow-up local validation passed:

- Library helper/runtime regressions, column/date filters, row navigation, standard-list
  metadata, theme lifecycle, embedded ListControl host and rule-value designers.
- Root `npm test` (20 tests) and SE validation (19 projects), compared with
  `b4cd11e9f649e774fcac92fa575018eb089688b6`.
- Both projects' clean, build, ship bundle and ship package tasks, followed by fresh
  package inspection: ListControl 1.0.0.120 and GridControl 1.0.0.98.
- Local browser fixtures using the actual React controls: pruning off retains empty
  and nonmatching folders; folder navigation remains inside the control; Enter/Space,
  nested folders, Up and pruning toggles work. Document links retain ListControl's
  Dialog behavior; Grid selection survives pruning and clears on folder navigation.
  The temporary fixture server/files were removed. These mocks are not SE-farm acceptance.

## File-count follow-up

The footer reports matching files in the current folder and its entire subtree,
not the number of immediate folder/document rows. Two folders containing three
matching files therefore report **3 files in this folder and subfolders**.
Drill-down scopes the total to that folder's subtree; empty folders report zero.
The count uses the committed, filtered document snapshot before folder wrapping
and pagination. It is the same with branch pruning on or off. The navigation
status remains library-wide; ordinary list totals and entry-based pagination
are unchanged. ListControl release: 1.0.0.123; GridControl release: 1.0.0.101.

Local validation passed: both library helper/runtime suites; column/date filters,
row navigation and ordinary-list metadata regressions; root tests (20 passed);
SE validation against the base above (19 projects); both clean/build/ship
bundle/package sequences and fresh package inspection. Target-farm acceptance
and deployment remain pending.

## 1. Confirmed requirements

The maintainer additionally requested hyperlink-style folder navigation and
backslash-separated breadcrumbs. Folder rows and the root breadcrumb now use
SharePoint's native `/_layouts/15/images/folder.gif` asset instead of the outline
font glyph. Links retain in-control navigation, selection clearing and Enter/Space
activation; Up has link styling, while toolbar command styling remains unchanged.
Release targets: ListControl 1.0.0.121 and GridControl 1.0.0.99.
Local runtime/helper/navigation regressions, root tests/base SE validation, both
production builds and fresh package inspections passed for these versions.
Browser fixtures verified underlined anchors with transparent backgrounds and no
button borders, backslash-separated nested paths, in-control link activation and
Up navigation. The native icon URL was checked; the mock server
does not contain SharePoint's images, so actual icon availability/appearance still
requires farm confirmation. The temporary fixture server and script were removed.

Implement option 2 in both ListControl and GridControl:

- Show documents matching filters authored through the existing control-level
  filter builder. A rolling Modified window is an example, not a dedicated mode
  or hard-coded requirement.
- Show only folders that contain matching documents, directly or through nested
  descendants. Keep every ancestor needed to reach those documents.
- A folder's own Modified value must not determine whether its branch appears.
- Use familiar folder drill-down, visible folder icons, breadcrumbs and Up.
- Refresh automatically every five minutes and retain manual Refresh.
- Use the library's default/configured view as the source, with filter
  conditions configured in the control.
- Enhance relative-time expressions so authors can combine
  `Modified >= now(-24h)` AND `Modified <= now` using existing filter conditions.
  Do not add a Last 24 hours option, preset, checkbox or separate date-filter UI.
- Harden document URLs, including libraries below the site-collection root.

The maintainer confirmed Modified for the example filter, drill-down rather than
an expandable tree, and five-minute automatic refresh. Authors remain free to
choose other fields, time offsets and supported filter combinations.

## 2. Scope and compatibility

- Add an opt-in **Matching document folders** library mode. Existing instances,
  standard lists and the current view-based display remain unchanged by default.
- Keep SPFx 1.5.1, TypeScript 2.4.2, existing React/tooling and SE-compatible
  REST/SPHttpClient patterns. Preserve Online behavior.
- Preserve solution, feature and component IDs and deployment scopes.
- Do not add Graph, modern SPO-only APIs, UI dependencies or a cross-project
  runtime package. Each top-level project remains independently buildable.
- Reuse existing filter, paging, button, diagnostics and URL helpers where
  suitable. Keep any necessary equivalent helpers local to each project.
- No folder creation, rename, move, folder deletion, arbitrary-folder starting
  location, or automatic tab switching in this release.
- Tab and Dynamic Forms should not require new packages for this feature.

See [Architecture](../Architecture.md), [SPFx baseline](../SPFx-Version.md),
[SE limitations](../SharePointSE-Limitations.md) and
[Agent Workflow](../Agent-Workflow.md).

## 3. Intended behavior

### Filtering and source view

1. Retain the selected/default view's document restrictions, displayed columns
   and applicable ordering.
2. In the new mode, explicitly query documents recursively throughout the
   library instead of inheriting a root-folder-only scope.
3. Apply the complete document filter expression, including preset filters and
   supported runtime column filters. Keep existing AND/OR and multi-select
   semantics; folder names must not independently satisfy document filters.
4. Build the visible folder hierarchy from the final matching documents.
   Ancestor folders are navigation context, not date-filtered document results.
5. Do not silently discard the source view's Where clause. Recommend an
   unfiltered default view when the control should own the date restriction.
   Documents excluded by other source-view conditions remain excluded.
6. Folder pruning must use the configured filters, not an injected Modified or
   24-hour condition. With no control filters, show the documents allowed by the
   source view and their nonempty ancestor branches.

### Relative time expressions in existing control filters

- Keep the existing field/operator/value-type/value/logical-condition workflow.
  Authors enter both conditions themselves; do not generate or insert them
  automatically.
- Extend the existing Expression value type with elapsed-hour offsets using
  `now(<signed integer>h)`, for example `now(-24h)`, `now(-6h)` or `now(-48h)`.
  This is generic relative-time support, not a special interpretation of 24.
- Example configuration in the existing filter builder:

  | Field | Operator | Value type | Value | Combine with previous |
  | --- | --- | --- | --- | --- |
  | Modified | Greater than or equal | Expression | `now(-24h)` | First condition |
  | Modified | Less than or equal | Expression | `now` | AND |

- Capture one UTC reference instant T per refresh/query generation. Resolve
  elapsed-hour offsets from that instant. The example above evaluates to
  `[T - 86,400,000 milliseconds, T]`, inclusive; other offsets and operators use
  their configured values without a fixed 24-hour limit.
- Use that same reference instant for server predicates, client predicates,
  all paging requests and folder pruning within the generation.
- Preserve the current meanings of `today`, `date(n)`, bare `now`, static
  values and expression mode. In particular, `date(-1)` remains calendar-based.
- Update expression validation, help and the shared evaluation surfaces
  consistently. Invalid offsets must show validation errors.
- Make read-only Created/Modified fields available to the filter designer
  without making them editable in GridControl.
- Reevaluate configured expressions and reload documents at refresh time, not
  continuously between the five-minute refreshes. Static and non-date filters
  remain valid and do not require a relative-time condition.

### Folder presentation and interaction

- At the library root, show matching documents stored directly at root and
  immediate child folders with matching descendants.
- Opening a folder shows its matching immediate documents and qualifying child
  folders. Breadcrumbs and Up navigate inside the control.
- Folder entries always show the built-in SharePoint/Fabric folder icon,
  including when command buttons use text-only mode. Icons are decorative;
  the folder link has an accessible name and keyboard activation.
- Keep folders first, ordered by name. Apply the view/current document sort to
  documents at the current level.
- Folder hierarchy replaces custom grouping in this opt-in mode. Show an
  explanatory settings note, preserve the saved grouping configuration, and
  restore its normal behavior when the author returns to the legacy mode.
- Folder activation is navigation, not document selection: do not publish
  folder item IDs to connected forms or enable document actions on them.
- Keep document row selection separate from opening its document link, as well
  as from the existing selected-row Go to details action.
- Preserve a selected document's Dynamic Data binding during background refresh;
  do not republish unchanged selection or steal focus.
- If the selected document ages out, indicate that it no longer appears in the
  results without automatically switching the connected form to another item.
- Explicit folder navigation clears the current document/bulk-delete selection
  and publishes the normal empty selection rather than leaving hidden selected
  items actionable. This proposed navigation rule is part of plan approval.
- Filter/query changes return to root. A refresh retains the current folder
  when it still has matches; otherwise move to its nearest visible ancestor and
  explain why. Use a real empty state when no documents match.
- Folder entries do not inflate matching-document counts or Dynamic Data
  document-result counts. Page the current level's folder/document rows
  consistently and distinguish row pagination from matching-document totals.

### Grid-specific behavior

- Folder rows have navigation only: no inline edit, history, document upload,
  Go to details, or bulk-delete checkbox on the folder row itself.
- The grid's Add document command uploads into the current browsed folder;
  root browsing uploads into the library root.
- Preserve file-collision policy, metadata validation and cleanup of a newly
  uploaded file when its metadata save fails. Never clean up an existing file.
- Save/cancel and upload completion reconcile the current filtered hierarchy.
  Explain successful uploads that do not match the other active filters.
- Defer automatic refresh and folder/filter changes while an inline edit,
  new-document draft or save is active. Show a pending-refresh notice, then
  refresh after save/cancel. Never discard draft values.
- Read-only users retain browsing and document links without edit commands.

## 4. Data and URL design

### Complete matching documents, then ancestor context

1. Resolve the authoritative library BaseType and root ServerRelativeUrl.
2. Request support fields regardless of whether they are visible columns:
   ID, file/folder type, FileRef, FileDirRef, FileLeafRef, Modified, Created and
   fields needed by active filters/sorting.
3. Use recursive, file-only candidate retrieval with the view's document Where
   restrictions and translatable control predicates. Prefer server-side
   filtering for configured DateTime ranges to avoid downloading unrelated
   documents; do not hard-code a Modified range.
4. Evaluate any remaining supported predicates client-side only after paging
   the complete candidate set. Never prune folders using only the first batch.
5. Derive and deduplicate ancestor paths for matching documents, stopping at
   the actual library root. Do not infer the root from the library display name.
6. Resolve required ancestor metadata through compatible, permission-trimmed
   SharePoint requests. Cache/deduplicate lookups per library/refresh and bound
   concurrency; avoid one ancestor request per document.
7. Build a path-keyed hierarchy and atomically publish a complete snapshot.
   Ignore stale responses after a filter, library or refresh-generation change.

Do not silently remove untranslatable filter conditions or broaden source-view
restrictions. A failed paging/ancestor request must be logged and surfaced, not
presented as an empty successful result. Retain the last complete snapshot with
a visible stale-data/error notice on refresh failure; show an error on initial
failure. Do not fabricate accessible folders when SharePoint denies access.

Large-library testing must check list-view-threshold behavior and indexes for
the configured filters, such as Modified. Document administrator prerequisites; do not
automatically create indexes or change library views/permissions.

### Harden URL handling

- Treat SharePoint FileRef/ServerRelativeUrl as authoritative document locations.
- Resolve server-relative paths against the SharePoint origin, not by appending
  them to the complete current-web URL. Resolve web-relative paths against the
  appropriate web base, with a trailing slash.
- Preserve valid escaping and avoid double encoding or decoding path separators.
  Escape OData string literals separately from URL path encoding.
- Cover spaces, apostrophes, Unicode and literal percent/hash/ampersand/plus
  characters where supported by the farm.
- Permit only appropriate HTTP(S) document destinations; do not turn unsafe
  schemes into clickable links.
- Do not send document-library rows through a fabricated
  `/Lists/<display title>/DispForm.aspx` fallback.
- In the new mode, folder links drill down locally and document links open the
  document. Keep optional metadata/detail navigation distinct and documented.
- Audit compatible existing URL helpers before extracting or adding helpers.
  Preserve ordinary list-item and external URL-column behavior.

## 5. Execution sequence

### Phase A: approval and fixtures

- Obtain approval of this complete design and authorization to implement.
  For GitHub issue execution, follow the proposal/approval workflow.
- Use an unprotected feature branch and record the actual PR base commit.
  Preserve all unrelated existing worktree changes.
- Add representative library/page fixtures and failing focused tests before
  changing the retrieval pipeline.

### Phase B: relative filters and URL helpers

- Update field discovery and expression validation/help/evaluation in both
  existing filter builders to support generic elapsed-hour offsets.
- Verify authors can construct the two-condition Modified example without a
  new 24-hour option or automatic condition generation.
- Add the per-query reference clock and safe URL/path normalization.
- Verify old expressions and ordinary list links remain unchanged.

### Phase C: recursive query and folder index

- Add opt-in library-mode props/state and authoritative root discovery.
- Implement recursive candidate retrieval, complete paging, document
  predicates, ancestor resolution and stale-request cancellation/ignoring.
- Build and test the pruned hierarchy before adding UI.

### Phase D: folder UI and document actions

- Add visible folder icons, folder drill-down, breadcrumbs, Up and empty states.
- Wire sorting, column filters, pagination, counts and document links.
- Audit grouping, sticky headers, action-column spans, keyboard handling,
  selected-row navigation and Dynamic Data selection for both row kinds.
- Adapt GridControl upload location, document-only actions and edit guards.
  Keep legacy grouping/display behavior unchanged outside the new mode.

### Phase E: five-minute refresh

- Add a non-overlapping 300,000 ms refresh cycle plus manual Refresh.
- Refetch documents, reevaluate configured expressions and prune branches;
  do not merely rerun filters over stale cached rows.
- Preserve folder context and focus, defer around grid drafts/saves, report
  pending or failed refreshes, and remove timers on disposal.

### Phase F: verification and release

- Run focused regressions, root checks, production builds and package inspection.
- Update both READMEs with settings, filter semantics, library/view requirements,
  folder navigation, selection rules and refresh behavior.
- Increment each affected project's npm/solution/feature versions consistently
  from the versions present at execution; preserve all IDs and deployment flags.
- Deliver fresh ListControl and GridControl packages and a human-reviewed PR.
- An authorized operator performs farm acceptance before release; do not treat
  local tests or package validation as App Catalog deployment evidence.

## 6. Primary implementation surfaces

- [ListControl web part](../../ListControl/src/webparts/listControl/ListControlWebPart.ts):
  properties, filter metadata/designer, runtime props and settings.
- [ListControl component](../../ListControl/src/webparts/listControl/components/ListControl.tsx):
  retrieval, filter evaluation, folder presentation, links, selection and refresh.
- [GridControl web part](../../GridControl/src/webparts/gridControl/GridControlWebPart.ts):
  equivalent settings/designer/runtime wiring.
- [GridControl component](../../GridControl/src/webparts/gridControl/components/GridControl.tsx):
  equivalent browsing plus upload/edit/delete protections.
- Each project's component CSS, localized strings, README, manifests/versions
  and fresh solution package; add small project-local helpers if justified.
- Reuse the source-transpilation/mocked-host patterns under [tests](../../tests).
  Extend focused coverage without adding a modern component toolchain.

## 7. Acceptance tests

| Area | Required verification |
| --- | --- |
| Pruning | Root recent document remains; old folder A containing nested folder B with a recent document remains; old-only branch C and empty branch D disappear. |
| Filter configuration | Build the example using two ordinary Expression conditions joined with AND. No dedicated 24-hour UI, preset or injected filter exists. |
| Parent age | A folder outside the configured document date range remains when a descendant document matches. |
| Window example | With `Modified >= now(-24h)` AND `Modified <= now`, include T and T-24h; exclude T+1 ms and T-24h-1 ms. Test UTC/local display and DST changes using elapsed milliseconds. |
| Generic filtering | Verify other offsets such as `now(-6h)` and `now(-48h)`, non-date filters, and no control filters. Folder pruning follows those conditions rather than assuming a 24-hour window. |
| Modified vs Created | An old-created, recently modified document matches the authored Modified conditions. Changing the field to Created changes the results appropriately. |
| Source view | Default root-folder scope is intentionally lifted; document Where conditions and configured columns/order are preserved. |
| Other filters | Choice/Lookup multi-selection, text/date filters and AND/OR rules prune branches from final matching documents. |
| Paging | A matching document found beyond the first batch still produces its full ancestor chain. |
| Identity | Equal folder names in different paths remain distinct; no path escapes the library root. |
| Navigation | Drill-down, breadcrumb and Up work across several nested levels without publishing folder IDs. |
| Icons/accessibility | Every folder shows its icon in all button modes; keyboard activation and accessible labels work; selection and ordinary text interaction do not navigate. |
| URLs | Root-web/subsite libraries and renamed library URLs work without duplicated site paths, broken escaping or unsafe schemes. |
| Grid safety | Folder rows cannot invoke document actions; uploads target the current folder; failed metadata saves clean up only the newly uploaded file. |
| Refresh | Fake-clock tests verify a five-minute cycle adds newly modified files and removes expired branches; requests do not overlap and stale completions cannot overwrite newer results. |
| Edit/focus safety | Grid drafts/saves defer refresh; unchanged document selection is not republished; connected forms are not switched by background refresh; focus is not stolen. |
| Errors/permissions | Initial failure, later-page failure, ancestor denial and refresh failure are explicit; no false empty-state success or unauthorized synthetic folder entries. |
| Compatibility | Standard lists, legacy library mode, grouping, filters, link options, selected-row navigation and existing grid editors retain their behavior. |
| Scale/farm | Test a library above 5,000 documents with appropriate indexes, nested paths and restricted permissions on the actual SE farm. |

Use the existing standalone filter, metadata, theme, row-navigation and host
tests, plus new focused library-folder, relative-time, URL and refresh tests.
Add browser checks with nested folders and more than one fetch batch.

Required repository checks, using the repository-script Node toolchain:

```powershell
npm test
npm run validate-se -- --base <actual-PR-base-commit>
```

In each affected project, using its legacy build toolchain, check every exit code:

```powershell
.\node_modules\.bin\gulp.cmd clean
.\node_modules\.bin\gulp.cmd build
.\node_modules\.bin\gulp.cmd bundle --ship
.\node_modules\.bin\gulp.cmd package-solution --ship
```

Inspect the freshly generated packages from the repository root:

```powershell
.\scripts\validate-package.ps1 -Project ListControl
.\scripts\validate-package.ps1 -Project GridControl
```

Report failed/unavailable checks, actual farm build/CU, App Catalog scope,
deployment results and rollback evidence as described in
[Deployment](../Deployment.md).

## 8. Local execution evidence (November 4, 2026)

- Working branch: `feature/list-grid-matching-document-folders`; base commit:
  `b4cd11e9f649e774fcac92fa575018eb089688b6`. Existing unrelated worktree changes
  were preserved. No commit, push, PR or deployment was performed.
- Fresh production packages: ListControl `1.0.0.118`, GridControl `1.0.0.96`.
  SPFx 1.5.1, IDs and deployment scopes remain unchanged.
- Passed new library helper/runtime regressions for both controls: exact inclusive
  time boundaries, other hour offsets/DST, no injected date filter, source query,
  complete paging, ancestor pruning/denial, safe paths/URLs, keyboard folder rows,
  document counts, stale generations, five-minute refresh, draft deferral, folder
  retention after cancel, current-folder uploads and new-file-only cleanup.
- Existing column-filter (47), row-navigation (47), date-filter (33),
  control-metadata (24), theme-lifecycle (82), embedded-host (21) and type-aware
  rule-designer checks passed. Root `npm test`: 20 passed. SE static validation
  against the base passed for 19 projects.
- Both projects passed clean/build, production bundle/package and inspection of
  their freshly generated solution packages with the legacy Node 10 toolchain.
- Browser fixture checks passed for both controls with nested folders and multiple
  fetch batches: Enter/Space drill-down, Up, document links, selection clearing,
  no document actions on folders, and manual refresh retaining keyboard focus.
  The local fixture verifies folder-icon elements; actual SharePoint icon-font
  rendering remains part of farm acceptance.
- Unavailable locally: actual SE farm build/CU, App Catalog deployment, restricted
  farm permissions, farm-specific percent/hash path support, >5,000-document
  threshold/index behavior and rollback evidence. A human-reviewed PR and authorized
  farm acceptance are still required before release.
