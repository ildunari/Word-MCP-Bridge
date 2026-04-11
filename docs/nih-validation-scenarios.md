# NIH Validation Scenarios

This checklist is the live validation harness for the Word MCP Bridge tool surface.
Use the NIH document family as follows:

- Source of truth: `docs/NIH_Biocoating_CraftV6 source-of-truth.docx`
- Copy A: `docs/NIH_Biocoating_CraftV6.docx`
- Copy B: `docs/NIH_Biocoating_CraftV6 copy.docx`

Never let two agent lanes mutate the same working copy at the same time.

## Evidence rules

For every mutation scenario:

1. Read the exact target scope first.
2. Capture the requested operation in structured form.
3. Read the same scope again after mutation.
4. Compare the literal text exactly.
5. Fail the scenario if any unintended character drift appears.

Use both:

- live bridge readback from the open Word session
- offline `officecli` checks against the `.docx` file on disk

## Easy

1. Read overall document text and structure.
2. Read one known paragraph by index and verify its text matches `officecli view ... text`.
3. Append text at document end on Copy A.
4. Insert a real new paragraph on Copy A.
5. Replace one active selection on Copy A.
6. Search for the inserted text and verify the match count.

## Medium

1. Use `word_get_text_range` on a dense scientific paragraph.
2. Replace an exact substring using `word_replace_text_range`.
3. Edit text adjacent to punctuation, percentages, and citation brackets.
4. Replace text at both the beginning and end of a paragraph.
5. Verify no extra blank paragraph was introduced.
6. Verify neighboring paragraphs are unchanged.

## Review

1. List comments from Copy B.
2. Insert a native comment on a targeted paragraph in Copy B.
3. Reply to the inserted or existing comment.
4. Resolve the comment.
5. Reopen the comment.
6. Delete the comment.
7. Confirm comment count and resolved state through both live readback and offline XML inspection when needed.

## Track Changes

1. Enable track changes on Copy B.
2. Make a targeted text-range edit.
3. Verify a revision appears in `word_get_revisions`.
4. Read its scope using `word_get_revision_scope`.
5. Accept one revision.
6. Reject one revision.
7. Accept all revisions on a separate mutation branch if needed.
8. Reject all revisions on a separate mutation branch if needed.

## Formatting

1. Apply paragraph style to a known heading.
2. Change paragraph spacing/alignment on one paragraph.
3. Apply inline bold to an exact substring.
4. Apply italic to another exact substring.
5. Apply underline, highlight, and font color on separate substrings.
6. Verify text equality separately from formatting equality.
7. Confirm unrelated runs in the same paragraph did not drift.

## Hard

1. Search for a repeated scientific phrase with multiple occurrences.
2. Use exact paragraph+offset targeting to mutate only one occurrence.
3. Add a comment on the same paragraph.
4. Enable track changes.
5. Perform a second exact-range edit.
6. Read revision scope.
7. Accept or reject only the intended revision.
8. Verify final paragraph text exactly.
9. Verify comment state exactly.

## Connection and UI

1. Watch `office-bridge events` while the panel is open.
2. Refresh the taskpane and confirm the logs show a hello path, not a silent failure.
3. Confirm tool inventory returns after reload.
4. Open both NIH copies and confirm the taskpane still matches the current document.
5. Confirm the helper app does not increment `connectionDropCount` during ordinary polling.

## Suggested commands

```bash
cd /Users/kosta/LocalDev/Word-MCP-Bridge

pnpm exec office-bridge list
pnpm exec office-bridge summary <session-id>
pnpm exec office-bridge events <session-id>

officecli view "docs/NIH_Biocoating_CraftV6 source-of-truth.docx" outline
officecli view "docs/NIH_Biocoating_CraftV6 source-of-truth.docx" text --max-lines 120
officecli query "docs/NIH_Biocoating_CraftV6 source-of-truth.docx" "35–200 kDa" --json
```
