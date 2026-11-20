# Apps Script synchronization

After making changes to the Apps Script source in this project, verify the changes and use `clasp push` to synchronize them to the linked Apps Script project as part of the task. The user authorizes this by default; do not ask for confirmation solely to push.

Before pushing, check for remote changes and unrelated local changes so they are not accidentally overwritten or published. Verify the push where possible. If network access, authentication, permissions, or another blocker prevents a safe push, report that clearly and state that the changes are not live.

# Spreadsheet formulas

- When writing or modifying formulas, first look for and reuse existing named formulas/functions and named ranges wherever practical. Minimize direct cell and range references; use them only when existing names do not adequately support the calculation.
- Preserve existing named-function calls. Do not replace them with direct references or inlined formulas without the user's explicit approval for that replacement.
