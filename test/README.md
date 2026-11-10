# In Loans calculation control

The bound spreadsheet exposes **SOL → In Loans calculations → Freeze / Unfreeze / Show status**.
The **In Loans!A1** dropdown also accepts **Freeze** and **Unfreeze**. The existing
`onInstallableEdit` trigger handles it and displays a toast. Menu operations and failed
dropdown operations synchronize A1 to the next available action: **Freeze** while live,
and **Unfreeze** while frozen. The Date header in C3
remains unchanged. Change the dropdown on its own and wait for the completion toast
before choosing another action.
Reload the spreadsheet after `clasp push` to refresh its existing installable menu trigger.

Freeze saves the formulas and array results as literal JSON in the hidden
`_in_loans_freeze_backup` sheet, then replaces calculated cells with their current
values. Backup and cell changes use one atomic Sheets API batch. Inputs, formatting,
validation, and other sheets are untouched. Unfreeze clears the saved spill results
and restores formulas, retaining changes to existing input cells.

While frozen, edit existing input cells only. Do not sort, move, insert, or delete
rows/columns, alter frozen results, or edit the backup sheet. Unfreeze detects changed
dimensions and modified results, but cannot reliably detect every same-size structural
rearrangement. Do not edit the sheet while a freeze/unfreeze command is running;
the document lock serializes script commands, not edits made by collaborators.

Freeze refuses formula errors, including loading errors. Unfreeze restores formulas;
the spreadsheet's existing automatic/manual calculation setting still applies.

Run local tests from the repository root:

```sh
node --test test/in_loans_calculations.test.js
```

For a live integration test, temporarily copy `test/apps_script_integration.js` into
`src/InLoansValidation.js`, push with clasp, and run `validateInLoansFreezeIntegration`
in the Apps Script editor. It creates a separate validation spreadsheet and tests
two freeze/unfreeze cycles, input preservation, array results, literal strings,
booleans, repeated freeze, and conflict detection. The execution log records its URL.
Remove the temporary source file and push again after testing.
Run `validateInLoansDropdownIntegration` in the same temporary test file to verify
dropdown actions, menu synchronization, input preservation, and failure recovery.
