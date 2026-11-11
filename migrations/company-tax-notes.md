# Company tax forecast

Apply `company-tax.json` once to the current workbook, deploy `CompanyTax.js` and
`RepaymentCalculations.js`, and run **SOL → Recalculate Loan Repayments**.
The migration clears the old tax/depreciation output formulas; the calculation
must populate the replacement outputs immediately. Do not replay the migration.

The only new global parameter is `financial year-end month`, set to 12 (December)
as confirmed by the user. Existing company tax rate and Units useful lives remain
unchanged. No BOI parameters or benefits are added.

## Calculation

- Depreciation uses the cost schedule matched to each Units type and that type's
  own useful life in column P. Each completion-month cohort is depreciated
  straight-line, starting in its completion month and stopping at its cost.
  Zero useful life excludes an asset from depreciation.
- Tax profit uses operating profit plus outgoing accrued interest, less unit
  depreciation, the existing setup depreciation input, and incoming accrued
  interest. Principal movements are not tax expenses or revenue.
- Losses are finalized annually, used oldest first, and expire after the fifth
  subsequent accounting period. Monthly taxable-profit movements can reverse
  within a year; this does not produce a cash refund.
- A full forecast accounting year prepays 50% of its forecast annual tax in
  accounting month eight. The balance falls in the month containing the 150th
  day after year-end. December year-end produces August and May payments.
  The annual summary retains settlements beyond the cash-flow horizon.
- Tax payments affect debt service, which affects accrued interest and tax.
  Repeated calculations in memory resolve this feedback to below 0.000001 THB,
  with a hard iteration limit and an error before any writes on nonconvergence.
  Saved tax payments and previous repayment outputs are never calculation inputs.
- The existing menu action saves incoming repayments, construction depreciation,
  monthly tax payments, loss balances, taxable-profit movements, tax expense,
  and the Company Tax annual summary together in one run. Other loan formulas
  and cash-flow array formulas remain formulas.

## Forecast assumptions and remaining classifications

This is a forecast under the workbook's existing accounting scope, not a filed
tax computation. Construction plan months represent completion/available-for-use
months. Opening losses and credits before the forecast are zero. Boundary periods
shorter than 12 months are explicitly identified and have no half-year prepayment.
Tax payments assume the full-year forecast is the estimate used for prepayment;
no payment extensions, holiday adjustments, withholding credits, or refunds are
modeled.

Setup depreciation remains the existing input (currently zero); no unsupported
asset classifications for infrastructure, furniture, vehicles, or setup costs
are invented. The forecast currently deducts accrued incoming interest; allocation
of borrowing costs to qualifying capital assets, entity-specific income and
expenses, and other tax adjustments still need supporting accounting data.

Official sources checked on 11 November 2026:

- [Revenue Department: Corporate Income Tax](https://www.rd.go.th/english/6044.html)
- [Revenue Code, sections 65–76](https://www.rd.go.th/english/37764.html)

## Verification

Business tests cover depreciation cohorts and lifetime caps, annual losses and
expiry, tax-accrual reversals, non-December years, partial periods, payment dates,
accrued rather than paid interest, calculation convergence, and write scope.
The live 96-month forecast was independently checked against monthly loan
accruals, repayment allocation, depreciation, annual tax, payment dates, and cash
rollforwards. All 11,156 numerical checks passed after the live script completed.
