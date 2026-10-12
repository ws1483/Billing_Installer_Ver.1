# Billing_Installer_Ver.1

## Payment Options (VBA source import required)

These exports add configurable aligner payment plans: one permanent full-total
invoice with a deposit request and monthly installment requests, filtered statements, deposit-paid release,
catch-up reminders and Outlook Classic email. **They are not active in your
workbook until imported.** Back up the `.xlsm` and try the import on a copy first.

### Import and Menu wiring

1. In Excel press **Alt+F11**. Remove/replace the existing modules named
   `modConfig`, `modHelpers`, `modConvert`, `modPayment`, `modStatement`, `modExportPDF`,
   `modMenu`, `modRenumber` and `ModMenuButtons` with the corresponding exports in this repo
   (do not keep duplicate modules). Keep the workbook's other existing modules,
   including its `NewInvoice`/`SaveInvoice` routines. Replace the `frmPayment`
   and `frmPaymentManage` form exports too (keep each matching `.frx` beside
   its `.frm`) for the payment-plan reference prompts.
2. Use **VBE > File > Import File** for `modPaymentOptions.bas` and
   `frmPaymentOptions.frm`. Keep `frmPaymentOptions.frx` next to the `.frm`
   during import; do not import the `.frx` separately. The form reuses the
   existing statement-form container and builds its Payment Options controls
   when opened. No Outlook project reference is needed.
3. Do **not** import `ThisWorkbook.cls` as a new class. Copy its
   `ShowInstallmentReminder` call into the workbook's existing `Workbook_Open`
   after the Menu refresh, preserving any other workbook event code.
4. On `Menu`, insert shapes/buttons and use **Assign Macro**:
   `BtnPaymentOptions`, `BtnDepositsStatement`, `BtnInstallmentsStatement`,
   and optionally `BtnPaymentPlanStatement`. `RefreshMenuSummary` adds a
   Payment Options summary at **B70**; reserve this cell.
5. Run **Debug > Compile VBAProject**, save as `.xlsm`, and enable macros only
   for the trusted workbook. Reopen to exercise the reminder.

### Alternative: build a clean Payment Options form

To avoid the reused `.frx` container, use `modBuildFormPaymentOptions.bas`
instead of importing the form in step 2 above. All other module imports and
Menu wiring still apply.

1. Work on a backed-up workbook copy. If `frmPaymentOptions` already exists,
   right-click it in the VBE and **Export File** to keep a backup, then remove
   it. Do not remove `frmStatement` or any other form.
2. Import `modBuildFormPaymentOptions.bas`. Temporarily enable **Trust access
   to the VBA project object model** in Excel's Trust Center > Macro Settings.
   The VBA project must be unlocked.
   If updating an older builder, remove the existing `modBuildFormPaymentOptions`
   module first, then import the updated file; do not keep duplicate builder modules.
3. Run `BuildFormPaymentOptions` and select the **updated, trusted**
   `frmPaymentOptions.frm` from this repository. No `.frx` is needed. The builder
   creates an empty designer and loads the form's code, including the `m_po*`
   handles and matching event handlers. It does not duplicate that code in the
   builder, so future changes to the selected export are included automatically.
   Existing components are never overwritten.
4. Disable Trust access again when finished. Run **Debug > Compile VBAProject**
   and save the workbook as `.xlsm`. An empty designer is expected: the controls
   are created when the form opens.

On a workbook copy, verify that cancelling the file picker or selecting another
form leaves the project unchanged; rerunning with an existing `frmPaymentOptions`
also leaves it unchanged. Open Payment Options through the Menu and check source
selection, preview changes and buttons, then exercise the reminder/mail flow using
the acceptance checks below. These checks require Excel/VBE.

The builder sets `StartUpPosition` through the VBA component's `Properties`
collection, not its MSForms designer (which can raise error 438 for that member).
If building still fails, report the error number and operation shown in the
dialog. Trust access is needed for project access; enabling it does not fix an
unsupported-property error.

If the builder fails, bypass it: in the VBE export/remove only the old
`frmPaymentOptions`, then **Insert > UserForm** in the billing workbook project.
Set **(Name)** to `frmPaymentOptions`, **Caption** to `Payment Options`, and
**StartUpPosition** to `1 - CenterOwner` in the Properties window. Leave the
designer empty. Open the updated `.frm` in a text editor and copy from
`Option Explicit` through the end into the form's **View Code** window; do not
paste the `VERSION`, designer or `Attribute` export headers. Compile, save and
run `BtnPaymentOptions`. This manual path needs no programmatic project access.

### Settings and schedule

The following key/value rows are added automatically to `Settings` columns A/B
below the existing counters (row 30 or later). Values must be numeric:

| Key | Default | Meaning |
| --- | --- | --- |
| `DepositPercent` | `0.5` | Fraction of total; 0 through 1 (50% is also valid as an Excel percentage) |
| `WeeksPerAligner` | `2` | Positive weeks worn per aligner |
| `WeeksPerMonth` | `4.345` | Positive treatment-week divisor |
| `ReminderLeadDays` | `7` | Nonnegative days before installment due date |

Enter the deposit percentage separately for each case in Payment Options.
The Settings percentage is only the default; a per-case entry does not overwrite
it. Enter `25` for 25% (not `0.25`). Type part of a saved quote/invoice number and select the matching document.
Choose one output folder for the deposit and all installment request PDFs.
After creating a plan, accept the offer to export all notices, or use **Export
all requests** on the saved plan later. **Export request** saves only the selected
row. All notices can be exported in advance (including Pending or zero-balance
rows), but only eligible unpaid released requests can be emailed.

Months are VBA `Round(aligners * WeeksPerAligner / WeeksPerMonth)`, minimum 1.
Amounts are VAT-inclusive, rounded to cents; the final month absorbs the
remainder. Tiny balances may have zero-valued installments, never negative ones.
R20,000 and 10 aligners at the defaults gives R10,000 deposit and five R2,000
installments. The deposit is due on the selected start date; monthly due dates
use `DateAdd("m", installmentNo, startDate)`.

`PaymentPlans` is auto-created with one row for the deposit (`InstallmentNo=0`)
and one per month. Each row refers to the **same permanent financial invoice**,
not to a separate installment invoice. Column constants and the expected header
layout live in `modConfig` and `EnsurePaymentPlansSheet`. Do not rearrange columns.
This model is for new plans; an incompatible legacy plan layout must not be
silently reinterpreted.

If dashboard, reminders and plan creation report **Incompatible PaymentPlans
layout**, and you have **never created any plans**, import the updated
`modPaymentOptions.bas`, then run **Alt+F8 > RecoverEmptyPaymentPlans** on a saved,
writable workbook copy. The macro checks all cells below the headers (including
hidden/filtered rows, formulas and data outside the PlanID column). If any data
exists, it stops without changing sheets: those records need reconciliation.
After confirmation, choose a short writable **local folder** (for example Documents)
for the backup, even if the workbook is stored in OneDrive / SharePoint. A web
address is not a filesystem backup destination. Cancelling the picker makes no
changes. The macro saves a uniquely named `PaymentPlansRecovery_<timestamp>` copy
with the workbook's original extension and verifies it exists and is nonempty
before renaming the empty sheet to `PaymentPlans_Backup` without deleting it, then
creates the new layout. An existing backup sheet is never overwritten. Reopen
Payment Options, select the saved unconverted quote and create the plan.
InvoiceLog, QuoteLog and Payments are not changed by recovery. Save after checking.
If creation fails after the rename, keep the archived sheet and workbook backup
and inspect the partial recovery before retrying. Do not manually set legacy
SchemaVersion values to 2 or remove validation guards.
For **Bad file name or number**, replace the old module with this updated export
and choose a short local folder. If it still fails, report the error number and
stage in the dialog, with personal path details redacted. A backup failure never
proceeds to renaming. On a workbook copy, verify cancellation, an unwritable
destination and an existing filename leave the original sheet intact or allocate
a unique filename; verify a successful backup opens before using the fresh layout.

The first 22 columns retain their existing positions; `RequestRef`,
`SchemaVersion`, `SendState` and `DepositPercent` are appended. New schedule
rows use schema version 2. Receipt metadata columns `PaymentPlanID` and
`PaymentRequestRef` are appended by header to `Payments`, without overwriting
its bank reference, notes or existing extra columns.

Select a saved invoice to retain its number, original details, tax, line items
and total. Select a saved quote to create **one full-total invoice** using the
normal invoice-number sequence and link the quote to it. A manual case creates
one full-total invoice as well. A payment request is a schedule notice, **not an
additional tax invoice or charge**.

For `INV-WA-00123`, the plan is `PP-WA-00123`, the deposit request is
`PP-WA-00123-DEP`, and installments are `PP-WA-00123-01`, `-02`, etc. All request
PDFs show the permanent invoice number. PDFs use the existing document/patient
filename style, with the request reference as document number; existing files
must not be silently overwritten.

Cancelling a plan cancels its schedule, **not the permanent invoice**. The
financial debt remains until payment or the normal credit/cancellation workflow.
An eligible amendment retains the invoice and old schedule history and gets a
new revision reference. Issued numbers and request references are never reused.
Requests already sent or payments already recorded can block automatic schedule
changes; use the practice's reviewed correction process rather than altering history.
Exported requests and unconfirmed export/send attempts also block automatic
amendment/cancellation. Revisions use `PP-WA-00123-R02`, then `-R03`, etc.
Existing medical invoices remain in `MedAidLog` with `PP-MC-...` plan references.

### Statements and mail

Deposits statements select recipient (doctor/private/med-claim), department and
account. The installments statement combines payable installments across plans;
the single-plan statement includes the deposit and every month, with paid status.
Payment Options statements show request amounts and allocated payments, not the
full invoice total repeated on every schedule row. They exclude account-level
credit because it cannot safely be attributed to one plan.

Normal monthly statements continue to show the original invoice and its existing
details once, with its full total. Actual deposit/installment payments are recorded
against that permanent invoice in `Payments`; they reduce its balance according
to the statement cutoff date. Request references are separate allocation metadata,
not replacement invoice numbers and not bank-reference replacements. There are
no additional installment invoice charges on normal statements.
In Record Payment, select the permanent invoice as usual. For an active plan,
enter an optional request reference and the actual amount received. Blank request
references allocate deposit-first, then oldest installment; an explicit reference
allocates to that request first, with any excess applied once to remaining
schedule rows. Manage Payments edits/deletions rebuild allocations. Future-dated
receipts do not release requests early; normal statements use the payment cutoff.

Installments remain pending until the deposit is paid. Payment reconciliation
and reminder scans update plan status; an unpaid/reversed deposit suspends
release. The reminder includes **all** unpaid, unsent released installments
due within the lead window, including older missed requests. Select a request
and click **Mail now**. The email is late-bound through Office 365 **Classic
Outlook**, with a clearly labelled nonfinancial payment-request PDF attachment.
Review the recipient and confirm sending; the mailed date is recorded only
after Outlook accepts the send. Outlook delivery is not guaranteed by that
timestamp. Missing email addresses, Outlook, or PDF export failures leave the
request eligible for retry; no external scheduler or automatic background send.
An actual Outlook send attempt is recorded conservatively: an unsuccessful or
unconfirmed send requires reviewing Outlook and AuditLog before any resend, to
avoid sending duplicate requests.

### Acceptance checks in Excel

Optionally import `modPaymentOptionsTests.bas` and run `TestPaymentOptions` in
the VBE. It uses explicit calculation settings and `Debug.Assert` for the worked
example, rounding, minimum term, tiny balances, 0%/100% deposits, request numbering,
revision IDs, receipt allocation, saved quote headers and invalid inputs.
No Excel/VBA runtime is included in this source-only repository, so these
assertions and the following integration checks must be run on a workbook copy:

- Use an existing R20,000 invoice / 10 aligners: preserve its number, all original
  lines, tax and total. Create six request references, not six new InvoiceLog
  charges. Check 25%, 0% and 100% case deposits without changing Settings.
- Select a quote: create exactly one full-total invoice, copy its saved details
  and lines, and verify both quote/invoice links. The new invoice has its actual
  issuance date, not the old quote date; the quote date is retained on the source.
  Retry and check duplicate prevention.
- Search using the middle/end of a quote/invoice number, select one match, and
  verify unmatched/ambiguous text is not used as a saved source.
- Select a folder, inspect all request PDFs for request number, original invoice,
  recipient, amount and due date, and check filename collision/failure handling.
  Simulate a PDF export failure: retain the recorded invoice and schedule, do not
  generate another financial invoice merely to retry exporting its requests.
  Export all notices before paying the deposit; ensure installment PDF saving
  works without releasing those installments for email.
- Record/reverse the deposit payment and verify installment release/suspension,
  separate request references and the original invoice's paid/balance totals.
- For R20,000 less R10,000 deposit paid and R2,000 installment paid, verify the
  normal statement has one R20,000 invoice charge and R8,000 remaining. Move the
  installment payment past the statement cutoff: only R10,000 should be deducted.
  Verify doctor, patient and medical recipients, and existing non-plan invoices.
- Backdate several installments; reopen and verify catch-up, failed-mail retry,
  successful-mail removal and the Menu count.
- Generate detailed plan statements, inspect PDF totals/aging and ensure each
  payment is allocated once, including partial and unallocated payments. Check
  zero-rated/mixed-tax sources without duplicating or inventing tax charges.
- Cancel/amend an unpaid unsent plan: preserve the original invoice and retained
  schedule/audit history, use new revision references and verify paid/sent
  schedules block changes. Cancel then recreate: never reuse old request IDs.
- Verify payment edits/deletions refresh both the original invoice and schedule,
  future-dated payments do not release requests early, and changing a request
  reference cannot move money to another patient's invoice.
- Test duplicate source, converted quote, numbering collision, missing Outlook,
  missing email, and PDF failure on the workbook copy before production use.