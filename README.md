# Billing_Installer_Ver.1

## Payment Options (VBA source import required)

These exports add configurable aligner payment plans: a deposit invoice and one
invoice per treatment month, filtered statements, deposit-paid release,
catch-up reminders and Outlook Classic email. **They are not active in your
workbook until imported.** Back up the `.xlsm` and try the import on a copy first.

### Import and Menu wiring

1. In Excel press **Alt+F11**. Remove/replace the existing modules named
   `modConfig`, `modHelpers`, `modConvert`, `modPayment`, `modStatement`, `modExportPDF`,
   `modMenu`, `modRenumber` and `ModMenuButtons` with the corresponding exports in this repo
   (do not keep duplicate modules). Keep the workbook's other existing modules,
   including its `NewInvoice`/`SaveInvoice` routines.
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

### Settings and schedule

The following key/value rows are added automatically to `Settings` columns A/B
below the existing counters (row 30 or later). Values must be numeric:

| Key | Default | Meaning |
| --- | --- | --- |
| `DepositPercent` | `0.5` | Fraction of total; 0 through 1 (50% is also valid as an Excel percentage) |
| `WeeksPerAligner` | `2` | Positive weeks worn per aligner |
| `WeeksPerMonth` | `4.345` | Positive treatment-week divisor |
| `ReminderLeadDays` | `7` | Nonnegative days before installment due date |

Months are VBA `Round(aligners * WeeksPerAligner / WeeksPerMonth)`, minimum 1.
Amounts are VAT-inclusive, rounded to cents; the final month absorbs the
remainder. Tiny balances may have zero-valued installments, never negative ones.
R20,000 and 10 aligners at the defaults gives R10,000 deposit and five R2,000
installments. The deposit is due on the selected start date; monthly due dates
use `DateAdd("m", installmentNo, startDate)`.

`PaymentPlans` is auto-created with one row for the deposit (`InstallmentNo=0`)
and one per month. Its columns track plan/source IDs, recipient, customer,
department, original total/deposit, installment number/count, due date, amount,
invoice number/generated flag, status, mailed date, creation/modification time,
patient name and PDF path. Column constants live in `modConfig`.
The exported layout is: `PlanID, SourceDocNo, RecipientType, CustID, Dept,
Total, Deposit, InstallmentNo, InstallmentCount, DueDate, Amount, InvoiceNo,
Generated, Status, Reminded, CreatedAt, ModifiedAt, PatientName, PDFPath,
Aligners, Paid, Balance`. Do not rearrange these headers or edit generated
invoices directly.

Select a saved quote/invoice, or leave source blank and enter the total plus
recipient, customer/patient and department. A saved source's authoritative total
is used. Already converted quotes and duplicate live plans are blocked. Replacing
an unpaid source invoice preserves a void/audit trail rather than charging both
the original total and the schedule. Paid or sent plans cannot be cancelled or
amended automatically; use the practice's credit/payment correction process.
Amending an eligible plan replaces its unsent invoices using the current source
and aligner count. Financial rows are never hard-deleted by this module.

### Statements and mail

Deposits statements select recipient (doctor/private/med-claim), department and
account. The installments statement combines payable installments across plans;
the single-plan statement includes the deposit and every month, with paid status.
All use `StatementTpl`, the existing statement formatting/aging, PDF export and
`StatementLog`. They exclude account-level credit because it cannot safely be
attributed to one plan. These filtered reports are not a replacement for the
practice's full account statements.

Installments remain pending until the deposit is paid. Payment reconciliation
and reminder scans update plan status; an unpaid/reversed deposit suspends
release. The reminder includes **all** unpaid, unsent released installments
due within the lead window, including older missed invoices. Select an invoice
and click **Mail now**. The email is late-bound through Office 365 **Classic
Outlook**, with the saved invoice exported as a one-page PDF attachment.
Review the recipient and confirm sending; the mailed date is recorded only
after Outlook accepts the send. Outlook delivery is not guaranteed by that
timestamp. Missing email addresses, Outlook, or PDF export failures leave the
invoice eligible for retry; no external scheduler or automatic background send.

### Acceptance checks in Excel

Optionally import `modPaymentOptionsTests.bas` and run `TestPaymentOptions` in
the VBE. It uses explicit calculation settings and `Debug.Assert` for the worked
example, rounding, minimum term, tiny balances, 0%/100% deposits and invalid
inputs. No Excel/VBA runtime is included in this source-only repository, so these
assertions and the following integration checks must be run on a workbook copy:

- Generate a R20,000 / 10-aligner plan: six unique correctly numbered invoices,
  total R20,000, correct bill-to, patient, tax and due dates.
- Record/reverse the deposit payment and verify installment release/suspension.
- Backdate several installments; reopen and verify catch-up, failed-mail retry,
  successful-mail removal and the Menu count.
- Generate each statement, inspect PDF totals/aging and its `StatementLog` row;
  ensure unrelated account invoices are excluded.
- Cancel/amend an unpaid unsent plan, verify the retained audit/void trail,
  and verify paid/sent invoices block changes.
- Test duplicate source, converted quote, numbering collision, missing Outlook,
  missing email, and PDF failure on the workbook copy before production use.