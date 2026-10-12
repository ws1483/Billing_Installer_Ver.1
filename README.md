# Billing_Installer_Ver.1

## Payment Options setup (Excel desktop)

Work on a backup of the macro-enabled workbook.

1. Enable **File > Options > Trust Center > Trust Center Settings > Macro Settings >
   Trust access to the VBA project object model**.
2. In the VBE (Alt+F11), import `modPaymentOptions.bas` and
   `modBuildPaymentOptionsForm.bas`. Replace the existing modules with the updated
   `modConfig.bas`, `modConvert.bas`, `modExportPDF.bas`, `modHelpers.bas`,
   `modPayment.bas`, `modRenumber.bas`, `modStatement.bas`, `modMenu.bas`, and
   `ModMenuButtons.bas` from this repository. Do not import duplicates.
   These support modules are required by the payment engine. Merge the
   `ShowInstallmentReminder` call from `ThisWorkbook.cls` into your existing
   `Workbook_Open` event, after `RefreshMenuSummary` (do not import a second
   ThisWorkbook class).
3. Select this workbook's project in the VBE. Run **BuildPaymentOptionsForm**
   once from Alt+F8. It creates the form and its named controls in the designer,
   then injects only VBA code starting with `Option Explicit`. Save the workbook
   to persist the form. No `frmPaymentOptions.frm` or `.frx` import is used.
   Close the form before rebuilding. Re-running asks before replacing it;
   choosing No leaves it unchanged. Replacement discards custom form edits.
   Locked projects must be unlocked first.
4. Assign the Menu button to **BtnPaymentOptions**. The wrapper resolves
   `frmPaymentOptions` by name and calls `.Show`, avoiding an undefined-form
   compile error before the first build. After building, `frmPaymentOptions.Show`
   also works directly from the Immediate window.
   Optional statement buttons: `BtnDepositsStatement`,
   `BtnInstallmentsStatement`, `BtnPaymentPlanStatement`.
   `BtnSavePaymentPlanPDFs` prompts for a Plan ID to save/retry invoice PDFs.
5. Choose **Debug > Compile VBAProject**, then open Payment Options. After the
   form is saved, trusted project access is no longer needed for ordinary use
   and can be disabled again.

### Style and layout

The builder matches `frmPayment`: teal `BackColor = &H00B0872E&`,
`ForeColor = &H80000012&`, Tahoma 9, white bold section labels, white inputs,
and system-grey buttons (`&H8000000F&`). It uses a borderless, flat,
approximately 680 x 450 point form, `Cycle = 0` (AllForms),
`KeepScrollBarsVisible = 3`, `StartUpPosition = 1` (CenterOwner), and
`ShowModal = True`. Tweak these values and control positions in
`modBuildPaymentOptionsForm.bas`, then rebuild.

Recipient/department are at the top-left, source/customer at the top-right,
with aligners, total, deposit date and plan-ID lookup above a read-only plan
summary and scrollable schedule preview. Buttons use two bottom rows to keep
all captions readable. The recipient choices Doctor / Private Patient / Med
Claim map to the engine's `doctor` / `patient` / `medclaim` values.
Departments include All/WA/WD/MC, matching the existing filter convention;
manual invoice generation requires WA or WD.

### Using payment plans

Type part of a document number, patient name or customer ID in **Source
Quote/Invoice**. Matching quotes, invoices and med claims appear in a
three-column dropdown (document / patient / customer); matching is
case-insensitive and does not require the complete number. Select a result
before generating (a single remaining match is resolved automatically).
Leave the source blank for manual entry. Enter the aligner count (a positive
whole number) and **Deposit %** (enter `30` for 30%, allowed range 0-100).
The field starts with the Settings default, but changing it applies only to
this plan, including generation and amendment; it does not change the default.
Loading a plan restores its deposit percentage from the saved deposit/total.
The preview shows the
deposit and each monthly installment, with the final amount adjusted for cents.
Generation uses the existing invoice numbering and layout. New Plan IDs use
**Patient Name - Source Quote/Invoice Number**, with `(2)`, `(3)`, etc. added
for repeat/replacement IDs. Manual plans use **Patient Name - Manual**.
A patient name is therefore required for all new plans. Existing Plan IDs
remain valid and are not renamed. Enter a **Plan ID**
to load a plan for amendment, cancellation or a single-plan statement.
Patient name and deposit due date are retained as separate inputs.
Totals remain locked for sourced plans because the engine uses the source
document's total when amending; manual-plan totals remain editable.

After generation or amendment, choose whether to save the deposit and monthly
invoices as PDFs. A **Save As** dialog lets you select each file's location and
name, defaulting to invoice number plus patient name. Existing files are never
overwritten: choose a new name. Cancel stops further PDF exports without
undoing saved invoices or deleting PDFs already created. Use
**BtnSavePaymentPlanPDFs** from Alt+F8 to retry later. Saved paths are tracked in
`PaymentPlans`; saving a PDF does not send mail or mark a reminder as sent.
PDFs use the existing one-page invoice format and saved VAT allocation.

If upgrading from the earlier builder, replace the updated modules, close
Payment Options, and run **BuildPaymentOptionsForm** again (confirm replacement).
The builder now injects all code in a single operation, avoiding blank lines
between `_` continuation statements that caused the reported syntax error.

The `PaymentPlans` sheet is created automatically. Settings keys (column A key,
column B value, starting at row 30) are `DepositPercent` (default 0.5 = 50%),
`WeeksPerAligner` (2), `WeeksPerMonth` (4.345), and `ReminderLeadDays` (7).
Missing keys are added with defaults. Deposit payment activates the plan;
the Menu dashboard and workbook-open reminders track upcoming and overdue
installments. **Mail Installment** opens the reminder list: select an invoice
and mail it using Outlook Classic. Deposits, combined installments and
single-plan statements use the existing statement engine. Cancellation and
amendment retain the audit trail and protect paid/sent invoices.

### Verification on a workbook copy

Optionally import `modPaymentOptionsTests.bas`. Run `TestPaymentOptions` for
calculation assertions and `TestPaymentOptionsForm` after building for controls,
style, initialization, code-behind and live preview assertions.
`TestPaymentOptionsCode` checks generated continuation-line adjacency without
requiring trusted access or a built form.

Also verify these manual cases in Excel:

- With trusted project access disabled, the builder explains the setting and
  exits without creating a component.
- Build with trust enabled; open via the Menu and via `frmPaymentOptions.Show`.
- Close and rebuild: choose No to retain the form, then Yes to replace it.
  Confirm that only one `frmPaymentOptions` exists.
- Select a source; confirm total, recipient, department and customer populate.
  Search using partial document numbers and patient names; confirm typing
  does not trigger exact-lookup errors, and ambiguous/unmatched text cannot
  generate invoices. Change aligners/total/date/deposit percentage (including
  0, 30, 100 and invalid values) and check the preview. Generate invoices on the
  copy, then load the resulting Plan ID to test statements/amend/cancel.
- Confirm new IDs contain patient/source and amendments add a unique suffix.
  Check that deposits match the chosen percentage. Test Save As location/name,
  cancellation, existing-file protection and retry, and verify exported
  subtotal/VAT/total against InvoiceLog. Confirm template formulas/events are
  restored after successful and failed export.
- With unpaid, active installments due, open the reminder, select an invoice,
  and test Outlook mailing. Test missing logs and unavailable Outlook for
  helpful messages rather than unhandled errors.

Excel/VBE and Outlook runtime checks require Windows desktop Office; they
cannot be executed in the repository's Linux environment.