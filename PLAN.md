# Billing System — Modernization Plan

> Working document. Derived from a full scan of the existing Excel/VBA application.
> Everything here is a proposal to be fleshed out and confirmed. Nothing is final.

---

## 1. What the current program is (from the code scan)

The existing app is an **Excel `.xlsm` VBA application** for a dental / medical-aid
billing practice (appears to be **South African** — WA/WD departments, medical-aid
claims, BHF numbers, Rand formatting). It is genuinely well-built for VBA:
clean module separation, a single central column-map, normalized ID matching,
overwrite guards, and an audit log.

### 1.1 Document types (all implemented today)
| Kind | Sheet | Log sheet | Lines sheet | Number format | Counter cell |
|------|-------|-----------|-------------|---------------|--------------|
| Invoice | `Invoice` | `InvoiceLog` | `InvoiceLines` | `INV-<WA/WD>-0000` | Settings B12/B13 |
| Quote | `Quote` | `QuoteLog` | `QuoteLines` | `Q-<WA/WD>-0000` | Settings B10/B11 |
| Med Claim | `Med Claim` | `MedAidLog` | `MedAidLines` | `MC-INV-<WA/WD>-0000` | Settings B20 |
| Credit Note | `CreditNote` | `CreditNotes` | `CreditNoteLines` | `CN-0000` | Settings B14 |
| Payments | — | `Payments` | — | `PAY-0000` | Settings B18 |

Supporting sheets: `Settings`, `Customers` (doctors), `Med Customers` (patients),
`Pricelist`, `AuditLog`, `PatientCredits` (auto-created), `Menu`.

### 1.2 Module map (what each `.bas`/form does)
- **modConfig** — the backbone. Column constants for every sheet + a `DocConfig`
  type + `GetDocConfig()` that unifies the 4 doc kinds. **This is the single best
  asset for the rewrite — it is effectively the data schema.**
- **modHelpers** — shared helpers: `NrmID` (normalize id), `NextDocNumber` /
  `NextUniqueDocNumber` (counter-based numbering with uniqueness loop),
  `FindLogRow`, `SafeToWriteNumber` (overwrite guard), `LogAudit`,
  customer/doctor lookups, doctor & patient credit stores.
- **modMenu** — builds the dashboard by writing summaries *into Menu cells*
  (B19:H68). Filters read from cells: dept `C5`, date-from `B17`, date-to `C17`,
  doctor `E17`, recipient `I17`, type `G17`. **This is the "menu = summary"
  coupling you want to break.**
- **modMenuNew / ModMenuButtons / modMenuSave** — button wrappers for New/Save.
- **modMedClaim** — Save/Recall med claims + autofill from `Med Customers`.
- **modRecall** — Recall invoice/quote/credit-note (resets sheet, applies
  recipient layout, loads lines).
- **modRecipient** — dynamic Doctor vs Private layout switching (rewrites labels
  + INDEX/MATCH formulas per recipient type).
- **modConvert** — Quote → Invoice conversion, links both logs, exports PDF.
- **modRenumber** — dept-change renumber, manual rename, revert-converted-quote.
  Cascades number changes across Payments + CreditNotes. Unpaid-only guards.
- **modCreditLink** — Credit Note ↔ Invoice linking; applies credit against
  invoice, routes **excess** into doctor/patient credit store; full void/reverse.
  (Most financially sensitive logic in the app.)
- **modPayment** — record/manage payments; `ReconcileDoc` recomputes
  Paid/Balance/Status from Payments rows.
- **modFormulas** — line price auto-fill from `Pricelist` (G = price incl,
  authoritative; E/F/H derive).
- **modExportPDF** — export/print active doc; forces one-page fit
  (`$A$1:$H$48`, FitToPagesTall=1).
- **modBackup** — timestamped `SaveCopyAs`, keeps newest 3; navigation helpers.
- **modCleanup / modNameAudit / modNameRepair** — integrity sweeps for orphan
  line rows and number consistency.
- **modCalendar / frmCalendar** — date picker.
- **Forms**: `frmPayment`, `frmPaymentManage`, `frmReport`, `frmSearch`,
  `frmStatement`, `frmCalendar`.
- **ThisWorkbook** — on open, activates Menu + refreshes summary.

### 1.3 Report types available today (`frmReport`)
VAT, Invoices, Quotes, Med Claims, Credit Notes, Full Summary — filtered by
date range + department (All/WA/WD/MC), Summary or Full mode.

---

## 2. Why the requested features need a rewrite (honest assessment)

Every feature you asked for is blocked by the **Excel-as-database, single-file,
single-user** architecture:

| Requested feature | Blocker in current design |
|---|---|
| Licensing / standalone program | It's an `.xlsm` needing Excel; no licensing possible |
| Multi-PC, central DB, "work from anywhere" | The workbook *is* the database; one file = one user at a time |
| Login with different users | Only `Environ$("USERNAME")` is captured, no auth |
| Menu shows options only; summaries in separate windows | Menu sheet **is** the summary (cells B19:H68) |
| VAT on/off, currency settings | VAT is baked into sheet formulas + fixed columns |
| Fully customizable page setups + headings/footers | Layout is hardcoded cell ranges (`$A$1:$H$48`) |
| CRM + Dental Lab modules | No module framework; data model is fixed sheets |
| Export to accountant (CSV/others) | No export layer today |

**Recommendation: rebuild as a C#/.NET desktop app on a real database**, reusing
the existing VBA as the *authoritative behavior spec*. The `modConfig` schema,
numbering rules, credit-note excess handling, and reconciliation logic all
transfer directly.

---

## 3. Recommended technology stack (for confirmation)

| Concern | Recommendation | Rationale |
|---|---|---|
| Language / UI | **C# + .NET 8, WPF** (WinForms acceptable) | Native Windows, mature, closest to VBA logic |
| ORM / data | **Entity Framework Core** | Maps `*Log`/`*Lines` sheets to tables cleanly |
| Database | **PostgreSQL** (central), **SQLite** local/offline option | True multi-user client/server |
| "From anywhere" | Cloud-hosted DB or on-prem server + VPN | Real remote access |
| PDF / documents | **QuestPDF** with user-editable templates | Customizable page setup + headings/footers |
| Reporting | LINQ queries + grid/export | Replaces cell-driven filters |
| Auth | Users + roles table, hashed passwords (BCrypt) | Multi-user login + per-user audit |
| Licensing | Signed license key + machine activation, offline grace | Your "standalone with licensing" need |
| Installer | MSI (WiX) or Inno Setup | Distribution |

> If staying in Excel/VBA is a hard requirement, licensing + central multi-user DB
> + separate summary windows are all severely limited. Please confirm the stack —
> this is the single biggest decision.

---

## 4. Proposed database schema (derived from `modConfig`)

Direct translation of the sheet column maps:

- **users** (id, username, password_hash, role, is_active, created_at)
- **settings** (key, value) — VAT enabled, VAT %, currency, counters, backup dir,
  bank details, headings/footers text
- **doctors** (was `Customers`: cust_id, name, street, suburb, city, postcode,
  id_no, tel, bhf_no, med_aid, vat_no, credit_balance)
- **patients** (was `Med Customers`: bill_to_key, name, address…, id_no, tel,
  email, med_aid_name, med_aid_number, dependant_code, main_member, credit_balance,
  **folder_path** ← for the "open folder" feature)
- **documents** (unified header for INV/QTE/MC/CN: id, doc_kind, doc_no, dept,
  recipient_type, doc_date, due_date, patient_id, doctor_id, appliance_type,
  subtotal, discount, vat, total, paid, balance, status, source_quote_no,
  converted_inv_no, source_inv_no, med_aid fields, notes, created_at, modified_at,
  created_by)
- **document_lines** (doc_id, line_no, code, zcode, description, qty, excl, vat,
  incl, total)
- **payments** (id, doc_id, date, amount, method, reference, notes, created_by)
- **audit_log** (timestamp, user, action, doc_no, old_val, new_val, comment)
- **pricelist** (code, description, zcode, price_incl, vatable)
- **document_templates** (doc_kind, page_setup json, header, footer, logo)
- **crm_*** and **lab_*** tables (see §6)

---

## 5. Feature-by-feature delivery plan

### 5.1 Core (Phase 1)
- PostgreSQL + EF Core schema above
- Users, login, roles, per-user audit (extends `LogAudit`)
- Settings screen: **VAT on/off + VAT %**, **currency**, bank details, counters

### 5.2 Documents (Phase 2)
- Port Invoice/Quote/Med Claim/Credit Note using existing rules:
  - Numbering + uniqueness loop (`NextUniqueDocNumber`)
  - Overwrite guard (`SafeToWriteNumber`)
  - Recipient (Doctor/Private) dynamic fields
  - Quote→Invoice conversion
  - Dept-change renumber + manual rename + revert
  - **Credit-note excess → credit store** (preserve exactly)
  - Payment reconciliation (`ReconcileDoc`)

### 5.3 Customizable templates (Phase 3)
- Per-doc-kind page setup, margins, orientation
- Editable **headings & footers**, logo, banking block
- Replaces the hardcoded `$A$1:$H$48` one-page fit

### 5.4 New Menu + summaries (Phase 4)
- **Menu = options only** (no data)
- Summaries open in **separate windows**
- **Open patient folder** button (opens the OS folder from `patients.folder_path`)

### 5.5 Reports + filters (Phase 5)
Existing filters: date range, department, doc type.
**Suggested additions:** status (Unpaid/Part-Paid/Paid), amount range,
**aging buckets 30/60/90**, medical-aid name, patient, doctor, balance>0 only,
created-by user, recipient type. Results in a sortable grid, exportable.

### 5.6 Accountant export (Phase 6)
- Universal **CSV** with column mapping
- Presets: generic ledger, **QuickBooks (IIF/CSV)**, **Xero CSV**, **Sage CSV**
- Date-range + doc-type scoped

### 5.7 CRM module (Phase 7)
- Unified contacts (patients + doctors as parties)
- Leads, follow-ups, reminders, communication log
- **Sees invoices/quotes** for a contact (read from `documents`)

### 5.8 Dental Lab module (Phase 8)
- Lab jobs/orders with status workflow: received → in-progress → ready → delivered
- Linked to patient + doctor
- Can auto-generate a Quote or Invoice via the existing document engine

### 5.9 Licensing, installer, remote hardening (Phase 9)
- Signed license keys, machine activation, offline grace period
- MSI/Inno installer
- Connection pooling, migrations, backup/restore for the central DB

---

## 6. Recommended improvements (found during the scan)

1. **Replace bubble sorts** (`modMenu`, `modBackup`) — fine in VBA, but in the
   rewrite use DB `ORDER BY` (menu summaries currently bubble-sort in memory).
2. **Kill duplicated helpers** — `Num`/`NrmID` are redefined in `modCreditLink`
   though they exist in `modHelpers`. Consolidate in the new codebase.
3. **Concurrency safety on counters** — counter-cell increment is not safe for
   multi-user. Use a DB sequence / transaction so two PCs can't grab the same
   number. (Your `NextUniqueDocNumber` loop is a good fallback but not atomic.)
4. **Move VAT out of formulas** — make it a setting so VAT/no-VAT and rate are
   configurable; recompute line excl/vat/incl in code.
5. **Foreign keys instead of string matching** — today docs match doctors/patients
   by normalized name/id strings; use real FK ids to avoid rename drift (already
   partly mitigated by the cascade repointing in `modRenumber`).
6. **Full audit** — capture real logged-in user (not just Windows username),
   before/after values on every write.
7. **Soft-delete / void trail** — the hard deletes in `RevertQuoteToSaved` and
   `VoidCreditNote` are well-guarded, but a soft-delete + audit trail is safer
   for financial records.
8. **Line-item cap** — current sheets cap at 15 lines (rows 16–30). The DB model
   removes that limit.
9. **Centralized error handling + logging** to a file/table instead of `MsgBox`.
10. **Automated tests** around numbering, credit-note excess, and reconciliation —
    the highest-risk money paths.

---

## 7. Open questions to confirm before building

1. **Stack:** OK with **C#/.NET + PostgreSQL**, or must it stay Excel/VBA?
2. **Users / PCs:** how many concurrent users and PCs?
3. **Hosting:** cloud-hosted DB, or office server + VPN?
4. **Country/VAT/currency:** confirm South Africa / ZAR / 15% VAT and med-aid
   claim format?
5. **Priority order:** is the phase order above right, or should CRM / Lab come
   earlier?
6. **Data migration:** do we need to import existing invoices/quotes/patients from
   the current workbook into the new DB?

---

## 8. Status

- [x] Scanned all existing VBA modules and forms
- [x] Documented current architecture + data model
- [x] Proposed stack, schema, phased plan, improvements
- [ ] Confirm open questions (§7)
- [ ] Lock scope for Phase 1
