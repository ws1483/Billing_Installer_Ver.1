# Billing System — Modernization Plan

> Working document. Derived from a full scan of the existing Excel/VBA application.
> **§7 decisions are now CONFIRMED.** Building against these.

---

## 1. What the current program is (from the code scan)

The existing app is an **Excel `.xlsm` VBA application** for a dental / medical-aid
billing practice (South African — WA/WD departments, medical-aid claims, BHF
numbers, Rand formatting). It is genuinely well-built for VBA: clean module
separation, a single central column-map, normalized ID matching, overwrite
guards, and an audit log.

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
  type + `GetDocConfig()` that unifies the 4 doc kinds. **The single best asset
  for the rewrite — effectively the data schema.**
- **modHelpers** — shared helpers: `NrmID`, `NextDocNumber` /
  `NextUniqueDocNumber` (counter numbering + uniqueness loop), `FindLogRow`,
  `SafeToWriteNumber` (overwrite guard), `LogAudit`, customer/doctor lookups,
  doctor & patient credit stores.
- **modMenu** — builds the dashboard by writing summaries *into Menu cells*
  (B19:H68); filters read from cells (dept C5, dates B17/C17, doctor E17,
  recipient I17, type G17). **The "menu = summary" coupling to break.**
- **modMenuNew / ModMenuButtons / modMenuSave** — New/Save button wrappers.
- **modMedClaim** — Save/Recall med claims + autofill from `Med Customers`.
- **modRecall** — Recall invoice/quote/credit-note.
- **modRecipient** — dynamic Doctor vs Private layout switching.
- **modConvert** — Quote → Invoice conversion, links both logs, exports PDF.
- **modRenumber** — dept-change renumber, manual rename, revert-converted-quote;
  cascades number changes across Payments + CreditNotes; unpaid-only guards.
- **modCreditLink** — Credit Note ↔ Invoice linking; applies credit, routes
  **excess** into doctor/patient credit store; full void/reverse. (Most
  financially sensitive logic.)
- **modPayment** — record/manage payments; `ReconcileDoc` recomputes
  Paid/Balance/Status from Payments rows.
- **modFormulas** — line price auto-fill from `Pricelist` (G incl authoritative).
- **modExportPDF** — export/print active doc; one-page fit (`$A$1:$H$48`).
- **modBackup** — timestamped `SaveCopyAs`, keeps newest 3; navigation helpers.
- **modCleanup / modNameAudit / modNameRepair** — integrity sweeps.
- **modCalendar / frmCalendar** — date picker.
- **Forms**: `frmPayment`, `frmPaymentManage`, `frmReport`, `frmSearch`,
  `frmStatement`, `frmCalendar`.
- **ThisWorkbook** — on open, activates Menu + refreshes summary.

### 1.3 Report types available today (`frmReport`)
VAT, Invoices, Quotes, Med Claims, Credit Notes, Full Summary — filtered by
date range + department (All/WA/WD/MC), Summary or Full mode.

---

## 2. Why the requested features need a rewrite

Every requested feature is blocked by the **Excel-as-database, single-file,
single-user** architecture:

| Requested feature | Blocker in current design |
|---|---|
| Licensing / standalone program | It's an `.xlsm` needing Excel |
| Multi-PC, central DB, "work from anywhere" | Workbook *is* the DB; one file = one user |
| Login with different users | Only `Environ$("USERNAME")` captured, no auth |
| Menu shows options only; summaries in separate windows | Menu sheet **is** the summary |
| VAT on/off, currency settings | VAT baked into sheet formulas + fixed columns |
| Fully customizable page setups + headings/footers | Layout is hardcoded cell ranges |
| CRM + Dental Lab modules | No module framework; fixed data model |
| Export to accountant (CSV/others) | No export layer today |

**Decision: rebuild as a C#/.NET desktop app on PostgreSQL**, reusing the VBA as
the *authoritative behavior spec*.

---

## 3. Confirmed technology stack

| Concern | Decision | Rationale |
|---|---|---|
| Language / UI | **C# + .NET 8, WPF** | Native Windows, closest to VBA logic |
| ORM / data | **Entity Framework Core** (Npgsql) | Clean table mapping + migrations |
| Database | **PostgreSQL** | True multi-user, ~10 concurrent users |
| Deployment | **BOTH** local-server+VPN **and** cloud-hosted (see §3.1) | User picks at install time |
| PDF / documents | **QuestPDF** with editable templates | Customizable page setup + headings/footers |
| Reporting | LINQ queries + grid/export | Replaces cell-driven filters |
| Auth | Users + roles, hashed passwords (BCrypt) | Multi-user login + per-user audit |
| Licensing | Signed license key + machine activation, offline grace | Standalone with licensing |
| Installer | Inno Setup (or WiX MSI) | Distribution + DB-mode selection |
| Config | i18n / multi-country settings (see §3.2) | Works outside South Africa |

### 3.1 Deployment — BOTH modes supported (confirmed)
The installer will let the customer choose one of:
1. **Local server + VPN** — PostgreSQL installed on an office server/PC;
   remote PCs connect over VPN. Fully on-prem, no cloud dependency.
2. **Cloud-hosted** — PostgreSQL on a managed cloud provider; any PC connects
   over TLS from anywhere.

Same application binary for both — only the **connection string** differs.
- Connection settings screen (host, port, db, user, password, SSL mode).
- Encrypted, per-machine stored connection profile.
- ~10 concurrent users target; connection pooling tuned accordingly.
- First-run wizard: "Local/VPN server" vs "Cloud server" → enter connection →
  test → run migrations if empty.

### 3.2 Multi-country support (confirmed)
Default profile = **South Africa (ZAR, 15% VAT, medical-aid claim fields)**, but
**users can manually configure for any country**:
- **Country profile** setting: country name, currency code + symbol, tax label
  (VAT/GST/Sales Tax), tax rate(s), tax-registration label (VAT No / GST No…).
- **VAT/tax on-off toggle** + editable rate.
- Currency symbol, decimal separator, thousands separator, date format.
- Medical-aid / claim fields become **optional** (toggle) for non-SA users.
- Number-format prefixes (INV/Q/MC/CN, dept segments) editable per install.
- All tax math done in code from the active country profile — never hardcoded.

---

## 4. Database schema (derived from `modConfig`)

- **users** (id, username, password_hash, role, is_active, created_at)
- **settings** (key, value) — includes **country_profile** (json: currency, tax
  label, tax rate, med-aid enabled, formats), VAT enabled, counters, bank
  details, headings/footers, deployment_mode
- **countries / tax_profiles** (id, country, currency_code, currency_symbol,
  tax_label, tax_rate, tax_reg_label, med_aid_enabled, date_format) — seed row
  for South Africa; users add/edit others
- **doctors** (was `Customers`: cust_id, name, street, suburb, city, postcode,
  id_no, tel, tax_reg_no, bhf_no, med_aid, credit_balance)
- **patients** (was `Med Customers`: bill_to_key, name, address…, id_no, tel,
  email, med_aid_name, med_aid_number, dependant_code, main_member,
  credit_balance, **folder_path** ← "open folder" feature)
- **documents** (unified INV/QTE/MC/CN header: id, doc_kind, doc_no, dept,
  recipient_type, doc_date, due_date, patient_id **(FK)**, doctor_id **(FK)**,
  appliance_type, subtotal, discount, tax_amount, total, paid, balance, status,
  source_quote_no, converted_inv_no, source_inv_no, med-aid fields, notes,
  is_voided **(soft-delete)**, created_at, modified_at, created_by **(FK users)**)
- **document_lines** (doc_id **(FK)**, line_no, code, zcode, description, qty,
  excl, tax, incl, total) — **no 15-line cap**
- **payments** (id, doc_id **(FK)**, date, amount, method, reference, notes,
  created_by, is_voided)
- **audit_log** (timestamp, user_id **(FK)**, action, doc_no, old_val, new_val,
  comment)
- **pricelist** (code, description, zcode, price_incl, taxable)
- **document_templates** (doc_kind, page_setup json, header, footer, logo)
- **doc_number_sequences** (doc_kind, dept, next_value) — **atomic** counters
- **crm_*** and **lab_*** tables (see §5.7 / §5.8)

---

## 5. Feature-by-feature delivery plan

### 5.1 Core (Phase 1)
- PostgreSQL + EF Core schema + migrations
- **First-run wizard: Local/VPN vs Cloud** connection setup + test
- Users, login, roles, per-user audit (real user, before/after values)
- Settings screen: **country profile**, **tax on/off + rate**, **currency**,
  formats, bank details, counters
- **Atomic doc-number sequences** (DB-side; safe for 10 concurrent users)

### 5.2 Documents (Phase 2)
- Port Invoice/Quote/Med Claim/Credit Note using existing rules:
  - Numbering (atomic) + overwrite guard equivalent
  - Recipient (Doctor/Private) dynamic fields
  - Quote→Invoice conversion
  - Dept-change renumber + manual rename + revert (**soft-delete/void trail**)
  - **Credit-note excess → credit store** (preserve exactly)
  - Payment reconciliation (`ReconcileDoc`)
  - Tax computed from **active country profile**

### 5.3 Customizable templates (Phase 3)
- Per-doc-kind page setup, margins, orientation
- Editable **headings & footers**, logo, banking block
- Replaces hardcoded `$A$1:$H$48` one-page fit

### 5.4 New Menu + summaries (Phase 4)
- **Menu = options only** (no data)
- Summaries open in **separate windows**
- **Open patient folder** button (from `patients.folder_path`)

### 5.5 Reports + filters (Phase 5)
Existing: date range, department, doc type. **Additions:** status
(Unpaid/Part-Paid/Paid), amount range, **aging buckets 30/60/90**, medical-aid
name, patient, doctor, balance>0 only, created-by user, recipient type. Sortable
grid, exportable.

### 5.6 Accountant export (Phase 6)
- Universal **CSV** with column mapping
- Presets: generic ledger, **QuickBooks (IIF/CSV)**, **Xero CSV**, **Sage CSV**
- Date-range + doc-type scoped

### 5.7 CRM module (Phase 7)
- Unified contacts (patients + doctors as parties)
- Leads, follow-ups, reminders, communication log
- **Sees invoices/quotes** for a contact (reads `documents`)

### 5.8 Dental Lab module (Phase 8)
- Lab jobs/orders workflow: received → in-progress → ready → delivered
- Linked to patient + doctor
- Can auto-generate a Quote or Invoice via the document engine

### 5.9 Licensing, installer, remote hardening (Phase 9)
- Signed license keys, machine activation, offline grace period
- Installer with **Local/VPN vs Cloud** DB-mode selection
- Connection pooling, migrations, backup/restore, TLS enforcement

---

## 6. Improvements — ALL adopted (confirmed)

1. **DB `ORDER BY`** replaces in-memory bubble sorts.
2. **Consolidated helpers** — one `Num`/`NrmID` equivalent, no duplicates.
3. **Atomic DB sequences** for document numbers (safe for 10 concurrent users).
4. **Tax out of formulas** — computed in code from the country profile.
5. **Foreign keys** for doctors/patients instead of string matching.
6. **Full audit** — real logged-in user + before/after values on every write.
7. **Soft-delete / void trail** for all financial records.
8. **No line-item cap** (DB model removes the 15-line limit).
9. **Centralized error handling + logging** to file/table (no `MsgBox`).
10. **Automated tests** around numbering, credit-note excess, reconciliation.

---

## 7. Confirmed decisions

| # | Question | Decision |
|---|---|---|
| 1 | Stack | ✅ **C#/.NET + PostgreSQL** |
| 2 | Concurrent users / PCs | ✅ **10** |
| 3 | Hosting | ✅ **BOTH** — local server+VPN **and** cloud, chosen at install |
| 4 | Country / tax / currency | ✅ Default **South Africa (ZAR, 15% VAT)**; **users can manually set any country** |
| 5 | Improvements | ✅ **Use all 10** |
| 6 | Priority order | Phase order per §5 (confirm or adjust) |
| 7 | Data migration | **Open** — import existing workbook data into the new DB? (please confirm) |

---

## 8. Status

- [x] Scanned all existing VBA modules and forms
- [x] Documented current architecture + data model
- [x] Proposed stack, schema, phased plan, improvements
- [x] Confirmed stack, concurrency, hosting (both), multi-country, improvements
- [ ] Confirm data migration (§7.7) + phase priority (§7.6)
- [ ] Lock scope for Phase 1 and begin scaffolding
