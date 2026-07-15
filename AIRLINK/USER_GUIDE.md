# AIRLINK ERP User Guide (English)

Recommended workflow and common mistakes.  
Applies from build `airlink-2026-07-14e`.

> In-app: **Settings → User Guide → Open User Guide** (`user-guide.html`, switch to **English**).

## Standard workflow

```
Client → Quotation (Accepted) → Job → Customer PO → Vendor PO → AP (pay vendor) → Job Completed → AR (invoice client)
```

Sidebar order: Clients → Quotation → Job → **Vendor PO** → AR → AP → Vendors

---

## 0. Before you start (Admin / Root)

1. Set top-right **Region** (HK / TW / SG / MY / ID). New data is tagged to the current region.
2. **Settings → General**: language, time zone, company name.
3. **Settings → Backup**: sync key / live sync. After critical actions, use **Sync now**.
4. **Settings → Users**: accounts, permissions, region access. Deleting a user must show “removed from cloud”.

---

## 1. Master data

### Clients
- **Clients → + Add Client**; next customer number is suggested.
- Optional: Invoice Title, payment terms (30/45/60/90 or custom), BU.

### Vendors
- Excel-aligned columns: Vendor No, Company Name, GST #, Legal Rep, Address, Postal Code, Company Phone, Main Contact, E-mail, Mobile, Job Title, Bank, Branch, Account #, SWIFT, Credit Term.

### Move to another region
Wrong region? Open the record → **Edit** → **Move to region** at the top → Save.  
Moving a Client does **not** auto-move Jobs / Invoices / POs.

---

## 2. Transaction flow

| Step | Module | Notes |
|------|--------|-------|
| 1 | Quotation | Client; must be **Accepted** to link a job |
| 2 | Job | Client, quotations, customer PO lines; detail shows BU / Vendor / Vendor PO |
| 3 | Vendor PO | See section 3 |
| 4 | AP | Link job → vendor → optional **AIRLINK PO** / **PO Amount** |
| 5 | Complete | Job status **Completed** |
| 6 | AR | Link job → invoice client |

Blue links in Detail open related records.

---

## 3. Vendor PO columns

Vendor Code → Name → For Project/ Inventory → Type → Airlink PO # → PO Date → Airlink PO Currency → Airlink PO Amt → **Airlink PO Amt in {local currency}** (region base, e.g. MY→MYR).

---

## 4. Reports / Import

- Monthly Report: click **Sync monthly** after changes.
- Import order: Clients → Vendors → Quotations → Jobs → **Vendor PO** → AR → AP.
- **Job import:** fill **Customer No** only — Client Company is optional and auto-mapped from the client master. Export includes Customer No. Fix company name later on the client record if needed.
- **No formal quotation:** on Add/Edit Job, tick **No quotation** to auto-assign HK00001 / TW00001 / SG00001… for that region and create an Accepted placeholder quotation on Save.

---

## 5. Common mistakes

- Job/quotation before client exists  
- Link job before quotation Accepted  
- Wrong region → use **Move to region**, don’t duplicate  
- Delete user without Sync now → may return in a new browser  
- Local-only storage without live sync  

See also the English panel in `user-guide.html`.
