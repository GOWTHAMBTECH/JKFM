# SAP B1 Add-on – Purchase Transport Details (UDO `PurTransportDetails`) Change Specification

## 1) Scope
This document converts the provided requirement into an implementable technical specification for the SAP B1 add-on UDO screen **Purchase Transport Details** (`FormType: AVA_PTDF`) with child tables:

- `@AVA_PTD1` (Truck)
- `@AVA_PTD2` (Train)
- `@AVA_PTD3` (Container)
- Header `@AVA_OPTD`

## 2) Requested UI / Data Columns
Apply to all 3 child matrices (`M_Truck`, `M_Train`, `M_Cont`):

1. Add/keep **Invoice Weight** column (`U_AVA_InvoiceWeight`).
2. Add **Loading Person** (code + name behavior via BP selection).
3. Add **UnLoading Person** (code + name behavior via BP selection).
4. Add **Loading** tick column (`U_AVA_Loading`, Y/N).
5. Add **UnLoading** tick column (`U_AVA_UnLoading`, Y/N).

> Note: your XML already contains these columns in all three matrices. The key remaining work is behavior/FMS/validation and PO/AP automation.

## 3) Master Data / FMS Rules

### 3.1 Loading Person / UnLoading Person FMS
Use BP master for CFL/FMS values.

- Source: OCRD (Business Partner)
- Filter: BP Group = **Load Person Gang** (as requested)
- Store: selected BP `CardCode` in `U_AVA_LoadingPerson` / `U_AVA_UnLoadingPerson`
- Display: optionally show `CardName` using linked display column or formatted text

Implementation options:
- Preferred: set CFL conditions on `CFL_Load11/12/21/22/31/32`
- Alternate: validate in `et_CHOOSE_FROM_LIST` after selection and reject if group mismatch

## 4) PO Creation Logic (Service PO)
When tick mark selected, create Service PO line(s) based on rules below.

### 4.1 Trigger
For each child row (Truck/Train/Container):
- If `U_AVA_Loading = 'Y'` and PO not already created for loading, create loading service PO line.
- If `U_AVA_UnLoading = 'Y'` and PO not already created for unloading, create unloading service PO line.

Recommended idempotency fields (new UDFs):
- `U_AVA_LoadPODocEntry`, `U_AVA_LoadPOLineNum`
- `U_AVA_UnLoadPODocEntry`, `U_AVA_UnLoadPOLineNum`

### 4.2 Loading Charges Rule
- Applicable **only for Rack** transport scenario.
- Not applicable for truck/container (as per requirement text).
- Charge value source: Item Master UDF `U_LCHARGES`.
- Service description example: `Wheat handling Goodshed - Loading`.
- Tax Code: Tax-Exempt.

### 4.3 UnLoading Charges Rule
- Fetch from selected Warehouse (`U_AVA_WHSE`) via warehouse-linked UDF / configuration table (tonnage-based unloading charge).
- Apply to Rack/Container/Truck per your rule text.
- Service description example: `Wheat handling charges - Unloading`.
- Tax Code: Tax-Exempt.

### 4.4 Amount Calculation
Recommended calculation:

- Quantity basis = `U_AVA_InvoiceWeight` when available, else `U_AVA_QTYINTONES`
- LineAmount = RatePerTon × QuantityBasis

Validate non-zero quantity and rate before PO creation.

## 5) Train Closure → AP Invoice Automation
Requirement interpretation:
- For **Train** rows, user updates `Remarks` and sets status to Closed/Done after all GRPO done.
- Then system should auto-create AP Invoice and submit for approval.

### 5.1 Preconditions
At header/row level:
- `U_AVA_TRANSTYPE = 'Train'` (or train tab row)
- All related GRPO completed (`U_AVA_ALLGRPO = Y` or equivalent check)
- Row/Header status transitioned to Closed/Done
- Required attachments available (calculation sheet + Avery table PDF)

### 5.2 Actions
1. Create AP Invoice (supplier from transport service provider / configured BP).
2. Attach documents:
   - Calculation sheet layout
   - Avery table PDF matched with transaction
3. Submit AP Invoice into approval process (draft/approval template route).
4. After AP Invoice posting, automatically create/push Loading and Unloading charges AP invoices as required (or add separate lines based on finance policy).

### 5.3 Audit Trail (recommended)
Add UDF references on PTD row/header:
- `U_AVA_APInvDocEntry`
- `U_AVA_APInvDocNum`
- `U_AVA_LoadAPDocEntry`
- `U_AVA_UnLoadAPDocEntry`
- `U_AVA_AutoPostLog`

## 6) Validation Matrix

- Loading tick cannot be selected without Loading Person.
- UnLoading tick cannot be selected without UnLoading Person.
- Prevent duplicate PO/AP creation if reference fields already exist.
- For loading charges, block if transport type is Truck/Container.
- For unloading charges, warehouse and charge config must exist.
- Prevent closing train record unless GRPO complete and attachments present.

## 7) Issues Noted in Current XML (to be fixed)

1. In container matrix (`M_Cont`), column **Loading** (`uid="Col_Load"`) is bound to `U_AVA_UnLoading`; this should be `U_AVA_Loading`.
2. Some labels contain typo variants (`Invioce`, `Weigth`, `Conatainer`, `Depature`) – optional cleanup.
3. Some databind aliases appear with trailing spaces in header edit texts (e.g., `U_AVA_REMARKS `, `U_AVA_TOKENDATE `) – should be normalized.

## 8) Proposed Technical Tasks

1. **Database/UDF alignment**
   - Ensure all required UDFs exist with correct data types/size for 4 tables.
2. **Form XML/UI update**
   - Fix `M_Cont.Col_Load` binding.
   - Confirm visibility/order/editability of five requested columns in all matrices.
3. **FMS/CFL enhancement**
   - BP group filter for loading/unloading persons.
4. **Business logic implementation**
   - Service PO creation on tick marks.
   - Charges fetch logic from item master and warehouse config.
5. **Auto AP process**
   - Train close/done event handling.
   - Attachment validation and linking.
   - Approval submission.
6. **Idempotency + logging**
   - Reference DocEntry fields and duplicate prevention.
7. **Testing/UAT scripts**
   - Truck/Train/Container scenarios, Rack-specific loading behavior, approval flow.

## 9) Open Clarifications Needed

1. Exact transport type value representing **Rack** in DB (`U_AVA_TRANSTYPE` vs separate field).
2. Exact UDF name/table for warehouse unloading rate (e.g., `OWHS.U_AVA_UNLOADRATE`).
3. Whether loading/unloading AP invoices must be separate documents or lines in one AP invoice.
4. Approval template code/name and whether submission is draft-based.
5. Attachment repository path and mandatory file naming format.

---

If needed, I can convert this into:
- a **DI API event-by-event pseudocode** spec, or
- a **developer checklist** mapped to specific class/method names in your add-on project.
