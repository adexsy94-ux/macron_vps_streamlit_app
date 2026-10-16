# Macron / VPS / Paymeter / KYC Reconciliation App

## Run locally

Use Python 3.11 or later (3.10+ supported). In the folder containing `macron_vps_streamlit_app.py` and `requirements.txt`:

```bash
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
# macOS / Linux
# source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: -r requirements.txt
python -m streamlit run macron_vps_streamlit_app.py
```

Open the local address printed by Streamlit (usually `http://localhost:8501`). Upload four files: Macron .xlsx, Paymeter .csv, VPS .xlsx, and KYC .xlsx, then select Generate reconciliation and Download Excel.

## Output sheets

- **Combined** — all Macron rows followed by VPS-only rows. Each bank receipt is used once; the other side is blank when missing. Starts with seven accounting fields: transaction_date, customer_name, account_number, cash_in, cash_out, interest, source.
- **VPS No Token** — VPS-only deposits, including payments <= NGN 1,000.
- **Macron No Payment** — Macron-only token purchases and PayBill entries.
- **Review Matches** — tentative pairs needing manual review.
- **Small Unmatched** — all unmatched payments <= NGN 1,000.
- **Possible Duplicates** — same virtual account/day/token value with more Macron token issuances than matching receipts.
- **KYC Mapping** — account links and KYC names.
- **Customer Exposure** — provisional account-level ledger, not confirmed debt.
- **PayBill Groups** — grouped PayBill allocations.
- **Paymeter No Reference** — Paymeter records whose reference is absent from Macron.
- **Summary**, **Methodology**.

## Accounting

- `cash_in = VPS transaction_amount_minor - 100 NGN` (missing VPS: 0)
- `cash_out = Macron AMOUNT` (missing Macron: 0)
- `actual_bank_charge = VPS gross - VPS settled_amount_minor`
- `interest = 100 NGN - actual_bank_charge` (missing VPS: 0)
- `source = Bank/Macron`, `Bank`, or `Macron`
- Dates use VPS local CBA date first, or UTC VPS created_at converted to Lagos; if missing, Macron date. Excel format `20-Feb-24`.

**Important:** The VPS file's `_minor` fields are treated as NGN because the supplied data confirms that `5000` is NGN 5,000. If the provider changes to actual kobo integers, adapt the units before relying on output.

PayBill is handled separately from prepaid tokens; one VPS receipt can cover multiple PayBill lines (shown once on a leader row). Paired transaction review candidates and grouped PayBill lines are **not** evidence of fully reconciled balances.

## Security and privacy

- Customer files are read in memory and not sent to an external API by this application. Run it on a private, access-controlled server or locally, not an unauthenticated public Streamlit URL.
- Excel formulas/macros/remote links from inputs are never executed. Output treats all customer-supplied text as literal strings to prevent spreadsheet formula injection.
- XLSX uploads have per-file and decompressed-size limits.
- Installation is pinned to Streamlit 1.66.0 and XlsxWriter 3.2.9. A known Streamlit Windows SSRF issue affecting versions before 1.54.0 was fixed in that release (CVE-2026-33682); these pins are newer, but **do not guarantee absence of future or transitive vulnerabilities**.
- Before deployment, independently audit the resolved dependency tree and inspect the environment: `python -m pip install pip-audit` then `python -m pip_audit`. Run `python -m pip check`. For stronger supply-chain protection, generate your own fully hash-locked dependency lockfile from a trusted package index.

## Verification principles

1. Each Macron source row appears exactly once; each VPS source row appears exactly once.
2. A single VPS payment is never linked to two Macron transaction rows.
3. Macron-to-Paymeter matching uses exact reference, and Paymeter-to-VPS matching uses the virtual account, gross payment, and timestamp.
4. KYC mapping is used even when a Paymeter transaction is missing. Reference-derived mapping is labeled unverified when full KYC is absent.
5. The match confidence prevents ambiguous or delayed transactions from silently becoming confirmed.
6. An unmatched payment or token issuance remains an exception to investigate; it does not prove debt, fraud, or duplicate issuance.

## Test the matching engine

```bash
python -m unittest -v test_reconciliation.py
```

The tests cover a payment linked only once, a repeat-token candidate, an unmatched small payment, empty source-side columns, the bank-fee calculation, and malicious spreadsheet-formula text remaining inactive in the exported workbook.
