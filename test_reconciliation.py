"""Adversarial sanity tests: run with `python -m unittest -v test_reconciliation.py`."""

from __future__ import annotations

import io
import unittest
import zipfile

from macron_vps_streamlit_app import as_cents, excel_bytes, reconcile


def sample_report():
    kyc = [
        {"PAYMETER - Account Number": "9640000001", "PAYMETER - Customer Name": "=HYPERLINK(\"http://bad.example\",\"Click\")",
         "MACRON ACCOUNT NUMBER": "701", "METER NUMBER": "123", "_row": 2},
        {"PAYMETER - Account Number": "9640000002", "PAYMETER - Customer Name": "Second Customer",
         "MACRON ACCOUNT NUMBER": "702", "METER NUMBER": "234", "_row": 3},
    ]
    macron = [
        {"DATE & TIME": "01/04/2026 10:05", "ACCOUNT NUMBER": "701", "ACTIVITY": "Token Purchase",
         "AMOUNT": "4900", "REFERENCE ID": "ref-1", "TRANSACTION ID": "t1", "_row": 2},
        {"DATE & TIME": "01/04/2026 10:06", "ACCOUNT NUMBER": "701", "ACTIVITY": "Token Purchase",
         "AMOUNT": "4900", "REFERENCE ID": "ref-2", "TRANSACTION ID": "t2", "_row": 3},
    ]
    vps = [
        {"virtual_acct_no": "9640000001", "transaction_amount_minor": "5000",
         "settled_amount_minor": "4980", "charge_amount_minor": "20",
         "cba_tran_time": "2026-04-01 10:03:00", "id": "id1", "_row": 2},
        {"virtual_acct_no": "9640000002", "transaction_amount_minor": "500",
         "settled_amount_minor": "498", "charge_amount_minor": "2",
         "cba_tran_time": "2026-04-01 10:03:00", "id": "id2", "_row": 3},
    ]
    paymeter = [
        {"Account Number": "9640000001", "Reference": "ref-1", "Input Amount": "5000",
         "Transaction Amount": "4900", "Created At": "01/04/2026 10:04", "_row": 2},
    ]
    return reconcile({"m": macron, "v": vps, "p": paymeter, "k": kyc})


class ReconciliationTests(unittest.TestCase):
    def test_one_receipt_never_used_twice(self):
        report = sample_report()
        self.assertEqual(len(report["rows"]), 3)
        self.assertEqual(report["matched_count"], 1)
        self.assertEqual(len([r for r in report["rows"] if r["status"] == "MACRON_ONLY_TOKEN"]), 1)
        self.assertEqual(len([r for r in report["rows"] if r["status"] == "VPS_ONLY_PAYMENT"]), 1)

    def test_missing_source_columns_are_blank(self):
        rows = sample_report()["rows"]
        m = next(r for r in rows if r["status"] == "MACRON_ONLY_TOKEN")
        v = next(r for r in rows if r["status"] == "VPS_ONLY_PAYMENT")
        self.assertIsNone(m["VPS | id"])
        self.assertIsNone(v["Macron | TRANSACTION ID"])
        self.assertEqual(m["cash_in"], 0)
        self.assertEqual(v["cash_out"], 0)
        self.assertEqual(v["cash_in"], 400)

    def test_fee_settlement_interest(self):
        report = sample_report()
        m = next(r for r in report["rows"] if r["source"] == "Bank/Macron")
        self.assertEqual(m["cash_in"], 4900)
        self.assertEqual(m["cash_out"], 4900)
        self.assertEqual(m["VPS_actual_bank_charge_NGN"], 20)
        self.assertEqual(m["interest"], 80)
        self.assertEqual(m["transaction_date"].isoformat(), "2026-04-01")

    def test_small_payment_and_repeat_token_flagged(self):
        report = sample_report()
        self.assertEqual(len(report["small"]), 1)
        self.assertEqual(len(report["duplicate"]), 1)

    def test_malicious_name_is_not_excel_formula(self):
        result = excel_bytes(sample_report())
        with zipfile.ZipFile(io.BytesIO(result)) as z:
            xml_sheets = [x for x in z.namelist() if x.startswith("xl/worksheets/sheet") and x.endswith(".xml")]
            self.assertTrue(xml_sheets)
            for name in xml_sheets:
                self.assertNotIn(b"<f>", z.read(name))
            self.assertIn(b"HYPERLINK", z.read("xl/sharedStrings.xml"))

    def test_invalid_amount_is_not_silently_zero(self):
        self.assertIsNone(as_cents("not-an-amount"))
        self.assertEqual(as_cents("100.25"), 10025)


if __name__ == "__main__":
    unittest.main()
