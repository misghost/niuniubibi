from datetime import date
import unittest

from crm_data import choose_owner, classify_status, parse_excel_date, select_latest_customer_rows, shift_months


class CRMDataTests(unittest.TestCase):
    def test_shift_months_handles_month_end(self):
        self.assertEqual(shift_months(date(2024, 1, 31), 1), date(2024, 2, 29))
        self.assertEqual(shift_months(date(2025, 3, 31), -1), date(2025, 2, 28))

    def test_parse_excel_date_accepts_string_and_serial(self):
        self.assertEqual(parse_excel_date("2026-03-12"), date(2026, 3, 12))
        self.assertEqual(parse_excel_date("45728"), date(2025, 3, 12))

    def test_choose_owner_uses_signer_only(self):
        self.assertEqual(choose_owner({"H": "负责人A", "J": "主笔A", "AC": "签约A"}), ("签约A", "AC"))
        self.assertEqual(choose_owner({"J": "主笔A"}), ("未分配", "system"))

    def test_classify_status(self):
        today = date(2026, 3, 15)
        self.assertEqual(classify_status(today, date(2026, 1, 1), date(2026, 3, 1)), "overdue")
        self.assertEqual(classify_status(today, date(2026, 3, 1), date(2026, 6, 1)), "due")
        self.assertEqual(classify_status(today, date(2026, 5, 1), date(2026, 8, 1)), "upcoming")

    def test_select_latest_customer_rows_keeps_newest_report(self):
        rows = [
            (2, {"A": "客户甲", "S": "2023-05-10", "AC": "签约人1"}),
            (3, {"A": "客户甲", "S": "2024-05-10", "AC": "签约人2"}),
            (4, {"A": "客户甲", "S": "2025-05-10", "AC": "签约人3"}),
            (5, {"A": "客户乙", "S": "2024-01-08", "AC": "签约人4"}),
        ]

        latest_rows = {cells["A"]: (row_number, report_date, cells["AC"]) for row_number, cells, report_date in select_latest_customer_rows(rows)}

        self.assertEqual(len(latest_rows), 2)
        self.assertEqual(latest_rows["客户甲"], (4, date(2025, 5, 10), "签约人3"))
        self.assertEqual(latest_rows["客户乙"], (5, date(2024, 1, 8), "签约人4"))


if __name__ == "__main__":
    unittest.main()
