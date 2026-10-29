from datetime import date, datetime
import importlib.util
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook


spec = importlib.util.spec_from_file_location("ote_forward_backend", Path(__file__).resolve().parents[1] / "henry_hub_ote_forward.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


class OteForwardTests(unittest.TestCase):
    def test_report_loader_reads_once_and_does_not_modify_input(self):
        path = self.workbook(['Datum', 'Nov-2026'], [[date(2026, 10, 29), 80]])
        original = path.read_bytes()
        with patch.object(m, 'resolve_input', return_value=path) as resolve, patch.object(m, 'read_ote_forward', wraps=m.read_ote_forward) as reader:
            payload = m.load_ote_forward_for_report(str(path))
        resolve.assert_called_once_with(str(path))
        reader.assert_called_once_with(path)
        self.assertIsNone(payload['error'])
        self.assertEqual(payload['rows'][0]['price'], 80)
        self.assertEqual(path.read_bytes(), original)

    def test_missing_or_invalid_file_yields_visible_error_and_empty_ote_series(self):
        for exc in (FileNotFoundError('Forwardy OTE.xlsx'), ValueError('Neznámý produkt')):
            with self.subTest(exc=type(exc)), patch.object(m, 'resolve_input', side_effect=exc):
                payload = m.load_ote_forward_for_report()
            self.assertEqual(payload['rows'], [])
            self.assertIsNone(payload['quote_date'])
            self.assertIn(str(exc), payload['error'])

    def test_each_forward_uses_its_quotation_date_and_existing_fed_gap_rule(self):
        rates = pd.DataFrame({'date': pd.to_datetime(['2000-01-03', '2026-10-23', '2026-10-28', '2026-10-29']),
                              'usd_per_eur': [1.0155, 1.0, 1.4, 1.5]})
        cme = {'rows': [{'contract_date': '2027-01-01'}], 'trade_date': '2026-10-26'}
        ote = {'rows': [{'delivery_start': '2030-01-01'}], 'quote_date': '2026-10-29'}
        m.attach_forward_exchange_rate(cme, rates, 'trade_date')
        m.attach_forward_exchange_rate(ote, rates, 'quote_date')
        self.assertEqual(cme['fx_rate'], 1.2)
        self.assertEqual(cme['fx_rate_method'], 'neighbor_mean')
        self.assertEqual(ote['fx_rate'], 1.5)
        self.assertEqual(ote['fx_rate_date'], '2026-10-29')
        self.assertEqual(ote['fx_rate_method'], 'exact')
        empty = {'rows': []}
        with patch.object(m, 'resolve_exchange_rates') as resolve:
            m.attach_forward_exchange_rate(empty, rates, 'quote_date')
        resolve.assert_not_called()

    def workbook(self, header, rows):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "Forwardy OTE.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(header)
        for row in rows:
            sheet.append(row)
        workbook.save(path)
        workbook.close()
        return path

    def test_nearest_unsorted_date_wins_and_values_are_not_filled_from_others(self):
        path = self.workbook(['Datum', 'Nov-2026', 'Q1-2027', 'Cal-2028'], [
            [date(2026, 10, 20), 999, 999, 999],
            [date(2026, 10, 28), None, 77, '-'],
            [date(2026, 10, 21), 888, 888, 888]])
        result = m.read_ote_forward(path, today=date(2026, 10, 29))
        self.assertEqual(result['quote_date'], '2026-10-28')
        self.assertEqual([x['price'] for x in result['rows']], [77])
        self.assertEqual(result['skipped_products'], ['Nov-2026', 'Cal-2028'])

    def test_equal_distance_prefers_earlier_date(self):
        path = self.workbook(['Datum', 'Nov-2026'], [[date(2026, 10, 30), 30], [date(2026, 10, 28), 28]])
        result = m.read_ote_forward(path, today=date(2026, 10, 29))
        self.assertEqual(result['quote_date'], '2026-10-28')
        self.assertEqual(result['rows'][0]['price'], 28)

    def test_future_nearest_date_is_allowed(self):
        path = self.workbook(['Datum', 'Nov-2026'], [[date(2026, 10, 30), 30], [date(2026, 10, 20), 20]])
        self.assertEqual(m.read_ote_forward(path, today=date(2026, 10, 29))['quote_date'], '2026-10-30')

    def test_defaults_to_prague_calendar_day(self):
        path = self.workbook(['Datum', 'Nov-2026'], [[date(2026, 10, 29), 29], [date(2026, 10, 30), 30]])
        with patch.object(m.pd.Timestamp, 'now', return_value=m.pd.Timestamp('2026-10-30T00:15:00+02:00')) as now:
            self.assertEqual(m.read_ote_forward(path)['quote_date'], '2026-10-30')
            now.assert_called_once_with(tz='Europe/Prague')

    def test_endpoints_cover_exact_delivery_periods(self):
        expected = {
            'Dec-2026': ('month', '2026-12-01', '2027-01-01'),
            'Q1-2027': ('quarter', '2027-01-01', '2027-04-01'),
            'Q2-2027': ('quarter', '2027-04-01', '2027-07-01'),
            'Q3-2027': ('quarter', '2027-07-01', '2027-10-01'),
            'Q4-2027': ('quarter', '2027-10-01', '2028-01-01'),
            'Cal-2028': ('calendar', '2028-01-01', '2029-01-01')}
        for label, (kind, start, end) in expected.items():
            with self.subTest(product=label):
                result = m._ote_forward_product(label)
                self.assertEqual((result['kind'], result['delivery_start'], result['delivery_end']), (kind, start, end))

    def test_zero_negative_and_text_decimal_prices(self):
        path = self.workbook(['Datum', 'Nov-2026', 'Q1-2027', 'Cal-2028', 'Dec-2026'], [
            ['29.10.2026', 0, '-1,25', '1 234,567', '8.75']])
        result = m.read_ote_forward(path, today=date(2026, 10, 29))
        self.assertEqual([row['price'] for row in result['rows']], [0, -1.25, 1234.567, 8.75])

    def test_identical_duplicate_quote_dates_allowed(self):
        path = self.workbook(['Datum', 'Nov-2026'], [['2026-10-29', 1], ['2026-10-29', '1,0']])
        self.assertEqual(len(m.read_ote_forward(path, today=date(2026, 10, 29))['rows']), 1)

    def test_conflicting_duplicate_quote_dates_rejected(self):
        path = self.workbook(['Datum', 'Nov-2026'], [['2026-10-29', 1], ['2026-10-29', 2]])
        with self.assertRaisesRegex(ValueError, 'rozdílné ceny'):
            m.read_ote_forward(path, today=date(2026, 10, 29))

    def test_invalid_selected_price_rejected(self):
        for bad in ('garbage', 'inf', True, '1,234.56'):
            with self.subTest(bad=bad):
                path = self.workbook(['Datum', 'Nov-2026'], [['2026-10-29', bad]])
                with self.assertRaisesRegex(ValueError, 'vadná cena'):
                    m.read_ote_forward(path, today=date(2026, 10, 29))

    def test_unselected_prices_are_not_used_or_validated(self):
        path = self.workbook(['Datum', 'Nov-2026'], [['2026-10-29', 3], ['2026-10-20', 'bad']])
        self.assertEqual(m.read_ote_forward(path, today=date(2026, 10, 29))['rows'][0]['price'], 3)

    def test_unknown_product_and_duplicate_product_rejected(self):
        for header in (['Datum', 'Q5-2027'], ['Datum', 'Notes'], ['Datum', 'Nov-2026', 'NOV-2026']):
            with self.subTest(header=header):
                path = self.workbook(header, [[date(2026, 10, 29), 1, 2][:len(header)]])
                with self.assertRaises(ValueError):
                    m.read_ote_forward(path, today=date(2026, 10, 29))

    def test_missing_date_or_empty_data_rejected(self):
        for header, rows in ((['Price', 'Nov-2026'], [[date(2026, 10, 29), 1]]),
                             (['Datum', 'Nov-2026'], []),
                             (['Datum', 'Nov-2026'], [[None, 1]])):
            with self.subTest(header=header, rows=rows):
                path = self.workbook(header, rows)
                with self.assertRaises(ValueError):
                    m.read_ote_forward(path, today=date(2026, 10, 29))

    def test_first_sheet_read_and_source_file_unchanged(self):
        path = self.workbook(['Date', 'Jan-2027'], [[date(2026, 10, 29), 11]])
        original = path.read_bytes()
        m.read_ote_forward(path, today=date(2026, 10, 29))
        self.assertEqual(path.read_bytes(), original)


if __name__ == '__main__':
    unittest.main()
