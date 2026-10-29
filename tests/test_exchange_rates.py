import importlib.util
import unittest
from email.message import Message
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError

import pandas as pd

spec = importlib.util.spec_from_file_location('report', Path(__file__).resolve().parents[1] / 'henry_hub_ote_forward.py')
fed = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fed)


def html_table(rows):
    body = ''.join('<tr><td>' + when + '</td><td>' + rate + '</td></tr>' for when, rate in rows)
    return '<html><table><tr><th>Date</th><th><b>Rate</b></th></tr>' + body + '</table></html>'


def rate_source():
    return fed.parse_exchange_rates_html(html_table([
        ('3-JAN-00', '1.0155'), ('7-JAN-00', '1.0'), ('10-JAN-00', 'ND'),
        ('11-JAN-00', ''), ('12-JAN-00', '1.5')]))


class ParsingTests(unittest.TestCase):
    def test_parses_only_original_quotes_and_ignores_unrelated_tables(self):
        html = '<table><tr><td>unrelated</td></tr></table>' + html_table([
            ('3-JAN-00', '1.0155'), ('4-Jan-00', 'ND'), ('5-JAN-00', ''),
            ('6-JAN-00', '1.02'), ('7-JAN-00', '1.03')])
        result = fed.parse_exchange_rates_html(html)
        self.assertEqual(result['date'].dt.strftime('%Y-%m-%d').tolist(),
                         ['2000-01-03', '2000-01-06', '2000-01-07'])
        self.assertEqual(result['usd_per_eur'].tolist(), [1.0155, 1.02, 1.03])
        self.assertTrue(pd.api.types.is_float_dtype(result['usd_per_eur']))

    def test_identical_duplicates_are_allowed_but_conflicts_fail(self):
        result = fed.parse_exchange_rates_html(html_table([('3-JAN-00', '1.0155'), ('3-JAN-00', '1.0155')]))
        self.assertEqual(len(result), 1)
        with self.assertRaisesRegex(ValueError, 'rozdílné kurzy'):
            fed.parse_exchange_rates_html(html_table([('3-JAN-00', '1.0155'), ('3-JAN-00', '1.1')]))

    def test_missing_anchor_is_explicit_error(self):
        for rows in [[('4-JAN-00', '1.02')], [('3-JAN-00', 'ND')]]:
            with self.subTest(rows=rows), self.assertRaisesRegex(ValueError, '3. 1. 2000'):
                fed.parse_exchange_rates_html(html_table(rows))

    def test_invalid_dates_and_rates_fail(self):
        for day, rate in [('31-FEB-00', '1.2'), ('3-XYZ-00', '1.2'), ('4-JAN-00', '0'),
                          ('4-JAN-00', '-1'), ('4-JAN-00', 'NaN'), ('4-JAN-00', 'inf'),
                          ('4-JAN-00', '1.23garbage'), ('4-JAN-00', '1_23')]:
            with self.subTest(day=day, rate=rate), self.assertRaises(ValueError):
                fed.parse_exchange_rates_html(html_table([('3-JAN-00', '1.0155'), (day, rate)]))

    def test_block_page_is_not_accepted_as_rates(self):
        with self.assertRaisesRegex(ValueError, 'Date a Rate'):
            fed.parse_exchange_rates_html('<html>Access denied</html>')

    def test_read_uses_remote_html_and_bounded_request(self):
        content = html_table([('3-JAN-00', '1.0155')]).encode()
        headers = Message()
        headers['Content-Type'] = 'text/html; charset=utf-8'
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, limit):
                self.limit = limit
                return content
        response = Response()
        response.headers = headers
        with patch.object(fed, 'urlopen', return_value=response) as opener:
            result = fed.read_exchange_rates()
        self.assertEqual(result.iloc[0]['usd_per_eur'], 1.0155)
        self.assertEqual(response.limit, 5_000_001)
        request = opener.call_args.args[0]
        self.assertEqual(request.full_url, fed.FED_EXCHANGE_RATE_SOURCE)
        self.assertEqual(opener.call_args.kwargs['timeout'], 20)

    def test_http_failure_has_clear_error_without_file_fallback(self):
        with patch.object(fed, 'urlopen', side_effect=HTTPError('url', 403, 'Forbidden', {}, None)):
            with self.assertRaisesRegex(RuntimeError, 'Fed.*403'):
                fed.read_exchange_rates()


class ResolutionTests(unittest.TestCase):
    def test_weekend_and_multiple_nd_use_same_neighbor_mean(self):
        dates = pd.Series(pd.to_datetime(['2000-01-08', '2000-01-09', '2000-01-10', '2000-01-11']))
        result = fed.resolve_exchange_rates(dates, rate_source())
        self.assertEqual(result['usd_per_eur'].tolist(), [1.25] * 4)
        self.assertEqual(result['rate_method'].tolist(), ['neighbor_mean'] * 4)
        self.assertTrue(result['rate_date'].eq(pd.Timestamp('2000-01-07')).all())
        self.assertTrue(result['next_rate_date'].eq(pd.Timestamp('2000-01-12')).all())
        self.assertEqual(len(rate_source()), 3)  # No generated quote becomes a future neighbor.

    def test_all_pre_anchor_days_use_january_3_quote(self):
        result = fed.resolve_exchange_rates(pd.Series(['1997-01-01', '1999-12-31', '2000-01-01', '2000-01-02']), rate_source())
        self.assertEqual(result['usd_per_eur'].tolist(), [1.0155] * 4)
        self.assertEqual(result['rate_method'].tolist(), ['pre_2000'] * 4)
        self.assertTrue(result['rate_date'].eq(pd.Timestamp('2000-01-03')).all())

    def test_exact_latest_and_beyond_last(self):
        result = fed.resolve_exchange_rates(pd.Series(['2000-01-03', '2000-01-12', '2026-10-29']), rate_source())
        self.assertEqual(result['usd_per_eur'].tolist(), [1.0155, 1.5, 1.5])
        self.assertEqual(result['rate_method'].tolist(), ['exact', 'exact', 'last_available'])
        self.assertTrue(result['next_rate_date'].isna().all())

    def test_target_order_index_and_duplicate_dates_are_preserved(self):
        dates = pd.Series(['2000-01-12', '2000-01-08', '2000-01-08', '1997-01-01'],
                          index=['repeat', 'repeat', 'third', 'fourth'])
        result = fed.resolve_exchange_rates(dates, rate_source())
        self.assertEqual(result.index.tolist(), dates.index.tolist())
        self.assertEqual(result['date'].dt.strftime('%Y-%m-%d').tolist(), dates.tolist())
        self.assertEqual(result['usd_per_eur'].tolist(), [1.5, 1.25, 1.25, 1.0155])

    def test_bad_source_or_target_is_not_silently_interpolated(self):
        for rates in [pd.DataFrame({'date': ['2000-01-03'], 'usd_per_eur': [float('inf')]}),
                      pd.DataFrame({'date': ['2000-01-04'], 'usd_per_eur': [1.2]})]:
            with self.assertRaises(ValueError):
                fed.resolve_exchange_rates(pd.Series(['2000-01-05']), rates)
        with self.assertRaisesRegex(ValueError, 'neplatné datum'):
            fed.resolve_exchange_rates(pd.Series(['garbage']), rate_source())

    def test_conversion_divides_both_units_and_preserves_other_data(self):
        prices = pd.DataFrame({'date': pd.to_datetime(['2000-01-12', '2000-01-08', '2000-01-08']),
                               'price_mmbtu': [3.0, 2.5, 5.0], 'price_mwh': [9.0, 7.5, 15.0],
                               'extra': ['latest', 'first', 'second']}, index=[7, 7, 2])
        original = prices.copy(deep=True)
        result = fed.convert_to_eur(prices, rate_source())
        self.assertEqual(result['extra'].tolist(), ['first', 'second', 'latest'])
        self.assertEqual(result['price_mmbtu'].tolist(), [2.0, 4.0, 2.0])
        self.assertEqual(result['price_mwh'].tolist(), [6.0, 12.0, 6.0])
        pd.testing.assert_frame_equal(prices, original)
        self.assertEqual(result['rate_method'].tolist(), ['neighbor_mean', 'neighbor_mean', 'exact'])


if __name__ == '__main__':
    unittest.main()
