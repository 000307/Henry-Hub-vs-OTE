import importlib.util
from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd
from pandas.testing import assert_frame_equal

_spec = importlib.util.spec_from_file_location('cnb_exchange_backend', Path(__file__).resolve().parents[1] / 'henry_hub_ote_forward.py')
cnb = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cnb)


def rates(rows):
    return pd.DataFrame(rows, columns=['date', 'czk_per_usd'])


def prices(dates, values=None):
    return pd.DataFrame({'date': dates, 'price_czk_mwh': values if values is not None else [500.0]*len(dates)})


class ParseTests(unittest.TestCase):
    def test_usd_column_is_found_by_currency_and_dates_sorted(self):
        text = '\ufeffDatum|1 EUR|1 USD|100 JPY\n05.01.2026|24,9|20,500|15,9\n02.01.2026|24,8|20.250|15,8\n'
        result = cnb.parse_cnb_usd_rates(text)
        self.assertEqual(result['date'].dt.strftime('%Y-%m-%d').tolist(), ['2026-01-02', '2026-01-05'])
        self.assertEqual(result['czk_per_usd'].tolist(), [20.25,20.5])

    def test_amount_100_usd_is_converted_to_one_usd(self):
        result = cnb.parse_cnb_usd_rates('Datum|100 USD|1 EUR\n02.01.2026|2050,0|24,8')
        self.assertEqual(result.loc[0,'czk_per_usd'],20.5)

    def test_duplicate_same_rate_allowed_conflicting_rate_rejected(self):
        text = 'Datum|1 USD\n02.01.2026|20,500\n02.01.2026|20.5'
        self.assertEqual(len(cnb.parse_cnb_usd_rates(text)),1)
        with self.assertRaisesRegex(ValueError,'rozdílné kurzy'):
            cnb.parse_cnb_usd_rates(text+'\n02.01.2026|21')

    def test_only_a_valid_header_may_represent_an_empty_year(self):
        result = cnb.parse_cnb_usd_rates('Datum|1 EUR|1 USD\n')
        self.assertTrue(result.empty)
        self.assertEqual(str(result['date'].dtype),'datetime64[ns]')
        for text in ('', 'Access denied', '<html>Blocked</html>', 'Datum|1 EUR'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                cnb.parse_cnb_usd_rates(text)

    def test_invalid_date_and_malformed_rows_are_rejected(self):
        for row in ('31.02.2026|20,5','02.01.2026','02.01.2026|20,5|extra','2026-01-02|20,5'):
            with self.subTest(row=row), self.assertRaises(ValueError):
                cnb.parse_cnb_usd_rates('Datum|1 USD\n'+row)

    def test_rates_must_be_finite_positive(self):
        for value in ('0','-1','NaN','Inf','Infinity','20,5junk','1e300',''):
            with self.subTest(value=value), self.assertRaises(ValueError):
                cnb.parse_cnb_usd_rates('Datum|1 USD\n02.01.2026|'+value)
        for header in ('Datum|0 USD','Datum|1 USD|100 USD','Datum|USD'):
            with self.subTest(header=header), self.assertRaises(ValueError):
                cnb.parse_cnb_usd_rates(header)


class FetchTests(unittest.TestCase):
    class Response:
        def __init__(self, body, content_type='text/plain; charset=UTF-8'):
            self.body=body
            self.headers={'Content-Type':content_type}
            self.limit=None
        def __enter__(self): return self
        def __exit__(self,*args): pass
        def read(self,limit):
            self.limit=limit
            return self.body[:limit]

    def test_fetch_has_bounded_read_timeout_and_verified_yearly_url(self):
        response=self.Response(b'Datum|1 USD\n02.01.2026|20,5')
        with patch.object(cnb,'urlopen',return_value=response) as opening:
            result=cnb.fetch_cnb_usd_rates(2026)
        self.assertEqual(opening.call_args.kwargs['timeout'],20)
        self.assertEqual(opening.call_args.args[0].full_url,cnb.CNB_YEARLY_RATE_TEMPLATE.format(year=2026))
        self.assertEqual(response.limit,cnb._CNB_MAX_TEXT_BYTES+1)
        self.assertEqual(result['czk_per_usd'].tolist(),[20.5])

    def test_html_200_cannot_be_mistaken_for_empty_current_year(self):
        for content_type in ('text/html; charset=UTF-8','text/plain'):
            with self.subTest(content_type=content_type), patch.object(cnb,'urlopen',return_value=self.Response(b'<html>Access denied</html>',content_type)), self.assertRaises(ValueError):
                cnb.fetch_cnb_usd_rates(2026)

    def test_header_only_is_valid_but_wrong_year_rows_are_not(self):
        with patch.object(cnb,'urlopen',return_value=self.Response(b'Datum|1 USD\n')):
            self.assertTrue(cnb.fetch_cnb_usd_rates(2026).empty)
        with patch.object(cnb,'urlopen',return_value=self.Response(b'Datum|1 USD\n31.12.2025|20,5')), self.assertRaisesRegex(ValueError,'jiného roku'):
            cnb.fetch_cnb_usd_rates(2026)

    def test_oversized_response_and_network_failure_are_explicit(self):
        with patch.object(cnb,'urlopen',return_value=self.Response(b'a'*(cnb._CNB_MAX_TEXT_BYTES+1))), self.assertRaisesRegex(ValueError,'limit'):
            cnb.fetch_cnb_usd_rates(2026)
        with patch.object(cnb,'urlopen',side_effect=TimeoutError('timed out')), self.assertRaisesRegex(RuntimeError,'2026'):
            cnb.fetch_cnb_usd_rates(2026)


class ConvertTests(unittest.TestCase):
    def test_same_day_both_units_extra_columns_and_input_order_preserved(self):
        data=prices(['2026-01-06','2026-01-05'],[480.0,500.0])
        data['source_url']=['ote/a','ote/b']
        data['extra']=[7,8]
        data.index=pd.Index([42,10],name='original')
        original=data.copy()
        with patch.object(cnb,'fetch_cnb_usd_rates',return_value=rates([('2026-01-05',20.0),('2026-01-06',24.0)])) as fetch:
            result=cnb.convert_ote_czk_to_usd(data)
        fetch.assert_called_once_with(2026)
        self.assertTrue(result.index.equals(data.index))
        self.assertEqual(result['source_url'].tolist(),['ote/a','ote/b'])
        self.assertEqual(result['extra'].tolist(),[7,8])
        self.assertEqual(result['price_mwh'].tolist(),[20.0,25.0])
        self.assertAlmostEqual(result['price_mmbtu'].iloc[0],20.0/3.4121416)
        self.assertEqual(result['cnb_rate_date'].dt.strftime('%Y-%m-%d').tolist(),['2026-01-06','2026-01-05'])
        assert_frame_equal(data,original)

    def test_weekend_uses_friday_not_future_monday(self):
        with patch.object(cnb,'fetch_cnb_usd_rates',return_value=rates([('2026-01-02',20.0),('2026-01-05',25.0)])):
            result=cnb.convert_ote_czk_to_usd(prices(['2026-01-03','2026-01-04']))
        self.assertEqual(result['czk_per_usd'].tolist(),[20.0,20.0])
        self.assertEqual(result['cnb_rate_date'].dt.strftime('%Y-%m-%d').tolist(),['2026-01-02']*2)

    def test_jan_first_loads_previous_year_only_once(self):
        def fetch(year):
            return {2026:rates([('2026-01-02',25.0)]),2025:rates([('2025-12-31',20.0)])}[year]
        with patch.object(cnb,'fetch_cnb_usd_rates',side_effect=fetch) as fetching:
            result=cnb.convert_ote_czk_to_usd(prices(['2026-01-01','2026-01-01','2026-01-02']))
        self.assertEqual([call.args[0] for call in fetching.call_args_list],[2026,2025])
        self.assertEqual(result['czk_per_usd'].tolist(),[20.0,20.0,25.0])

    def test_current_year_can_be_header_only_on_jan_first(self):
        def fetch(year):
            return cnb.parse_cnb_usd_rates('Datum|1 USD\n') if year==2026 else rates([('2025-12-31',20.0)])
        with patch.object(cnb,'fetch_cnb_usd_rates',side_effect=fetch):
            result=cnb.convert_ote_czk_to_usd(prices(['2026-01-01']))
        self.assertEqual(result['cnb_rate_date'].iloc[0],pd.Timestamp('2025-12-31'))

    def test_loads_only_input_years_without_unneeded_intermediate_years(self):
        def fetch(year):
            return rates([(f'{year}-06-01',20.0)])
        with patch.object(cnb,'fetch_cnb_usd_rates',side_effect=fetch) as fetching:
            cnb.convert_ote_czk_to_usd(prices(['2026-06-02','2023-06-02','2026-06-03']))
        self.assertEqual([call.args[0] for call in fetching.call_args_list],[2023,2026])

    def test_empty_input_does_not_fetch(self):
        data=prices([])
        with patch.object(cnb,'fetch_cnb_usd_rates') as fetching:
            result=cnb.convert_ote_czk_to_usd(data)
        fetching.assert_not_called()
        self.assertTrue(result.empty)
        self.assertIn('cnb_rate_date',result)

    def test_no_previous_rate_fails_instead_of_using_first_future_rate(self):
        def fetch(year):
            return rates([('2026-01-02',20.0)]) if year==2026 else rates([])
        with patch.object(cnb,'fetch_cnb_usd_rates',side_effect=fetch), self.assertRaisesRegex(ValueError,'předchozí kurz'):
            cnb.convert_ote_czk_to_usd(prices(['2026-01-01']))

    def test_invalid_input_dates_prices_and_conflicting_rates_are_rejected(self):
        for value in ('31.02.2026','not a date',None,45000):
            with self.subTest(value=value), patch.object(cnb,'fetch_cnb_usd_rates') as fetching, self.assertRaises(ValueError):
                cnb.convert_ote_czk_to_usd(prices([value]))
            fetching.assert_not_called()
        for value in (float('nan'),float('inf')):
            with self.subTest(value=value), patch.object(cnb,'fetch_cnb_usd_rates') as fetching, self.assertRaises(ValueError):
                cnb.convert_ote_czk_to_usd(prices(['2026-01-02'],[value]))
            fetching.assert_not_called()
        with patch.object(cnb,'fetch_cnb_usd_rates',return_value=rates([('2026-01-02',20.0),('2026-01-02',21.0)])), self.assertRaisesRegex(ValueError,'rozdílné kurzy'):
            cnb.convert_ote_czk_to_usd(prices(['2026-01-02']))

    def test_zero_negative_prices_and_decimal_comma_are_valid(self):
        with patch.object(cnb,'fetch_cnb_usd_rates',return_value=rates([('2026-01-02',20.0)])):
            result=cnb.convert_ote_czk_to_usd(prices(['2026-01-02']*3,[0.0,-20.0,'40,5']))
        self.assertEqual(result['price_mwh'].tolist(),[0.0,-1.0,2.025])


if __name__=='__main__':
    unittest.main()
