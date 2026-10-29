import base64
import importlib.util
import unittest
import warnings
from contextlib import redirect_stdout
from io import StringIO
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError

import pandas as pd

HERE = Path(__file__).parent
spec = importlib.util.spec_from_file_location('report_eia', HERE.parent / 'henry_hub_ote_forward.py')
eia = importlib.util.module_from_spec(spec)
spec.loader.exec_module(eia)
FIXTURES = HERE / 'fixtures'


def fixture(name='eia_henry_hub_data1'):
    return base64.b64decode((FIXTURES / (name + '.xls.b64')).read_text())


def synthetic_frame(days=32):
    rows = [['Title', None], ['Sourcekey', 'RNGWHHD'], ['Date', eia._EIA_PRICE_COLUMN]]
    rows.extend([[datetime(1997, 1, 2) + timedelta(days=number), 3.82] for number in range(days)])
    return pd.DataFrame(rows)


class EiaWorkbookTests(unittest.TestCase):
    def test_real_biff_xls_uses_data1_and_third_row_headers(self):
        content = fixture()
        self.assertTrue(content.startswith(eia._EIA_XLS_SIGNATURE))
        result = eia.parse_henry_hub_xls(content)
        self.assertEqual(list(result.columns), ['date', 'price_mmbtu', 'price_mwh'])
        self.assertEqual(len(result), 44)
        self.assertEqual(result.iloc[0]['date'], pd.Timestamp('1997-01-02'))
        self.assertEqual(result.iloc[-1]['date'], pd.Timestamp('1997-02-15'))
        self.assertTrue(result['date'].is_monotonic_increasing)
        self.assertFalse(result['date'].duplicated().any())
        self.assertFalse(result['price_mmbtu'].eq(99).any())  # Contents has misleading valid-looking data.

    def test_reordered_columns_numeric_comma_and_serial_dates(self):
        result = eia.parse_henry_hub_xls(fixture()).set_index('date')
        # 3 January was an unformatted numeric Excel serial; 9 January has text 3,82.
        self.assertAlmostEqual(result.loc['1997-01-03', 'price_mmbtu'], 3.82)
        self.assertAlmostEqual(result.loc['1997-01-09', 'price_mmbtu'], 3.82)
        self.assertAlmostEqual(result.loc['1997-01-04', 'price_mmbtu'], 3.82)  # ISO string.
        self.assertAlmostEqual(result.loc['1997-01-13', 'price_mmbtu'], 3.82)  # US date string.
        self.assertAlmostEqual(result.loc['1997-01-02', 'price_mwh'], 3.82 * 3.4121416)

    def test_duplicates_normalize_to_day_and_average_without_filling_gaps(self):
        result = eia.parse_henry_hub_xls(fixture()).set_index('date')
        self.assertEqual(result.loc['1997-01-07', 'price_mmbtu'], 3.0)  # (2 + 4) / 2, one entry has 15:30.
        self.assertNotIn(pd.Timestamp('1997-01-05'), result.index)
        self.assertFalse(result.index.to_series().dt.hour.any())
        self.assertFalse((result.index >= pd.Timestamp('1997-03-01')).any())  # Invalid price rows were dropped.

    def test_exact_data1_sheet_and_xlrd_engine_are_requested_once(self):
        original = pd.read_excel
        with patch.object(eia.pd, 'read_excel', wraps=original) as reader:
            eia.parse_henry_hub_xls(fixture())
        self.assertEqual(reader.call_count, 1)
        self.assertIsInstance(reader.call_args.args[0], BytesIO)
        self.assertEqual(reader.call_args.kwargs, {'sheet_name': 'Data 1', 'header': None, 'engine': 'xlrd'})

    def test_missing_data1_never_uses_contents(self):
        with self.assertRaisesRegex(ValueError, "chybí požadovaný list 'Data 1'"):
            eia.parse_henry_hub_xls(fixture('eia_henry_hub_missing_sheet'))

    def test_wrong_header_does_not_assume_first_two_columns(self):
        with self.assertRaisesRegex(ValueError, 'chybí záhlaví Date'):
            eia.parse_henry_hub_xls(fixture('eia_henry_hub_wrong_header'))

    def test_html_empty_and_damaged_binary_are_clear_errors(self):
        with self.assertRaisesRegex(ValueError, 'HTML/XML místo souboru XLS'):
            eia.parse_henry_hub_xls(b'\n <!DOCTYPE html><html>Access denied</html>')
        with self.assertRaisesRegex(ValueError, 'není platný binární soubor XLS'):
            eia.parse_henry_hub_xls(b'')
        with redirect_stdout(StringIO()):
            with self.assertRaisesRegex(ValueError, 'poškozený|nepodařilo přečíst'):
                eia.parse_henry_hub_xls(fixture()[:100])

    def test_dependency_error_explains_xlrd_installation(self):
        with patch.object(eia.pd, 'read_excel', side_effect=ImportError('xlrd missing')):
            with self.assertRaisesRegex(RuntimeError, r'pip install.*xlrd'):
                eia.parse_henry_hub_xls(fixture())

    def test_invalid_nonfinite_dates_prices_and_nonpositive_prices_are_dropped(self):
        raw = synthetic_frame()
        extra = pd.DataFrame([
            [float('inf'), 3.82], ['garbage', 3.82], [pd.NaT, 3.82],
            [datetime(2000, 1, 1), float('inf')], [datetime(2000, 1, 2), float('nan')],
            [datetime(2000, 1, 3), -1], [datetime(2000, 1, 4), 0],
            [datetime(2000, 1, 5), '3.82garbage'], [datetime(2000, 1, 6), 1e308],
        ])
        with patch.object(eia.pd, 'read_excel', return_value=pd.concat([raw, extra], ignore_index=True)):
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', RuntimeWarning)
                result = eia.parse_henry_hub_xls(fixture())
        self.assertEqual(len(result), 32)
        self.assertTrue(result[['price_mmbtu', 'price_mwh']].apply(lambda column: column.map(eia.math.isfinite)).all().all())

    def test_minimum_is_thirty_unique_valid_days_after_grouping(self):
        raw = synthetic_frame(29)
        raw = pd.concat([raw, raw.iloc[3:]], ignore_index=True)
        with patch.object(eia.pd, 'read_excel', return_value=raw):
            with self.assertRaisesRegex(ValueError, 'méně než 30'):
                eia.parse_henry_hub_xls(fixture())


class EiaDownloadTests(unittest.TestCase):
    def test_download_occurs_once_and_passes_binary_body_directly(self):
        content = fixture()
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, limit):
                self.limit = limit
                return content
        response = Response()
        with patch.object(eia, 'urlopen', return_value=response) as opener:
            result = eia.read_henry_hub()
        self.assertEqual(len(result), 44)
        self.assertEqual(opener.call_count, 1)
        self.assertEqual(response.limit, 20_000_001)
        request = opener.call_args.args[0]
        self.assertEqual(request.full_url, eia.EIA_PRICE_XLS_URL)
        self.assertEqual(request.get_header('Referer'), eia.EIA_PRICE_SOURCE_URL)
        self.assertIn('application/vnd.ms-excel', request.get_header('Accept'))
        self.assertEqual(opener.call_args.kwargs['timeout'], 20)

    def test_http403_has_clear_eia_error_and_no_fallback(self):
        with patch.object(eia, 'urlopen', side_effect=HTTPError('url', 403, 'Forbidden', {}, None)) as opener:
            with self.assertRaisesRegex(RuntimeError, 'EIA.*403'):
                eia.read_henry_hub()
        self.assertEqual(opener.call_count, 1)

    def test_oversize_response_is_rejected_before_workbook_parse(self):
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, limit): return b'x' * limit
        with patch.object(eia, 'urlopen', return_value=Response()):
            with self.assertRaisesRegex(RuntimeError, 'limit 20 MB'):
                eia.read_henry_hub()


if __name__ == '__main__':
    unittest.main()
