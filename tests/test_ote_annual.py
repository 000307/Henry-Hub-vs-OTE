import importlib.util
import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError
from zipfile import ZipFile, ZIP_DEFLATED

import pandas as pd
from openpyxl import Workbook

spec = importlib.util.spec_from_file_location('ote', Path(__file__).resolve().parents[1] / 'henry_hub_ote_forward.py')
ote = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ote)
YEAR = pd.Timestamp.now(tz="Europe/Prague").year
PAGE_URL = ote.OTE_ANNUAL_SOURCE_TEMPLATE.format(year=YEAR)


def workbook(rows=None, header_height=2, daily_word=False, wrong_sheet=False):
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Other data' if wrong_sheet else 'Odchylky NC BAL'
    sheet.append(['Roční zpráva – test', None, None])
    negative = 'Použitelná cena pro záporné ' + ('denní ' if daily_word else '') + 'vyrovnávací množství'
    positive = 'Použitelná cena pro kladné vyrovnávací množství'
    if header_height == 1:
        sheet.append([negative + ' (Kč/MWh)', positive + ' (Kč/MWh)', 'Plynárenský den'])
    elif header_height == 2:
        sheet.append([negative, positive, 'Plynárenský den'])
        sheet.append(['(Kč/MWh)', '(Kč/MWh)', None])
    else:
        sheet.append(['Použitelná cena pro záporné', 'Použitelná cena pro kladné', 'Plynárenský den'])
        sheet.append(['vyrovnávací množství', 'vyrovnávací množství', None])
        sheet.append(['(Kč/MWh)', '(Kč/MWh)', None])
    if rows is None:
        serial = (datetime(YEAR, 1, 2) - datetime(1899, 12, 30)).days
        rows = [
            ['1 234,56', 999999, datetime(YEAR, 1, 1, 12)],
            [-12.5, 999999, serial], [0, 999999, f'3.1.{YEAR}'],
            ['3,4', 999999, f'{YEAR}-01-04'], ['ND', 999999, f'5.1.{YEAR}'],
            [None, 999999, f'6.1.{YEAR}'], [12, 999999, datetime(YEAR - 1, 12, 31)],
        ]
    for row in rows:
        sheet.append(row)
    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


def archive(files):
    stream = BytesIO()
    with ZipFile(stream, 'w', ZIP_DEFLATED) as z:
        for name, content in files:
            z.writestr(name, content)
    return stream.getvalue()


def gas_zip(content=None, version=0):
    return archive([(f'podsložka/Rocni_zprava_o_trhu_plyn_{YEAR}_V{version}.xlsx', content or workbook())])


class DiscoveryTests(unittest.TestCase):
    def test_prefers_daily_v0_over_lagged_v2_and_ignores_electricity_and_other_year(self):
        html = f'''<h1>Roční zpráva</h1>
        <a href="/pubweb/{YEAR}/Rocni_zprava_o_trhu_{YEAR}_V2_plyn.zip">Plyn V2</a>
        <a HREF="/pubweb/{YEAR}/Rocni_zprava_o_trhu_{YEAR}_V0_plyn.zip" TITLE="Plyn">ZIP</a>
        <a href="/pubweb/{YEAR}/Rocni_zprava_o_trhu_{YEAR}_V0_elektrina.zip">Elektřina</a>
        <a href="/pubweb/{YEAR-1}/Rocni_zprava_o_trhu_{YEAR-1}_V0_plyn.zip">Plyn</a>'''
        result = ote.find_ote_gas_zip_url(html, YEAR, PAGE_URL)
        self.assertEqual(result, f'https://www.ote-cr.cz/pubweb/{YEAR}/Rocni_zprava_o_trhu_{YEAR}_V0_plyn.zip')

    def test_historical_fallback_highest_version_and_unicode_quote(self):
        html = '''<a href="/zprávy/Rocni_zprava_o_trhu_2020_V1_plyn.zip">ZIP</a>
        <a href="/zprávy/Rocni_zprava_o_trhu_2020_V2_plyn.zip">Plyn</a>'''
        result = ote.find_ote_gas_zip_url(html, 2020, PAGE_URL)
        self.assertIn('/zpr%C3%A1vy/', result)
        self.assertTrue(result.endswith('_V2_plyn.zip'))

    def test_ambiguous_v0_is_error(self):
        html = f'<a href="/a/{YEAR}_V0_plyn.zip">Plyn</a><a href="/b/{YEAR}_V0_plyn.zip">Plyn</a>'
        with self.assertRaisesRegex(ValueError, 'nejednoznačných'):
            ote.find_ote_gas_zip_url(html, YEAR, PAGE_URL)

    def test_current_year_without_v0_is_error(self):
        with self.assertRaisesRegex(ValueError, 'verzi V0'):
            ote.find_ote_gas_zip_url(f'<a href="/{YEAR}_V2_plyn.zip">Plyn</a>', YEAR, PAGE_URL)

    def test_external_insecure_and_similar_domains_rejected(self):
        for prefix in ('https://evil.example', 'http://www.ote-cr.cz', 'https://www.ote-cr.cz.evil.example'):
            with self.subTest(prefix=prefix), self.assertRaisesRegex(ValueError, 'HTTPS přímo na ote-cr.cz'):
                ote.find_ote_gas_zip_url(f'<a href="{prefix}/{YEAR}_V0_plyn.zip">Plyn</a>', YEAR, PAGE_URL)


class WorkbookTests(unittest.TestCase):
    def test_actual_zip_xlsx_headers_dates_prices_and_wrong_neighbor_column(self):
        result = ote.parse_ote_annual_zip(gas_zip(), YEAR)
        self.assertEqual(list(result.columns), ['date', 'price_czk_mwh'])
        self.assertEqual(result['date'].dt.strftime('%Y-%m-%d').tolist(),
                         [f'{YEAR}-01-01', f'{YEAR}-01-02', f'{YEAR}-01-03', f'{YEAR}-01-04'])
        self.assertEqual(result['price_czk_mwh'].tolist(), [1234.56, -12.5, 0.0, 3.4])

    def test_one_two_and_three_header_rows_and_optional_daily_word(self):
        for height, daily in ((1, False), (2, False), (3, False), (1, True), (2, True)):
            with self.subTest(height=height, daily=daily):
                result = ote.parse_ote_annual_zip(gas_zip(workbook(header_height=height, daily_word=daily)), YEAR)
                self.assertEqual(result['price_czk_mwh'].tolist(), [1234.56, -12.5, 0.0, 3.4])

    def test_xlsx_selection_prefers_v0_and_ignores_electricity_other_year_and_directory(self):
        actual = workbook()
        decoy = workbook(rows=[[99999, 0, datetime(YEAR, 1, 1)]])
        zipped = archive([
            (f'folder/Rocni_zprava_o_trhu_plyn_{YEAR}_V2.xlsx', decoy),
            (f'folder/Rocni_zprava_o_trhu_plyn_{YEAR}_V0.xlsx', actual),
            (f'Rocni_zprava_o_trhu_elektrina_{YEAR}_V0.xlsx', decoy),
            (f'Rocni_zprava_o_trhu_plyn_{YEAR-1}_V0.xlsx', decoy),
            ('Contents.txt', b'decoy'),
        ])
        self.assertEqual(ote.parse_ote_annual_zip(zipped, YEAR)['price_czk_mwh'].iloc[0], 1234.56)

    def test_conflicting_duplicate_days_error_but_identical_duplicates_merge(self):
        for second, conflict in ((1.0, False), (2.0, True)):
            content = gas_zip(workbook(rows=[[1.0, 999, datetime(YEAR, 1, 1)], [second, 999, f'1.1.{YEAR}']]))
            if conflict:
                with self.assertRaisesRegex(ValueError, 'rozdílné ceny'):
                    ote.parse_ote_annual_zip(content, YEAR)
            else:
                self.assertEqual(len(ote.parse_ote_annual_zip(content, YEAR)), 1)

    def test_positive_column_alone_is_rejected(self):
        wb = Workbook()
        wb.active.title = 'Odchylky NC BAL'
        wb.active.append(['Plynárenský den', 'Použitelná cena pro kladné vyrovnávací množství (Kč/MWh)'])
        wb.active.append([datetime(YEAR, 1, 1), 10])
        stream = BytesIO()
        wb.save(stream)
        with self.assertRaisesRegex(ValueError, 'záporné vyrovnávací'):
            ote.parse_ote_annual_zip(gas_zip(stream.getvalue()), YEAR)

    def test_missing_sheet_non_zip_and_oversize_member_are_errors(self):
        with self.assertRaisesRegex(ValueError, 'Odchylky NC BAL'):
            ote.parse_ote_annual_zip(gas_zip(workbook(wrong_sheet=True)), YEAR)
        with self.assertRaisesRegex(ValueError, 'Archiv nebo sešit'):
            ote.parse_ote_annual_zip(b'<html>Access denied</html>', YEAR)
        huge = archive([(f'Rocni_zprava_o_trhu_plyn_{YEAR}_V0.xlsx', b'0' * 30_000_001)])
        with self.assertRaisesRegex(ValueError, 'příliš velký'):
            ote.parse_ote_annual_zip(huge, YEAR)

    def test_ambiguous_internal_preferred_version_is_error(self):
        zipped = archive([(f'a/plyn_{YEAR}_V0.xlsx', workbook()), (f'b/plyn_{YEAR}_V0.xlsx', workbook())])
        with self.assertRaisesRegex(ValueError, 'nejednoznačných'):
            ote.parse_ote_annual_zip(zipped, YEAR)


class FetchTests(unittest.TestCase):
    def test_fetches_page_then_discovered_zip_and_adds_provenance(self):
        html = f'<h1>Roční zpráva</h1><a href="/pubweb/{YEAR}/plyn_{YEAR}_V0.zip">Plyn ZIP</a>'.encode()
        with patch.object(ote, '_ote_download', side_effect=[html, gas_zip()]) as download:
            result = ote.fetch_ote_annual_prices(YEAR)
        self.assertEqual(download.call_count, 2)
        self.assertEqual(download.call_args_list[0].args[0], PAGE_URL)
        self.assertEqual(result['source_url'].unique().tolist(), [f'https://www.ote-cr.cz/pubweb/{YEAR}/plyn_{YEAR}_V0.zip'])

    def test_only_recognized_annual_page_without_gas_zip_returns_none(self):
        for html in ('<title>Roční zpráva | OTE</title><h1>Roční zpráva</h1>',
                     '<h1>Roční zpráva</h1><a href="/electricity.zip">Elektřina</a>'):
            with patch.object(ote, '_ote_download', return_value=html.encode()):
                self.assertIsNone(ote.fetch_ote_annual_prices(YEAR))
        for html in ('<title>Access Denied</title>', '<html>unknown response</html>',
                     '<title>Access Denied</title><h1>Roční zpráva</h1>'):
            with patch.object(ote, '_ote_download', return_value=html.encode()):
                with self.assertRaisesRegex(ValueError, 'rozpoznanou stránku'):
                    ote.fetch_ote_annual_prices(YEAR)

    def test_http403_is_explicit_and_does_not_try_other_sources(self):
        with patch.object(ote, 'urlopen', side_effect=HTTPError(PAGE_URL, 403, 'Forbidden', {}, None)) as opener:
            with self.assertRaisesRegex(RuntimeError, 'OTE.*403'):
                ote.fetch_ote_annual_prices(YEAR)
        self.assertEqual(opener.call_count, 1)


if __name__ == '__main__':
    unittest.main()
