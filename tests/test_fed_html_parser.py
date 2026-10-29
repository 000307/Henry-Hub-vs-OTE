import importlib.util
import unittest
from email.message import Message
from unittest.mock import patch
from pathlib import Path


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


PRODUCTION = Path(__file__).resolve().parents[1] / 'henry_hub_ote_forward.py'
fixed = load_module('report_html', PRODUCTION)
parse_fixed = fixed.parse_exchange_rates_html


class FedParserRegressions(unittest.TestCase):
    def assert_original_rates(self, html):
        rates = parse_fixed(html)
        self.assertEqual(rates['date'].dt.strftime('%Y-%m-%d').tolist(), ['2000-01-03', '2000-01-06'])
        self.assertEqual(rates['usd_per_eur'].tolist(), [1.0155, 1.2])
        return rates

    def test_optional_cell_and_row_end_tags_preserve_headers_and_quotes(self):
        html = '<table><tr><th>Date<th>Rate<tr><td>3-JAN-00<td>1.0155<tr><td>6-JAN-00<td>1.2</table>'
        self.assert_original_rates(html)

    def test_all_standard_explicit_and_implicit_tag_variants(self):
        rows = [('Date', 'Rate', 'th'), ('3-JAN-00', '1.0155', 'td'),
                ('4-JAN-00', 'ND', 'td'), ('5-JAN-00', '', 'td'), ('6-JAN-00', '1.2', 'td')]
        for close_cell in (False, True):
            for close_row in (False, True):
                for close_table in (False, True):
                    with self.subTest(close_cell=close_cell, close_row=close_row, close_table=close_table):
                        html = '<table>'
                        for day, rate, tag in rows:
                            end_cell = f'</{tag}>' if close_cell else ''
                            html += f'<tr><{tag}>{day}{end_cell}<{tag}>{rate}{end_cell}'
                            html += '</tr>' if close_row else ''
                        html += '</table>' if close_table else ''
                        self.assert_original_rates(html)

    def test_case_attributes_comments_nested_formatting_and_row_groups(self):
        html = '''<TABLE id="fx" data-test="Date Rate"><THEAD><TR>
        <TH scope="col"><SPAN>dAte</SPAN><TH><b>rAtE</b>
        </THEAD><TBODY><TR><TD align="right">3-JAN-00<TD>1.0155
        <!-- ignored <td>missing</td> --><TR><TD>4-JAN-00<TD>ND
        <TR><TD>5-JAN-00<TD>&nbsp;
        </TBODY><TFOOT><TR><TD><FONT>6-JAN-00</FONT><TD>1.2</TFOOT></TABLE>'''
        self.assert_original_rates(html)

    def test_starting_row_group_implicitly_finishes_header_row(self):
        self.assert_original_rates('''<table><thead><tr><th>Date<th>Rate
        <tbody><tr><td>3-JAN-00<td>1.0155<tr><td>6-JAN-00<td>1.2</table>''')

    def test_table_boundary_flushes_last_cell_and_row(self):
        parser = fixed._FedRateTables()
        parser.feed('<table><tr><td>one<td>two</table><table><tr><td>three<td>four')
        parser.close()
        self.assertEqual(parser.tables, [[['one', 'two']], [['three', 'four']]])

    def test_nested_tables_are_separate_and_do_not_contaminate_outer_text(self):
        parser = fixed._FedRateTables()
        parser.feed('<table><tr><td>before<table><tr><td>inner</table>after<td>outer</table>')
        parser.close()
        self.assertEqual(parser.tables, [[['inner']], [['beforeafter', 'outer']]])
        self.assert_original_rates('''<table><tr><td>unrelated<table><tr><th>Date<th>Rate
        <tr><td>3-JAN-00<td>1.0155<tr><td>6-JAN-00<td>1.2</table></table>''')

    def test_script_and_style_text_never_enters_cells(self):
        self.assert_original_rates('''<script>var fake="<table><tr><th>Date<th>Rate";</script>
        <table><tr><th>Date<style>Rate { color:red }</style><th>Rate
        <tr><td>3-JAN-00<td><script>"garbage"</script>1.0155
        <tr><td>6-JAN-00<td>1.2</table>''')

    def test_footnote_markers_on_header_are_narrowly_normalized(self):
        self.assert_original_rates('''<table><tr><th>Date<sup>1</sup><th>Rate&nbsp;*
        <tr><td>3-JAN-00<td>1.0155<tr><td>6-JAN-00<td>1.2</table>''')
        self.assertEqual(fixed._fed_header_label('DATE [1]'), 'date')
        self.assertEqual(fixed._fed_header_label('Rate²'), 'rate')
        self.assertEqual(fixed._fed_header_label('Exchange Rate'), 'exchange rate')

    def test_nd_gap_remains_neighbor_mean_after_parser_fix(self):
        rates = self.assert_original_rates('''<table><tr><th>Date<th>Rate
        <tr><td>3-JAN-00<td>1.0155<tr><td>4-JAN-00<td>ND
        <tr><td>5-JAN-00<td>ND<tr><td>6-JAN-00<td>1.2''')
        result = fixed.resolve_exchange_rates(fixed.pd.Series(['2000-01-04', '2000-01-05']), rates)
        self.assertEqual(result['usd_per_eur'].tolist(), [(1.0155 + 1.2) / 2] * 2)
        self.assertEqual(result['rate_method'].tolist(), ['neighbor_mean'] * 2)

    def test_access_denied_title_is_available_for_diagnostics(self):
        html = '<html><head><title> Access&nbsp; Denied </title></head><body>Request refused</body></html>'
        parser = fixed._FedRateTables()
        parser.feed(html)
        parser.close()
        self.assertEqual(parser.page_title, 'Access Denied')
        self.assertEqual(parser.tables, [])
        with self.assertRaisesRegex(ValueError, 'Date a Rate.*0.*Access Denied'):
            parse_fixed(html)

    def test_http_200_error_page_is_reported_with_downloaded_page_title(self):
        class Response:
            def __init__(self):
                self.headers = Message()
                self.headers['Content-Type'] = 'text/html; charset=utf-8'
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, limit):
                return b'<html><title>Access Denied</title><body>Request refused</body></html>'
        with patch.object(fixed, 'urlopen', return_value=Response()):
            with self.assertRaisesRegex(RuntimeError, 'Date a Rate.*0.*Access Denied'):
                fixed.read_exchange_rates()

    def test_page_title_handles_chunked_feed_and_eof(self):
        for html, expected in [('<title> Historical Rates </title>', 'Historical Rates'),
                               ('<title>Access Denied', 'Access Denied'),
                               ('<title>One</title><script>"<title>Fake</title>"</script>', 'One')]:
            with self.subTest(html=html):
                parser = fixed._FedRateTables()
                for character in html:
                    parser.feed(character)
                parser.close()
                self.assertEqual(parser.page_title, expected)

    def test_streamed_feed_retains_same_result(self):
        html = '<table><tr><th>Date<th>Rate<tr><td>3-JAN-00<td>1.0155</table>'
        parser = fixed._FedRateTables()
        for character in html:
            parser.feed(character)
        parser.close()
        self.assertEqual(parser.tables, [[['Date', 'Rate'], ['3-JAN-00', '1.0155']]])


if __name__ == '__main__':
    unittest.main()
