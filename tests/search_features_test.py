import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from search_features import build_search_query, location_label, result_snippet


class SearchFeaturesTests(unittest.TestCase):
    def test_boolean_precedence_and_exact_phrase(self):
        query = build_search_query('EBV OR "chronic pain" AND intervention NOT pediatric')
        alternatives = query['query']['bool']['must'][0]['bool']['should']
        self.assertEqual(len(alternatives), 2)
        conjunction = alternatives[1]['bool']['must']
        self.assertEqual(conjunction[0]['multi_match']['query'], 'chronic pain')
        self.assertEqual(conjunction[0]['multi_match']['type'], 'phrase')
        self.assertEqual(conjunction[2]['bool']['must_not'][0]['multi_match']['query'], 'pediatric')

    def test_grouped_exclusion_and_filename_scope(self):
        query = build_search_query('protocol NOT (draft OR old)', search_in='filename')
        required = query['query']['bool']['must'][0]['bool']['must']
        self.assertEqual(required[0]['multi_match']['fields'], ['file_name'])
        self.assertEqual(len(required[1]['bool']['must_not'][0]['bool']['should']), 2)

    def test_bad_syntax_has_readable_errors(self):
        for text in ('"unclosed', '(EBV', 'EBV)', 'EBV AND', 'OR EBV', '""'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                build_search_query(text)

    def test_folder_boundary_literal_and_case_insensitive(self):
        query = build_search_query(folder=r'D:\Research?\Project*')
        wildcard = query['query']['bool']['filter'][0]['wildcard']['file_path']
        self.assertTrue(wildcard['case_insensitive'])
        self.assertEqual(wildcard['value'], r'D:\\Research\?\\Project\*\\*')

    def test_dates_include_entire_end_day_and_use_keyword_types(self):
        query = build_search_query(file_type='docx,pdf', modified_from='2026-11-01', modified_to='2026-11-15')
        filters = query['query']['bool']['filter']
        self.assertEqual(filters[0], {'terms': {'file_type.keyword': ['docx', 'pdf']}})
        self.assertEqual(filters[1]['range']['modified_date'], {'gte': '2026-11-01', 'lt': '2026-11-16'})
        with self.assertRaises(ValueError):
            build_search_query(modified_from='2026-11-15', modified_to='2026-11-01')

    def test_blank_search_supports_filters_and_pagination(self):
        query = build_search_query(file_type='docx', page=3, page_size=100, sort_by='date_desc')
        self.assertEqual(query['from'], 200)
        self.assertTrue(query['track_total_hits'])
        self.assertEqual(query['query']['bool']['must'], [{'match_all': {}}])

    def test_network_scope_and_labels_use_indexed_paths(self):
        network = build_search_query(location='network')['query']['bool']['filter'][0]
        local = build_search_query(location='local')['query']['bool']['filter'][0]
        self.assertEqual(local, {'bool': {'must_not': [network]}})
        self.assertEqual(location_label(r'\\NAS\Backups\file.docx'), 'NAS / network')
        self.assertEqual(location_label(r'D:\Backups\file.docx'), 'Local')

    def test_snippets_escape_document_markup(self):
        self.assertEqual(result_snippet({'_source': {'content': '<script>bad()</script>'}}), '&lt;script&gt;bad()&lt;/script&gt;')
        self.assertEqual(build_search_query('text')['highlight']['encoder'], 'html')


if __name__ == '__main__':
    unittest.main()
