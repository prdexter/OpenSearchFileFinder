"""Read-only search queries for the existing documents index."""
import html
import re
from datetime import date, timedelta


def _boolean_query(text, fields):
    tokens = re.findall(r'"(?:\\.|[^"\\])*"|[()]|[^\s()"]+', text)
    if re.sub(r'"(?:\\.|[^"\\])*"', '', text).count('"'):
        raise ValueError('Close the quotation marks around your exact phrase.')
    position = 0

    def peek():
        return tokens[position].upper() if position < len(tokens) else None

    def primary():
        nonlocal position
        token = peek()
        if token in ('NOT', '-'):
            position += 1
            return {'bool': {'must_not': [primary()]}}
        if token == '(':
            position += 1
            value = either()
            if peek() != ')':
                raise ValueError('Close the parentheses in your search.')
            position += 1
            return value
        if token is None or token in ('AND', 'OR', ')'):
            raise ValueError('Enter a word or phrase after each search operator.')
        value = tokens[position]
        position += 1
        if value.startswith('-') and len(value) > 1:
            return {'bool': {'must_not': [{'multi_match': {'query': value[1:], 'fields': fields}}]}}
        quoted = value.startswith('"')
        if quoted:
            value = re.sub(r'\\(["\\])', r'\1', value[1:-1])
        if not value.strip():
            raise ValueError('Enter text inside the quotation marks.')
        return {'multi_match': {'query': value, 'fields': fields, 'type': 'phrase' if quoted else 'best_fields'}}

    def both():
        nonlocal position
        parts = [primary()]
        while peek() not in (None, 'OR', ')'):
            if peek() == 'AND':
                position += 1
            parts.append(primary())
        return parts[0] if len(parts) == 1 else {'bool': {'must': parts}}

    def either():
        nonlocal position
        parts = [both()]
        while peek() == 'OR':
            position += 1
            parts.append(both())
        return parts[0] if len(parts) == 1 else {'bool': {'should': parts, 'minimum_should_match': 1}}

    value = either()
    if position != len(tokens):
        raise ValueError('Check the parentheses in your search.')
    return value


def _wildcard_literal(value):
    return ''.join('\\' + character if character in '\\*?' else character for character in value)


def build_search_query(text='', sort_by='relevance', page=1, page_size=100,
                       search_in='both', folder='', file_type='', modified_from='', modified_to='', location='all'):
    fields = {'both': ['file_name^5', 'content^3'], 'contents': ['content'], 'filename': ['file_name']}
    if search_in not in fields or location not in ('all', 'local', 'network'):
        raise ValueError('Choose a valid search scope and location.')
    must = [_boolean_query(text.strip(), fields[search_in])] if text.strip() else [{'match_all': {}}]
    filters = []
    if file_type:
        extensions = [extension.strip().lower().lstrip('.') for extension in file_type.split(',')]
        if any(not re.fullmatch('[a-z0-9]+', extension) for extension in extensions):
            raise ValueError('Choose a valid file type.')
        filters.append({'terms': {'file_type.keyword': extensions}})
    if folder.strip():
        folder = folder.strip().replace('/', '\\').rstrip('\\')
        filters.append({'wildcard': {'file_path': {'value': _wildcard_literal(folder + '\\') + '*', 'case_insensitive': True}}})
    network = {'wildcard': {'file_path': {'value': _wildcard_literal('\\\\') + '*'}}}
    if location == 'network':
        filters.append(network)
    elif location == 'local':
        filters.append({'bool': {'must_not': [network]}})
    dates = {}
    try:
        if modified_from:
            dates['gte'] = date.fromisoformat(modified_from).isoformat()
        if modified_to:
            dates['lt'] = (date.fromisoformat(modified_to) + timedelta(days=1)).isoformat()
        if modified_from and modified_to and modified_from > modified_to:
            raise ValueError('The start date must be on or before the end date.')
    except (ValueError, OverflowError):
        raise ValueError('Use valid modified dates, with the start on or before the end.')
    if dates:
        filters.append({'range': {'modified_date': dates}})
    sorts = {
        'date_desc': [{'modified_date': {'order': 'desc', 'missing': '_last'}}],
        'date_asc': [{'modified_date': {'order': 'asc', 'missing': '_last'}}],
        'name_asc': [{'file_name.keyword': {'order': 'asc', 'missing': '_last'}}],
        'size_desc': [{'file_size': {'order': 'desc', 'missing': '_last'}}],
    }
    return {'query': {'bool': {'must': must, 'filter': filters}},
            'sort': sorts.get(sort_by, ['_score', {'modified_date': 'desc'}]) + [{'file_path': 'asc'}],
            'from': (max(1, page) - 1) * page_size, 'size': page_size, 'track_total_hits': True,
            'highlight': {'encoder': 'html', 'pre_tags': ['<mark>'], 'post_tags': ['</mark>'],
                          'fields': {'content': {'fragment_size': 240, 'number_of_fragments': 2},
                                     'file_name': {'number_of_fragments': 0}}}}


def location_label(path):
    return 'NAS / network' if path.replace('/', '\\').startswith('\\\\') else 'Local'


def result_snippet(hit):
    fragments = hit.get('highlight', {}).get('content', [])
    if fragments:
        return ' … '.join(fragments)
    content = hit.get('_source', {}).get('content') or ''
    return html.escape(content[:300]) + ('…' if len(content) > 300 else '') if content else 'No indexed text available for this document.'
