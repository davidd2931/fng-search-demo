import functools
import os
import re
import threading
import numpy as np
import pandas as pd
import pyarrow.feather as feather
import demo_ecn
SUPPORTED_FORMAT = '2'
DRAWING_NUMBER_LENGTH = 7
PART_NUMBER_LENGTH = 10
MIN_FILES_FOR_REAL_HEADING = 5
REVISION_MAX_LEVELS = 6
REVISION_MAX_ROUNDS = 12
REVISION_MAX_NODES = 3000
BLOCK_BANNER = 'component information'
ECN_CENTURY_BREAK = 40
CURRENT_YEAR_CEILING = 2100
COLUMN_MAP = [('part_number', '^part\\s*(number|no\\.?|#)$'), ('drawing', '^(dwg|drawing)\\s*(number|no\\.?|#)$'), ('part_name', '^part\\s*name$|^description$|^name$'), ('quantity', '^qty\\s*/?\\s*asm$|^qty\\.?$|^quantity$'), ('level', '^indent$|^level$|^bom\\s*level$'), ('material', "^material\\s*spec\\w*$|^mat'?l\\s*/\\s*finish$|^material$"), ('material_no', "^raw\\s*mat'?l\\s*#$|^mat'?l\\s*#$"), ('finish', '^surface\\s*finish\\s*/?\\s*coating$|^finish$'), ('thickness_min', '(?:.*\\s)?thk[-\\s]*min$'), ('thickness_max', '(?:.*\\s)?thk[-\\s]*max$'), ('weight', '^net\\s*weight(\\s*lbs)?$'), ('commodity_code', '^commodity\\s*code$'), ('manufacturing_location', '^ml$')]
CANONICAL_LABELS = {'part_number': 'Part Number', 'drawing': 'Drawing', 'part_name': 'Part Name', 'quantity': 'Qty', 'level': 'BOM level', 'material': 'Material', 'material_no': 'Material No', 'finish': 'Finish', 'thickness_min': 'Thickness min', 'thickness_max': 'Thickness max', 'weight': 'Weight', 'commodity_code': 'Commodity code', 'manufacturing_location': 'Plant'}
PART_COLUMNS = [('level', 'Lvl'), ('part_number', 'Part Number'), ('drawing', 'Drawing'), ('part_name', 'Part Name'), ('quantity', 'Qty'), ('material', 'Material'), ('thickness_min', 'Thk min'), ('thickness_max', 'Thk max'), ('manufacturing_location', 'Plant')]
COMPONENT_VALUE_COLUMNS = [('part_number', 'Part Number'), ('part_name', 'Part Name'), ('drawing', 'Drawing'), ('material', 'Material'), ('material_no', 'Material No'), ('thickness_min', 'Thk min'), ('thickness_max', 'Thk max'), ('finish', 'Finish'), ('manufacturing_location', 'Plant')]
COMPONENT_COLUMNS = COMPONENT_VALUE_COLUMNS[:1] + [('part_name', 'Part Name'), ('drawing', 'Drawing'), ('material', 'Material'), ('thickness_min', 'Thk min'), ('thickness_max', 'Thk max'), ('manufacturing_location', 'Plant'), ('Folder', 'Commodity'), ('BOMs', 'Used in')]
ATTRIBUTE_SEARCH_COLUMNS = [('BOM #', 'BOM number'), ('Internal Assembly No:', 'Assembly number'), ('Assembly Name:', 'Assembly name'), ('Customer Assembly No:', 'Customer assembly number'), ('Customer & Platform:', 'Customer / platform'), ('Customer Service No:', 'Customer service number'), ('ECN#:', 'ECN'), ('Filename', 'File name')]
SEARCH_SCOPES = [('all', 'Everything', None, None), ('part_no', 'Part number', ('part_number',), ('Internal Assembly No:', 'Customer Assembly No:')), ('part_name', 'Part name', ('part_name',), ('Assembly Name:',)), ('drawing', 'Drawing number', ('drawing',), ()), ('material', 'Material spec', ('material', 'material_no'), ()), ('thickness', 'Material thickness', ('thickness_min', 'thickness_max'), ()), ('bom', 'BOM details', (), ('BOM #', 'ECN#:', 'Customer & Platform:', 'Customer Service No:', 'Filename'))]
SCOPE_LABELS = {key: label for (key, label, _, _) in SEARCH_SCOPES}

def scope_columns(scope):
    if not scope or scope == 'all':
        return (None, None)
    for (key, _, columns, attributes) in SEARCH_SCOPES:
        if key == scope:
            return (columns, attributes)
    return (None, None)
FILE_COLUMNS = [('BOM #', 'BOM #'), ('Internal Assembly No:', 'Assembly No'), ('Assembly Name:', 'Assembly Name'), ('Customer & Platform:', 'Customer / Platform'), ('Customer Assembly No:', 'Customer Asm No'), ('ECN#:', 'ECN'), ('ML', 'ML'), ('Folder', 'Commodity')]
TEMPLATE_NOISE = re.compile('insert\\s+(?:new\\s+)?rows?|last\\s+row|do\\s+not\\s+delete|end\\s+of\\s+bom|ctrl\\s*\\+\\s*shift', re.IGNORECASE)
SIGNER_LABEL = re.compile('^\\s*(?:name|signature)\\s*:?\\s*$', re.IGNORECASE)
DATE_LABEL = re.compile('(?i)^date\\s*:?$')
NOT_A_SIGNER = {'bill of material', 'assembly information', 'component information', 'part name', 'part number', 'net weight', 'assembled parts', 'do not use', 'not applicable', 'see note', 'material spec', 'surface finish', 'drawing number', 'customer assembly', 'internal assembly'}
SIGNER_SHAPE = re.compile("^[A-Za-z][A-Za-z.'-]*(?:\\s+[A-Za-z][A-Za-z.'-]*){1,2}$")

def looks_like_a_signer(value):
    text = ' '.join(str(value).split())
    if not 6 <= len(text) <= 34:
        return False
    if text.lower() in NOT_A_SIGNER:
        return False
    return bool(SIGNER_SHAPE.match(text))
NOT_A_PART = {'name', 'date', 'signature', 'n/a', 'na', '-', 'none', 'prepared by', 'approved by', 'checked by', 'rev', 'note'}
EXCEL_EPOCH = pd.Timestamp('1899-12-30')
DATE_SERIAL = re.compile('^\\d{5}(\\.\\d+)?$')
TEXT_DATE = re.compile('^([A-Za-z]{3})-(\\d{1,2})-(\\d{2})$')
MONTHS = {name: number for (number, name) in enumerate(('jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'), start=1)}

def read_date(value):
    text = str(value).strip()
    if DATE_SERIAL.match(text):
        try:
            when = EXCEL_EPOCH + pd.Timedelta(days=float(text))
        except (ValueError, OverflowError):
            return None
        return when if 1990 <= when.year <= 2030 else None
    match = TEXT_DATE.match(text)
    if match:
        month = MONTHS.get(match.group(1).lower())
        if not month:
            return None
        (day, year) = (int(match.group(2)), int(match.group(3)))
        year += 2000 if year <= 30 else 1900
        try:
            return pd.Timestamp(year=year, month=month, day=day)
        except ValueError:
            return None
    return None

def ecn_year(value):
    text = str(value).strip()
    match = re.match('^(\\d{2})\\d{3}', text)
    if not match:
        return ''
    two = int(match.group(1))
    year = 2000 + two if two <= ECN_CENTURY_BREAK else 1900 + two
    if year > CURRENT_YEAR_CEILING:
        return ''
    return str(year)

@functools.lru_cache(maxsize=1 << 16, typed=True)
def looks_like_part_number(value):
    text = str(value).strip()
    if not text or text.lower() in NOT_A_PART:
        return False
    if text.endswith(':'):
        return False
    if TEMPLATE_NOISE.search(text):
        return False
    if text.isdigit() and len(text) < 6:
        return False
    return len(text) >= 6

def ecn_key(value):
    match = re.match('\\s*(\\d{5})', str(value))
    return match.group(1) if match else ''

def load_projects(path):
    if not path or not os.path.exists(path):
        return {}
    try:
        frame = feather.read_table(path).to_pandas()
    except Exception:
        return {}
    if 'ECN' not in frame.columns or 'Project' not in frame.columns:
        return {}
    projects = {}
    for (_, row) in frame.iterrows():
        key = ecn_key(row['ECN'])
        project = str(row['Project']).strip()
        if not key or not project:
            continue
        confidence = str(row.get('Confidence', '')).strip()
        existing = projects.get(key)
        if existing and existing['confidence'] == 'certain' and (confidence != 'certain'):
            continue
        projects[key] = {'project': project, 'confidence': confidence, 'source': str(row.get('Source', ''))}
    return projects

def load_project_folders(path):
    if not path or not os.path.exists(path):
        return {}
    try:
        frame = feather.read_table(path).to_pandas()
    except Exception:
        return {}
    if 'Project' not in frame.columns or 'Path' not in frame.columns:
        return {}
    folders = {}
    for (_, row) in frame.iterrows():
        number = str(row['Project']).strip()
        path_text = str(row['Path']).strip()
        if not number or not path_text:
            continue
        status = str(row.get('Status', '')).strip()
        found = folders.setdefault(number, [])
        entry = {'path': path_text, 'name': str(row.get('Name', '')), 'status': status}
        if status == 'active':
            found.insert(0, entry)
        else:
            found.append(entry)
    return folders
DFMEA_TABLE_FORMAT = '1'

def _revision_rank(revision):
    text = str(revision or '').upper()
    if not text:
        return (0, ())
    return (1, (len(text),) + tuple((ord(c) for c in text)))

def load_dfmea(path):
    if not path or not os.path.exists(path):
        return None
    try:
        table = feather.read_table(path)
    except Exception:
        return None
    metadata = {key.decode(): value.decode() for (key, value) in (table.schema.metadata or {}).items()}
    if metadata.get('demo_format_version') != DFMEA_TABLE_FORMAT:
        return None
    frame = table.to_pandas()
    needed = ('Folder', 'Filename', 'Numbers', 'Revision', 'Modified')
    if any((column not in frame.columns for column in needed)):
        return None
    found = {}
    for (folder, filename, numbers, revision, modified) in zip(frame['Folder'], frame['Filename'], frame['Numbers'], frame['Revision'], frame['Modified']):
        revision = str(revision or '').strip().upper()
        record = {'folder': str(folder), 'filename': str(filename), 'revision': revision, 'next': demo_ecn.next_revision(revision) if revision else '', 'modified': str(modified)}
        for number in str(numbers).split():
            held = found.get(number)
            if held is None or (_revision_rank(revision), record['modified']) > (_revision_rank(held['revision']), held['modified']):
                found[number] = record
    return found
REVISION_FAMILY = re.compile('\\s*(\\d{7})(\\d)\\d\\d(?!\\d)')

def revision_family(value):
    match = REVISION_FAMILY.match(str(value))
    return match.group(1) + match.group(2) if match else ''

def drawing_owners(numbers):
    grouped = {}
    for value in numbers:
        number = str(value).strip()
        if len(number) == PART_NUMBER_LENGTH and number.isdigit():
            grouped.setdefault(number[:DRAWING_NUMBER_LENGTH], []).append(number)
        else:
            grouped.setdefault(number, [number])
    owners = {}
    for group in grouped.values():
        owner = min(group)
        for number in group:
            owners[number] = owner
    return owners

def restore_leading_zero(text, width):
    if text.isdigit() and len(text) == width - 1 and (not text.startswith('0')):
        return text.zfill(width)
    return text
_mark_re = re.compile("^['*]+\\s*(?=\\d)")

def strip_mark(value):
    return _mark_re.sub('', str(value).strip())

def part_spelling(value):
    return restore_leading_zero(strip_mark(value), PART_NUMBER_LENGTH)

@functools.lru_cache(maxsize=1 << 16, typed=True)
def part_number_key(value):
    text = strip_mark(value)
    if not text:
        return ''
    token = text.split()[0].strip(',;')
    if len(token) < 6 or not any((character.isdigit() for character in token)):
        return ''
    token = restore_leading_zero(token, PART_NUMBER_LENGTH)
    return token if looks_like_part_number(token) else ''

def as_text(series):
    if not isinstance(series.dtype, pd.CategoricalDtype):
        return series.astype(str)
    categories = np.asarray(series.cat.categories.astype(str), dtype=object)
    codes = series.cat.codes.to_numpy()
    text = categories[codes] if len(categories) else np.full(len(codes), 'nan', dtype=object)
    if (codes < 0).any():
        text = np.where(codes < 0, 'nan', text)
    return pd.Series(text, index=series.index, name=series.name)

def text_matches(series, pattern):
    if not isinstance(series.dtype, pd.CategoricalDtype):
        return np.array([bool(pattern.match(str(value))) for value in series], dtype=bool)
    hits = np.array([bool(pattern.match(str(value))) for value in series.cat.categories] + [False], dtype=bool)
    return hits[series.cat.codes.to_numpy()]

def _codes(series):
    if isinstance(series.dtype, pd.CategoricalDtype):
        return series.cat.codes.to_numpy().astype(np.int64) + 1
    return pd.factorize(series)[0].astype(np.int64) + 1
_normalise_re = re.compile('\\s+')

def normalise_heading(text):
    return _normalise_re.sub(' ', str(text).strip().lower())

def canonical_column(heading):
    name = normalise_heading(heading)
    for (canonical, pattern) in COLUMN_MAP:
        if re.match(pattern, name):
            return canonical
    return None

class IndexError_(Exception):
    pass

class Index:

    def __init__(self, cells, files, pdfs, info):
        self.cells = cells
        for column in ('BOM #', 'Internal Assembly No:'):
            if column in files.columns:
                files = files.assign(**{column: files[column].map(lambda value: strip_mark(value) if isinstance(value, str) else value)})
        self.files = files
        self.pdfs = pdfs
        self.info = info
        headings = cells['ColHeader'].astype(str)
        self._canonical = {h: canonical_column(h) for h in headings.unique()}
        self.cells = cells.assign(Canonical=headings.map(self._canonical).fillna(''))
        self._stack_split_headings()
        self._prepare()

    def _stack_split_headings(self):
        cells = self.cells
        header = cells[cells['IsHeader'].to_numpy(dtype=bool)][['Filename', 'Sheet', 'Row', 'Col', 'Value']]
        if header.empty:
            return
        header = header.sort_values(['Filename', 'Sheet', 'Col', 'Row'], kind='stable')
        group = header.groupby(['Filename', 'Sheet', 'Col'], sort=False, observed=True).ngroup().to_numpy()
        rows = header['Row'].to_numpy()
        starts = np.empty(len(header), dtype=bool)
        starts[0] = True
        starts[1:] = (group[1:] != group[:-1]) | (rows[1:] != rows[:-1] + 1)
        first = np.flatnonzero(starts)
        last = np.append(first[1:], len(header))
        values = header['Value'].astype(object).to_numpy()
        bottom = values[last - 1]
        file_codes = header['Filename'].cat.codes.to_numpy()
        sheet_codes = header['Sheet'].cat.codes.to_numpy()
        columns = header['Col'].to_numpy()
        heading_codes = {text: code for (code, text) in enumerate(cells['ColHeader'].cat.categories)}
        remembered = dict(self._canonical)

        def canonical_of(text):
            if text not in remembered:
                remembered[text] = canonical_column(text)
            return remembered[text]
        stacked = {}
        worth_stacking = last - first > 1
        for position in np.flatnonzero(worth_stacking):
            bottom_text = str(bottom[position])
            if canonical_of(bottom_text) is not None:
                continue
            joined = ' '.join((text for text in (str(value) for value in values[first[position]:last[position]]) if text.strip()))
            canonical = canonical_of(joined)
            if canonical is None:
                continue
            heading_code = heading_codes.get(bottom_text)
            if heading_code is None:
                continue
            stacked[file_codes[first[position]], sheet_codes[first[position]], columns[first[position]], heading_code] = canonical
        self._stacked_headings = stacked
        if not stacked:
            return
        rescued = np.zeros(len(cells['ColHeader'].cat.categories) + 1, dtype=bool)
        for key in stacked:
            rescued[key[3]] = True
        wanted = rescued[cells['ColHeader'].cat.codes.to_numpy()] & (cells['Canonical'].to_numpy() == '')
        positions = np.flatnonzero(wanted)
        if len(positions) == 0:
            return
        keys = zip(cells['Filename'].cat.codes.to_numpy()[positions], cells['Sheet'].cat.codes.to_numpy()[positions], cells['Col'].to_numpy()[positions], cells['ColHeader'].cat.codes.to_numpy()[positions])
        canonical_values = cells['Canonical'].to_numpy(dtype=object).copy()
        canonical_values[positions] = [stacked.get(key, '') for key in keys]
        self.cells = self.cells.assign(Canonical=canonical_values)

    def _prepare(self):
        headed = self.cells[self.cells['ColHeader'] != '']
        per_heading = headed.groupby('ColHeader', observed=True)['Filename'].nunique()
        self._real_headings = set(per_heading[per_heading >= MIN_FILES_FOR_REAL_HEADING].index)
        self._attribute_cache = {}
        self._numbers = None
        self._drawing_by_prefix = {}
        for name in self.pdfs['PDF_Name'].astype(str):
            for run in re.findall('\\d+', name):
                if len(run) >= DRAWING_NUMBER_LENGTH:
                    self._drawing_by_prefix.setdefault(run[:DRAWING_NUMBER_LENGTH], name)
        self._components = None
        self._components_lock = threading.Lock()
        self._uses = None
        self._parents = None
        self._part_names = None
        self._details = None
        self.projects = {}
        self.project_folders = {}
        self.obsolete = pd.DataFrame()
        self.dfmea = None
        self.cad = pd.DataFrame(columns=['Folder', 'Name'])
        self.prototype = None
        self.prototype_error = ''
        self.sampled = set()
        self._dates = None
        self._signers = None
        self._plants = None
        self._part_plants = None
        self._assemblies = None
        self._layouts = None
        self._layout_lock = threading.RLock()

    @property
    def built_at(self):
        return self.info.get('demo_built_at', 'unknown')

    @property
    def workbook_count(self):
        return len(self.files)

    def age_in_days(self):
        try:
            built = pd.to_datetime(self.built_at)
        except (ValueError, TypeError):
            return None
        return (pd.Timestamp.now() - built).days

    def built_on(self):
        try:
            return pd.to_datetime(self.built_at).strftime('%d %b').lstrip('0')
        except (ValueError, TypeError):
            return ''

    def describe(self):
        days = self.age_in_days()
        if days is None:
            when = ''
        elif days == 0:
            when = ' Built today.'
        elif days == 1:
            when = ' Built yesterday.'
        else:
            when = f' Built {days} days ago.'
        return f'{self.workbook_count:,} BOMs indexed.{when}'

    def search(self, term, limit=None, scope=None):
        term = term.strip()
        if not term:
            return pd.DataFrame()
        (columns, attributes) = scope_columns(scope)
        matches = self.cells[self._matching_rows(term)]
        if columns is not None:
            matches = matches[matches['Canonical'].astype(str).isin(columns)]
        attribute_matches = self._matching_attributes(term, exclude=set(matches['Filename'].astype(str)) if not matches.empty else set(), columns=attributes)
        if matches.empty and attribute_matches.empty:
            return pd.DataFrame()
        counts = matches.groupby('Filename', observed=True).size() if not matches.empty else pd.Series(dtype=int)
        first = matches.drop_duplicates(subset=['Filename']) if not matches.empty else matches
        if not attribute_matches.empty:
            first = pd.concat([first, attribute_matches], ignore_index=True)
        results = first.merge(self.files, on=['Folder', 'Filename'], how='left')
        results['Matches'] = results['Filename'].map(counts).fillna(1).astype(int)
        results['MatchedIn'] = [self.label_for(canonical, heading) for (canonical, heading) in zip(results['Canonical'].astype(str), results['ColHeader'].astype(str))]
        if '_AttributeLabel' in results.columns:
            from_attribute = results['_AttributeLabel'].astype(str) != ''
            results.loc[from_attribute, 'MatchedIn'] = results.loc[from_attribute, '_AttributeLabel']
            results = results.drop(columns=['_AttributeLabel'])
        results = results.rename(columns={'Value': 'MatchedValue'})
        results['MatchedValue'] = results['MatchedValue'].astype(str)
        own_number = results.get('BOM #', pd.Series(dtype=str)).astype(str)
        is_own = (results['MatchedIn'] == 'elsewhere in the BOM') & (results['MatchedValue'].map(strip_mark) == own_number)
        results.loc[is_own, 'MatchedIn'] = 'BOM number'
        results['Drawing'] = [self.drawing_for(number) for number in results.get('BOM #', pd.Series(dtype=str))]
        results = results.sort_values(['Folder', 'Filename'])
        return results.head(limit) if limit else results

    def _matching_attributes(self, term, exclude, columns=None):
        frame = self.files
        if columns is not None and (not columns):
            return pd.DataFrame()
        hit = pd.Series(False, index=frame.index)
        matched_in = pd.Series('', index=frame.index)
        matched_value = pd.Series('', index=frame.index)
        for (column, label) in ATTRIBUTE_SEARCH_COLUMNS:
            if column not in frame.columns:
                continue
            if columns is not None and column not in columns:
                continue
            values = self._attribute_text(column)
            found = values.str.contains(term, case=False, na=False, regex=False)
            fresh = found & ~hit
            matched_in[fresh] = label
            matched_value[fresh] = values[fresh]
            hit = hit | found
        if not hit.any():
            return pd.DataFrame()
        rows = frame[hit][['Folder', 'Filename']].copy()
        rows['Value'] = matched_value[hit]
        rows['ColHeader'] = ''
        rows['Canonical'] = ''
        rows['_AttributeLabel'] = matched_in[hit]
        rows = rows[~rows['Filename'].astype(str).isin(exclude)]
        return rows

    def _attribute_text(self, column):
        text = self._attribute_cache.get(column)
        if text is None:
            text = self.files[column].astype(str)
            self._attribute_cache[column] = text
        return text

    def _matching_rows(self, term):
        values = self.cells['Value']
        categorical = isinstance(values.dtype, pd.CategoricalDtype)
        pool = values.cat.categories if categorical else values.astype(str)

        def spread(matched):
            return np.asarray(matched)[values.cat.codes.to_numpy()] if categorical else np.asarray(matched)
        mask = spread(pool.str.contains(term, case=False, na=False, regex=False))
        self.zero_padded_hits = 0
        stripped = term.lstrip('0')
        if term.isdigit() and stripped and (stripped != term):
            text = pd.Series(pool.astype(str).to_numpy() if hasattr(pool, 'to_numpy') else pool.astype(str))
            digits = text.str.fullmatch('\\d+').fillna(False)
            short = digits & ((text == stripped) | text.str.startswith(stripped) & (text.str.len() == PART_NUMBER_LENGTH - 1))
            extra = spread(short) & ~mask
            self.zero_padded_hits = int(extra.sum())
            mask = mask | extra
        return mask

    def label_for(self, canonical, heading):
        if canonical:
            return CANONICAL_LABELS.get(canonical, canonical.replace('_', ' '))
        heading = str(heading).strip()
        if heading in self._real_headings and len(heading) <= 28:
            return heading
        return 'elsewhere in the BOM'

    def components(self):
        if self._components is not None:
            return self._components
        with self._components_lock:
            if self._components is None:
                self._components = self._build_components()
        return self._components

    def _build_components(self):
        rows = self.cells[~self.cells['IsHeader'].astype(bool) & (self.cells['Canonical'].astype(str) != '')]
        if rows.empty:
            return pd.DataFrame()
        rows = rows.assign(Value=as_text(rows['Value']))
        table = rows.pivot_table(index=['Folder', 'Filename', 'Sheet', 'Row'], columns='Canonical', values='Value', aggfunc='first', observed=True).reset_index()
        if 'part_number' not in table.columns:
            return pd.DataFrame()
        table = table[table['part_number'].notna()]
        table = table.assign(part_number=table['part_number'].map(part_spelling))
        table = table[table['part_number'].map(looks_like_part_number)]
        table = table[table['part_number'].map(part_number_key) != '']
        if table.empty:
            return pd.DataFrame()
        for (canonical, _) in COMPONENT_VALUE_COLUMNS:
            if canonical not in table.columns:
                table[canonical] = ''
        table = table.astype(object)
        table = table.where(table.notna(), '')
        grouped = table.groupby(['part_number', 'Folder'], observed=True)
        merged = {}
        for (canonical, _) in COMPONENT_VALUE_COLUMNS:
            if canonical == 'part_number':
                continue
            values = table[canonical].astype(str)
            ranked = table.assign(_v=values, _blank=values.str.strip() == '')
            ranked = ranked.sort_values('_blank', kind='stable')
            merged[canonical] = ranked.groupby(['part_number', 'Folder'], observed=True)['_v'].first()
        components = pd.DataFrame(merged).reset_index()
        components['BOMs'] = grouped['Filename'].nunique().values
        components['Filename'] = grouped['Filename'].first().values
        components['drawing'] = [self.drawing_spelling(value, part) for (value, part) in zip(components['drawing'], components['part_number'])]
        blank = components['drawing'].astype(str).str.strip() == ''
        components.loc[blank, 'drawing'] = [self.drawing_for(number) for number in components.loc[blank, 'part_number']]
        components = components.sort_values(['Folder', 'part_number'])
        return components.reset_index(drop=True)

    def search_components(self, term, limit=None, scope=None):
        term = term.strip()
        if not term:
            return pd.DataFrame()
        components = self.components()
        if components.empty:
            return pd.DataFrame()
        (columns, _) = scope_columns(scope)
        searchable = [name for (name, _) in COMPONENT_VALUE_COLUMNS]
        if columns:
            narrowed = [name for name in columns if name in searchable]
            if narrowed:
                searchable = narrowed
        hit = pd.Series(False, index=components.index)
        matched_in = pd.Series('', index=components.index)
        matched_value = pd.Series('', index=components.index)
        for name in searchable:
            values = components[name].astype(str)
            found = values.str.contains(term, case=False, na=False, regex=False)
            fresh = found & ~hit
            matched_in[fresh] = CANONICAL_LABELS.get(name, name)
            matched_value[fresh] = values[fresh]
            hit = hit | found
        if not hit.any():
            return pd.DataFrame()
        results = components[hit].copy()
        results['MatchedIn'] = matched_in[hit]
        results['MatchedValue'] = matched_value[hit]
        results = results.sort_values(['Folder', 'part_number'])
        return results.head(limit) if limit else results

    def resolve_part_numbers(self, part_number):
        number = str(part_number).strip()
        uses = self._part_uses()
        if not number or number in uses:
            return [number] if number else []
        if len(number) == DRAWING_NUMBER_LENGTH and number.isdigit():
            return sorted((known for known in uses if known[:DRAWING_NUMBER_LENGTH] == number))
        return [number]

    def revision_impact(self, part_number, max_rounds=REVISION_MAX_ROUNDS):
        seeds = self.resolve_part_numbers(part_number)
        if not seeds:
            return pd.DataFrame()
        uses = self._part_uses()
        parents = self._parents_by_file()
        revving = {seed: 0 for seed in seeds}
        affected = {}
        frontier = list(seeds)
        for round_number in range(max_rounds):
            spawned = []
            for part in frontier:
                for filename in uses.get(part, ()):
                    chain = self._chain_in_file(filename, part, parents)
                    if filename not in affected:
                        why = 'uses it directly' if round_number == 0 else f'lists {part}, which revs'
                        steps = ' → '.join([part] + chain)
                        affected[filename] = (round_number, part, why, steps)
                    for ancestor in chain:
                        if ancestor not in revving:
                            revving[ancestor] = round_number + 1
                            spawned.append(ancestor)
            if not spawned:
                break
            frontier = spawned
        if not affected:
            return pd.DataFrame()
        frame = pd.DataFrame([(level, filename, part, why, steps) for (filename, (level, part, why, steps)) in affected.items()], columns=['Level', 'Filename', 'RevPart', 'MatchedIn', 'MatchedValue'])
        frame = frame.merge(self.files, on='Filename', how='left')
        frame['Drawing'] = [self.drawing_for(value) for value in frame.get('BOM #', pd.Series(dtype=str))]
        return frame.sort_values(['Level', 'Folder', 'Filename']).reset_index(drop=True)

    def revision_impact_with_siblings(self, part_number, max_rounds=REVISION_MAX_ROUNDS):
        siblings = self.siblings_sharing_drawing(part_number)
        frames = [self.revision_impact(part_number, max_rounds)]
        frames += [self.revision_impact(sibling, max_rounds) for sibling in siblings]
        frames = [frame for frame in frames if not frame.empty]
        if not frames:
            return (pd.DataFrame(), siblings)
        joined = pd.concat(frames, ignore_index=True).drop_duplicates('Filename').sort_values(['Level', 'Folder', 'Filename']).reset_index(drop=True)
        return (joined, siblings)

    def revision_parts(self, part_number, max_rounds=REVISION_MAX_ROUNDS):
        seeds = self.resolve_part_numbers(part_number)
        if not seeds:
            return {}
        (uses, parents) = (self._part_uses(), self._parents_by_file())
        (revving, frontier) = ({seed: 0 for seed in seeds}, list(seeds))
        for round_number in range(max_rounds):
            spawned = []
            for part in frontier:
                for filename in uses.get(part, ()):
                    for ancestor in self._chain_in_file(filename, part, parents):
                        if ancestor not in revving:
                            revving[ancestor] = round_number + 1
                            spawned.append(ancestor)
            if not spawned:
                break
            frontier = spawned
        return revving

    def bom_dates(self):
        if self._dates is not None:
            return self._dates
        is_label = text_matches(self.cells['Value'], DATE_LABEL)
        labels = self.cells[is_label]
        wanted = {(str(filename), str(sheet), int(row), int(column) + 1) for (filename, sheet, row, column) in zip(labels['Filename'], labels['Sheet'], labels['Row'], labels['Col'])}
        if not wanted:
            self._dates = {}
            return self._dates
        near = self._rows_holding(is_label)
        found = {}
        for (filename, sheet, row, column, value) in zip(as_text(near['Filename']), as_text(near['Sheet']), near['Row'], near['Col'], as_text(near['Value'])):
            if (filename, sheet, int(row), int(column)) not in wanted:
                continue
            when = read_date(value)
            if when is not None and (filename not in found or when > found[filename]):
                found[filename] = when
        self._dates = found
        return self._dates

    def bom_signers(self):
        if self._signers is not None:
            return self._signers
        near = self._rows_holding(text_matches(self.cells['Value'], DATE_LABEL) | text_matches(self.cells['Value'], SIGNER_LABEL))
        values = as_text(near['Value'])
        files = as_text(near['Filename'])
        sheets = as_text(near['Sheet'])
        rows = near['Row']
        columns = near['Col']
        at = {}
        for (filename, sheet, row, column, value) in zip(files, sheets, rows, columns, values):
            at[filename, sheet, int(row), int(column)] = value
        date_rows = {}
        for (filename, sheet, row, column, value) in zip(files, sheets, rows, columns, values):
            if not DATE_LABEL.match(str(value)):
                continue
            when = read_date(at.get((filename, sheet, int(row), int(column) + 1), ''))
            date_rows[filename, sheet, int(row)] = when
        best = {}
        for (filename, sheet, row, column, value) in zip(files, sheets, rows, columns, values):
            if not SIGNER_LABEL.match(str(value)):
                continue
            (row, column) = (int(row), int(column))
            (when, distance) = (None, None)
            for step in (0, -1, 1, -2, 2):
                if (filename, sheet, row + step) in date_rows:
                    when = date_rows[filename, sheet, row + step]
                    distance = abs(step)
                    break
            if distance is None:
                continue
            for step in range(1, 4):
                beside = at.get((filename, sheet, row, column + step), '')
                text = ' '.join(str(beside).split())
                if not text:
                    continue
                if looks_like_a_signer(text):
                    previous = best.get(filename)
                    if previous is None or (when is not None and (previous[1] is None or when > previous[1])):
                        best[filename] = (text, when)
                break
        self._signers = {name: signer for (name, (signer, _)) in best.items()}
        return self._signers

    def _rows_holding(self, mask):
        cells = self.cells
        sheets = _codes(cells['Sheet'])
        rows = cells['Row'].to_numpy().astype(np.int64)
        key = (_codes(cells['Filename']) * (int(sheets.max()) + 1) + sheets) * (int(rows.max()) + 1) + rows
        return cells[np.isin(key, key[mask])]

    def bom_signer(self, filename):
        return self.bom_signers().get(str(filename), '')

    def bom_plants(self):
        if self._plants is not None:
            return self._plants
        (body, canonical, written) = self._plant_lines()
        assemblies = self._assembly_numbers()
        numbers = body[canonical == 'part_number']
        found = {}
        for (filename, sheet, row, value) in zip(as_text(numbers['Filename']), as_text(numbers['Sheet']), numbers['Row'], numbers['Value']):
            if filename in found:
                continue
            plant = written.get((filename, sheet, int(row)))
            if plant and part_number_key(str(value)) == assemblies.get(filename):
                found[filename] = plant
        self._plants = found
        return self._plants

    def bom_plant(self, filename):
        return self.bom_plants().get(str(filename), '')

    def _plant_lines(self):
        body = self.cells[~self.cells['IsHeader'].astype(bool)]
        canonical = body['Canonical'].astype(str)
        plants = body[canonical == 'manufacturing_location']
        written = {}
        for (filename, sheet, row, value) in zip(as_text(plants['Filename']), as_text(plants['Sheet']), plants['Row'], plants['Value']):
            text = ' '.join(str(value).split())
            if text:
                written[filename, sheet, int(row)] = text
        return (body, canonical, written)

    def part_plants(self, part_number):
        if self._part_plants is None:
            (body, canonical, written) = self._plant_lines()
            numbers = body[canonical == 'part_number']
            found = {}
            for (filename, sheet, row, value) in zip(as_text(numbers['Filename']), as_text(numbers['Sheet']), numbers['Row'], numbers['Value']):
                plant = written.get((filename, sheet, int(row)))
                key = part_number_key(str(value))
                if plant and key:
                    held = found.setdefault(key, [])
                    if plant not in held:
                        held.append(plant)
            self._part_plants = found
        return list(self._part_plants.get(part_number_key(str(part_number or '')), []))

    def prototype_state(self):
        return 'ready' if getattr(self, 'prototype', None) is not None else 'not indexed'

    def dfmea_state(self):
        return 'ready' if getattr(self, 'dfmea', None) is not None else 'not built'

    def dfmea_for(self, number):
        table = getattr(self, 'dfmea', None)
        if not table:
            return None
        key = part_number_key(str(number or ''))
        return table.get(key) or table.get(str(number or '').strip())

    def dfmea_for_bom(self, assembly, bom_number):
        if part_number_key(str(assembly or '')):
            return self.dfmea_for(assembly)
        bare = re.sub('-.*$', '', str(bom_number or '').strip())
        return self.dfmea_for(bare) if bare else None

    def bom_date(self, filename, ecn=''):
        when = self.bom_dates().get(str(filename))
        if when is not None:
            return when.strftime('%Y-%m-%d')
        return ecn_year(ecn)

    def project_for(self, ecn):
        found = self.projects.get(ecn_key(ecn))
        return found['project'] if found else ''

    def project_folder_for(self, ecn):
        project = self.project_for(ecn)
        if not project:
            return None
        found = self.project_folders.get(project)
        return found[0] if found else None

    def project_folders_for(self, ecn):
        project = self.project_for(ecn)
        return list(self.project_folders.get(project, ())) if project else []

    def project_detail(self, ecn):
        return self.projects.get(ecn_key(ecn))

    def projects_for_assembly(self, assembly_number):
        wanted = str(assembly_number).strip()
        if not wanted:
            return []
        sources = []
        assemblies = self.files.get('Internal Assembly No:')
        if assemblies is not None:
            live = self.files[assemblies.astype(str).str.strip() == wanted]
            for (_, row) in live.iterrows():
                sources.append((True, str(row.get('BOM #', '')), str(row.get('ECN#:', '')), str(row.get('Filename', '')), str(row.get('Folder', '')), ''))
        if isinstance(self.obsolete, pd.DataFrame) and (not self.obsolete.empty):
            assemblies = self.obsolete['Assembly'].astype(str).str.strip()
            family = revision_family(wanted)
            if family:
                older = self.obsolete[assemblies.map(revision_family) == family]
            else:
                older = self.obsolete[assemblies == wanted]
            for (_, row) in older.sort_values('Filename', ascending=False).iterrows():
                sources.append((False, str(row.get('BOM #', '') or row.get('Filename', '')), str(row.get('ECN#:', '') or row.get('ECN', '')), str(row.get('Filename', '')), str(row.get('Folder', '')), str(row.get('Path', '') or '')))
        (found, order) = ({}, [])
        for (current, bom, ecn, filename, folder, path) in sources:
            project = self.project_for(ecn)
            if not project:
                continue
            bom_file = {'label': bom, 'folder': folder, 'filename': filename, 'path': path}
            if project in found:
                found[project]['bom'] += ', ' + bom
                found[project]['bom_files'].append(bom_file)
                found[project]['current'] = found[project]['current'] or current
                continue
            folders = self.project_folders.get(project, ())
            found[project] = {'project': project, 'name': folders[0]['name'] if folders else '', 'path': folders[0]['path'] if folders else '', 'folders': list(folders), 'ecn': ecn, 'bom': bom, 'bom_files': [bom_file], 'current': current}
            order.append(project)
        return sorted((found[project] for project in order), key=lambda entry: not entry['current'])

    def part_name(self, part_number):
        if self._part_names is None:
            self._parents_by_file()
        return self._part_names.get(str(part_number).strip(), '')

    def revision_tree(self, part_number, max_nodes=REVISION_MAX_NODES, info=None):
        seeds = self.resolve_part_numbers(part_number)
        if info is not None:
            info['truncated'] = False
        if not seeds:
            return []
        uses = self._part_uses()
        parents = self._parents_by_file()
        details = self._file_details()
        assemblies = self.assembly_numbers()
        budget = [max_nodes]
        expanded = set()

        def out_of_budget():
            if budget[0] <= 0 and info is not None:
                info['truncated'] = True
            return budget[0] <= 0

        def part_node(part, skip_file=None, depth=0):
            node = {'kind': 'part', 'part': part, 'name': self.part_name(part), 'boms': len(uses.get(part, ())), 'repeat': False, 'children': []}
            if part in expanded:
                node['repeat'] = True
                return node
            if depth >= REVISION_MAX_LEVELS or out_of_budget():
                return node
            expanded.add(part)
            for filename in sorted(uses.get(part, ())):
                if filename == skip_file or out_of_budget():
                    continue
                budget[0] -= 1
                node['children'].append(bom_node(filename, part, depth))
            return node

        def climb_from_assembly(node, assembly, filename, depth):
            if assembly in expanded:
                node['repeat'] = True
                return
            if depth + 1 >= REVISION_MAX_LEVELS or out_of_budget():
                return
            expanded.add(assembly)
            for other in sorted(uses.get(assembly, ())):
                if other == filename or out_of_budget():
                    continue
                budget[0] -= 1
                node['children'].append(bom_node(other, assembly, depth + 1))

        def bom_node(filename, part, depth):
            node = dict(details.get(filename, {'Folder': '', 'BOM #': ''}))
            node.update({'kind': 'bom', 'filename': filename, 'assembly': assemblies.get(filename, ''), 'children': []})
            holder = node
            for ancestor in self._chain_in_file(filename, part, parents):
                if ancestor == node['assembly']:
                    climb_from_assembly(node, ancestor, filename, depth)
                    break
                if out_of_budget():
                    break
                budget[0] -= 1
                child = part_node(ancestor, skip_file=filename, depth=depth + 1)
                holder['children'].append(child)
                holder = child
            return node
        return [part_node(seed) for seed in seeds]

    def _file_details(self):
        if self._details is None:
            wanted = ['Folder', 'BOM #', 'Assembly Name:', 'Customer & Platform:', 'Customer Assembly No:']
            columns = [column for column in wanted if column in self.files.columns]
            frame = self.files[['Filename'] + columns].astype(str)
            self._details = {row['Filename']: {column: row[column] for column in columns} for (_, row) in frame.iterrows()}
        return self._details

    def warm_revision_walk(self):
        self._parents_by_file()
        self.bom_dates()
        self.bom_signers()
        self.bom_plants()

    def revision_chain(self, filename, part_number):
        return self._chain_in_file(str(filename), str(part_number).strip(), self._parents_by_file())

    def _chain_in_file(self, filename, part, parents=None):
        if parents is None:
            parents = self._parents_by_file()
        within = parents.get(filename)
        if not within:
            return []
        (chain, seen) = ([], {part})
        current = within.get(part)
        while current and current not in seen and (len(chain) < REVISION_MAX_LEVELS):
            chain.append(current)
            seen.add(current)
            current = within.get(current)
        return chain

    def _sheet_layouts(self):
        if self._layouts is not None:
            return self._layouts
        with self._layout_lock:
            if self._layouts is not None:
                return self._layouts
            cells = self.cells
            values = cells['Value']
            if isinstance(values.dtype, pd.CategoricalDtype):
                banners = {name for name in values.cat.categories if BLOCK_BANNER in str(name).lower()}
                is_banner = values.isin(banners)
            else:
                is_banner = values.astype(str).str.lower().str.contains(BLOCK_BANNER, na=False, regex=False)
            splits = cells[is_banner].groupby(['Filename', 'Sheet'], observed=True)['Col'].min().astype(int).to_dict()
            header = cells[cells['IsHeader'].astype(bool)][['Filename', 'Sheet', 'Row', 'Col', 'Canonical']].copy()
            header['Filename'] = as_text(header['Filename'])
            header['Sheet'] = as_text(header['Sheet'])
            per_row = header.groupby(['Filename', 'Sheet', 'Row'], observed=True).size().reset_index(name='cells').sort_values(['Filename', 'Sheet', 'Row'])
            previous = per_row.groupby(['Filename', 'Sheet'], observed=True)['Row'].shift(1)
            starts_run = per_row['Row'] - previous != 1
            per_row['run'] = starts_run.groupby([per_row['Filename'], per_row['Sheet']]).cumsum()
            widest = per_row.groupby(['Filename', 'Sheet', 'run'], observed=True)['cells'].sum().reset_index().sort_values('cells').groupby(['Filename', 'Sheet'], observed=True).tail(1)
            kept = per_row.merge(widest[['Filename', 'Sheet', 'run']], on=['Filename', 'Sheet', 'run'])
            header = header.merge(kept[['Filename', 'Sheet', 'Row']], on=['Filename', 'Sheet', 'Row'])

            def columns_named(canonical):
                rows = header[header['Canonical'].astype(str) == canonical]
                return rows.groupby(['Filename', 'Sheet'], observed=True)['Col'].apply(lambda column: sorted({int(value) for value in column})).to_dict()
            numbers = columns_named('part_number')
            names = columns_named('part_name')
            levels = columns_named('level')
            layouts = {}
            for (key, number_columns) in numbers.items():
                level_columns = levels.get(key, [])
                layouts[key] = {'numbers': number_columns, 'names': names.get(key, []), 'indent': level_columns[0] if level_columns else None, 'split': splits.get(key)}
            self._layouts = layouts
            return self._layouts

    def _parents_by_file(self):
        if self._parents is not None:
            return self._parents
        with self._layout_lock:
            if self._parents is not None:
                return self._parents
            self._build_hierarchy()
            return self._parents

    def _part_uses(self):
        if self._uses is not None:
            return self._uses
        with self._layout_lock:
            if self._uses is not None:
                return self._uses
            self._build_hierarchy()
            return self._uses

    def _build_hierarchy(self):
        layouts = self._sheet_layouts()
        data = self.cells[~self.cells['IsHeader'].astype(bool)][['Filename', 'Sheet', 'Row', 'Col', 'Value']].copy()
        data['Filename'] = as_text(data['Filename'])
        data['Sheet'] = as_text(data['Sheet'])
        data['Value'] = as_text(data['Value'])
        wanted = [(filename, sheet, column) for ((filename, sheet), layout) in layouts.items() for column in layout['numbers']]
        wanted += [(filename, sheet, column) for ((filename, sheet), layout) in layouts.items() for column in layout['names']]
        indent_columns = {(filename, sheet): layout['indent'] for ((filename, sheet), layout) in layouts.items() if layout['indent'] is not None}
        wanted += [(filename, sheet, column) for ((filename, sheet), column) in indent_columns.items()]
        keys = pd.MultiIndex.from_arrays([data['Filename'], data['Sheet'], data['Col']])
        occurrences = data[keys.isin(pd.MultiIndex.from_tuples(wanted))]
        heads = self.cells[self.cells['IsHeader'].astype(bool)]
        first_heading = heads.groupby([as_text(heads['Filename']), as_text(heads['Sheet']), heads['Col']], observed=True)['Row'].min().to_dict()
        top = pd.Series(list(zip(occurrences['Filename'], occurrences['Sheet'], occurrences['Col']))).map(first_heading).to_numpy()
        occurrences = occurrences[pd.isna(top) | (occurrences['Row'].to_numpy() > top)]
        ordered = occurrences.sort_values(['Filename', 'Sheet'], kind='stable')
        keys = list(zip(ordered['Filename'], ordered['Sheet']))
        all_rows = ordered['Row'].tolist()
        all_columns = ordered['Col'].tolist()
        all_values = ordered['Value'].tolist()
        (bounds, start) = ([], 0)
        for position in range(1, len(keys) + 1):
            if position == len(keys) or keys[position] != keys[start]:
                bounds.append((keys[start], start, position))
                start = position
        (parents, uses, names, seen) = ({}, {}, {}, {})
        for ((filename, sheet), start, end) in bounds:
            layout = layouts.get((filename, sheet))
            if layout is None:
                continue
            by_row = {}
            for (row, column, value) in zip(all_rows[start:end], all_columns[start:end], all_values[start:end]):
                by_row.setdefault(int(row), {})[int(column)] = value
            if layout['indent'] is not None:
                (links, found) = self._tree_by_indent(by_row, layout)
            elif layout['split'] is not None:
                (links, found) = self._tree_by_halves(by_row, layout)
            else:
                (links, found) = ({}, [(part_number_key(value), '') for cells in by_row.values() for value in cells.values()])
            within = parents.setdefault(filename, {})
            for (child, parent) in links.items():
                within.setdefault(child, parent)
            for (value, name) in found:
                if value:
                    uses.setdefault(value, set()).add(filename)
                    seen.setdefault(filename, set()).add(value)
                    if name and value not in names:
                        names[value] = name
        assemblies = self._assembly_numbers()
        self._assemblies = assemblies
        for (filename, assembly) in assemblies.items():
            uses.setdefault(assembly, set()).add(filename)
        for (filename, values) in seen.items():
            assembly = assemblies.get(filename, '')
            if not assembly:
                continue
            within = parents.setdefault(filename, {})
            for value in values:
                if value != assembly:
                    within.setdefault(value, assembly)
        self._parents = parents
        self._uses = uses
        self._part_names = names

    def assembly_numbers(self):
        if self._assemblies is None:
            self._assemblies = self._assembly_numbers()
        return self._assemblies

    def owns_a_bom(self, part_number):
        number = str(part_number).strip()
        return number in set(self.assembly_numbers().values())

    def siblings_sharing_drawing(self, part_number):
        number = str(part_number).strip()
        if len(number) <= DRAWING_NUMBER_LENGTH:
            return []
        prefix = number[:DRAWING_NUMBER_LENGTH]
        if not prefix.isdigit():
            return []
        return sorted((known for known in self._part_uses() if known != number and known[:DRAWING_NUMBER_LENGTH] == prefix))

    def _assembly_numbers(self):
        column = 'Internal Assembly No:'
        if column not in self.files.columns:
            return {}
        numbers = {}
        bom_numbers = self.files['BOM #'].astype(str) if 'BOM #' in self.files.columns else pd.Series('', index=self.files.index)
        for (filename, internal, bom) in zip(self.files['Filename'].astype(str), self.files[column].astype(str), bom_numbers):
            found = part_number_key(internal)
            if not found:
                stem = str(bom).strip().rsplit('-', 1)[0]
                found = part_number_key(stem)
            if found:
                numbers[filename] = found
        return numbers

    @staticmethod
    def _tree_by_indent(by_row, layout):
        indent_column = layout['indent']
        number_column = layout['numbers'][0]
        name_column = layout['names'][0] if layout['names'] else None
        (ordered, links, found) = (sorted(by_row), {}, [])
        (depth_at, part_at) = ({}, {})
        for row in ordered:
            cells = by_row[row]
            depth = cells.get(indent_column, '').strip()
            part = part_number_key(cells.get(number_column, ''))
            if part:
                part_at[row] = part
                found.append((part, cells.get(name_column, '').strip() if name_column is not None else ''))
            if depth.isdigit():
                depth_at[row] = int(depth)
        for row in ordered:
            (depth, part) = (depth_at.get(row), part_at.get(row))
            if part is None or not depth:
                continue
            above = [candidate for candidate in ordered if candidate < row and depth_at.get(candidate) == depth - 1]
            if above:
                parent = part_at.get(above[-1])
                if parent and parent != part:
                    links.setdefault(part, parent)
        return (links, found)

    @staticmethod
    def _tree_by_halves(by_row, layout):
        split = layout['split']
        owner_column = min((column for column in layout['numbers'] if column < split), default=None)
        child_column = min((column for column in layout['numbers'] if column >= split), default=None)
        if owner_column is None or child_column is None:
            return ({}, [(part_number_key(value), '') for cells in by_row.values() for value in cells.values()])
        split = layout['split']
        owner_name_column = min((column for column in layout['names'] if column < split), default=None)
        child_name_column = min((column for column in layout['names'] if column >= split), default=None)

        def name_at(row, column):
            return by_row[row].get(column, '').strip() if column is not None else ''
        ordered = sorted(by_row)
        owner_at = {row: part_number_key(by_row[row].get(owner_column, '')) for row in ordered}
        owner_rows = [row for row in ordered if owner_at.get(row)]
        links = {}
        found = [(owner_at[row], name_at(row, owner_name_column)) for row in owner_rows]
        owner_index = 0
        current_owner = None
        for row in ordered:
            while owner_index < len(owner_rows) and owner_rows[owner_index] <= row:
                current_owner = owner_at.get(owner_rows[owner_index])
                owner_index += 1
            child = part_number_key(by_row[row].get(child_column, ''))
            if not child:
                continue
            found.append((child, name_at(row, child_name_column)))
            if current_owner and child != current_owner:
                links.setdefault(child, current_owner)
        return (links, found)

    def parts(self, filename):
        rows = self.cells[(self.cells['Filename'] == filename) & ~self.cells['IsHeader'] & (self.cells['Canonical'] != '')]
        if rows.empty:
            return pd.DataFrame()
        rows = rows.assign(Value=rows['Value'].astype(str))
        table = rows.pivot_table(index=['Sheet', 'Row'], columns='Canonical', values='Value', aggfunc='first', observed=True)
        table = table.reset_index()
        if 'part_number' not in table.columns:
            return pd.DataFrame()
        table = table[table['part_number'].notna()]
        table = table.assign(part_number=table['part_number'].map(part_spelling))
        table = table[table['part_number'].map(looks_like_part_number)]
        for (canonical, _) in PART_COLUMNS:
            if canonical not in table.columns:
                table[canonical] = ''
        table = table.astype(object)
        table = table.where(table.notna(), '')
        table['drawing'] = [self.drawing_spelling(value, part) for (value, part) in zip(table['drawing'], table['part_number'])]
        if 'drawing' in table.columns:
            missing = table['drawing'] == ''
            table.loc[missing, 'drawing'] = [self.drawing_for(value) for value in table.loc[missing, 'part_number']]
        table = table[[canonical for (canonical, _) in PART_COLUMNS]]
        merged = []
        seen = {}
        for (_, row) in table.iterrows():
            number = str(row['part_number'])
            if number in seen:
                target = merged[seen[number]]
                for column in table.columns:
                    if not str(target[column]).strip():
                        target[column] = row[column]
            else:
                seen[number] = len(merged)
                merged.append(row.copy())
        return pd.DataFrame(merged, columns=table.columns).reset_index(drop=True)

    def alternative_terms(self, term, scope=None, limit=4):
        text = str(term).strip()
        if len(text) < 4:
            return []
        (tried, found) = ({text.lower()}, [])

        def consider(candidate):
            candidate = candidate.strip()
            if len(candidate) < 3 or candidate.lower() in tried:
                return
            tried.add(candidate.lower())
            hits = self.search(candidate, scope=scope)
            if hits is not None and (not hits.empty):
                found.append((candidate, len(hits)))
        if text.lower().endswith('s'):
            consider(text[:-1])
        words = [word for word in re.split('[\\s,/_-]+', text) if len(word) >= 3]
        if len(words) > 1:
            for word in words:
                consider(word)
                if word.lower().endswith('s'):
                    consider(word[:-1])
        found.sort(key=lambda pair: (-len(pair[0]), -pair[1]))
        return found[:limit]

    def _known_numbers(self):
        if self._numbers is not None:
            return self._numbers
        known = set()
        values = self.cells['Value'].cat.categories if isinstance(self.cells['Value'].dtype, pd.CategoricalDtype) else self.cells['Value']
        for value in values.astype(str):
            text = value.strip()
            if 5 <= len(text) <= 10 and text.isdigit():
                known.add(text)
        for name in self.pdfs['PDF_Name'].astype(str):
            match = re.match('\\d+', name)
            if match:
                known.add(match.group(0))
        self._numbers = known
        return known

    def similar_numbers(self, term, limit=5):
        wanted = str(term).strip()
        if not 4 <= len(wanted) <= 10 or not wanted.isdigit():
            return []
        known = self._known_numbers()
        candidates = set()
        digits = '0123456789'
        for position in range(len(wanted)):
            for digit in digits:
                candidates.add(wanted[:position] + digit + wanted[position + 1:])
            candidates.add(wanted[:position] + wanted[position + 1:])
        for position in range(len(wanted) + 1):
            for digit in digits:
                candidates.add(wanted[:position] + digit + wanted[position:])
        candidates.discard(wanted)
        near = sorted((candidate for candidate in candidates if candidate in known))
        near.sort(key=lambda number: (len(number) != len(wanted), number))
        return near[:limit]

    def drawings_matching(self, term, limit=8):
        wanted = str(term).strip().upper()
        if len(wanted) < 4 or not any((character.isdigit() for character in wanted)):
            return pd.DataFrame()
        names = self.pdfs['PDF_Name'].astype(str)
        stems = names.str.upper().str.replace('\\.PDF$', '', regex=True)
        found = stems.str.startswith(wanted) | stems.str.contains('[_\\-]' + re.escape(wanted), regex=True)
        if not found.any():
            return pd.DataFrame()
        return self.pdfs[found].head(limit).reset_index(drop=True)

    def drawing_spelling(self, drawing, part_number):
        text = strip_mark(drawing)
        padded = restore_leading_zero(text, DRAWING_NUMBER_LENGTH)
        if padded != text and (str(part_number)[:DRAWING_NUMBER_LENGTH] == padded or padded in self._drawing_by_prefix):
            return padded
        return text

    def drawing_for(self, part_number):
        if prototype_key(part_number) is not None:
            found = self.prototype_files('drawing', part_number)
            if len(found) > 1:
                return f'{len(found)} files'
            return os.path.splitext(os.path.basename(found[0]))[0] if found else ''
        digits = re.match('\\d+', str(part_number).strip())
        if not digits:
            return ''
        name = self._drawing_by_prefix.get(digits.group(0)[:DRAWING_NUMBER_LENGTH])
        return drawing_label(os.path.splitext(name)[0]) if name else ''

    def prototype_files(self, kind, number):
        wanted = prototype_key(number)
        if wanted is None:
            return []
        files = self._prototype_file_map(kind)
        return sorted({path for (design, path) in files.get(wanted[:2], ()) if serves(wanted, wanted[:2] + (design,))})

    def _prototype_file_map(self, kind):
        maps = getattr(self, '_prototype_maps', None)
        if maps is None:
            maps = self._prototype_maps = {}
        if kind not in maps:
            (table, column) = (self.pdfs, 'PDF_Name') if kind == 'drawing' else (self.cad, 'Name')
            built = {}
            for (folder, name) in zip(table['Folder'].astype(str), table[column].astype(str)):
                for (project, part, design) in prototype_keys(name):
                    built.setdefault((project, part), []).append((design, os.path.join(folder, name)))
            maps[kind] = built
        return maps[kind]

    def columns_present(self, filename):
        rows = self.cells[(self.cells['Filename'] == filename) & (self.cells['Canonical'] != '')]
        return set(rows['Canonical'].unique())
PRODUCTION = 'Production'
PROTOTYPE = 'Prototype'
PROTOTYPE_NUMBER = re.compile('(?<!\\d)(\\d{6})[-_ ]+(\\d{4})(?!\\d)(?:[-_ ]+(\\d{1,3})(?!\\d))?')
CAD_TABLE = 'bom_cad.feather'
_DRAWING_LABEL = re.compile('^(\\d{7}(?:[_ ,&]+\\d{7})*-[A-Z]{1,2})(?![A-Za-z])')

def drawing_label(stem):
    found = _DRAWING_LABEL.match(stem)
    return found.group(1) if found else stem

def _prototype_groups(project, part, design):
    return (project, part, int(design) if design else None)

def prototype_key(number):
    match = PROTOTYPE_NUMBER.search(str(number))
    return _prototype_groups(*match.groups()) if match else None
_SHARED_PART = re.compile('[-_ ]+(\\d{4})(?!\\d)')

def prototype_keys(name):
    text = str(name)
    keys = []
    for match in PROTOTYPE_NUMBER.finditer(text):
        keys.append(_prototype_groups(*match.groups()))
        at = match.end()
        while True:
            more = _SHARED_PART.match(text, at)
            if not more:
                break
            keys.append(_prototype_groups(match.group(1), more.group(1), None))
            at = more.end()
    return keys

def serves(wanted, found):
    return wanted[:2] == found[:2] and (wanted[2] is None or found[2] is None or wanted[2] == found[2])

def load_cad(path):
    if not os.path.isfile(path):
        return pd.DataFrame(columns=['Folder', 'Name'])
    return _read(path, 'CAD list')[0]
INDEX_TABLES = ('bom_cells.feather', 'bom_files.feather', 'bom_pdfs.feather')

def load_prototype(folder):
    if not folder:
        return None
    paths = [os.path.join(folder, name) for name in INDEX_TABLES]
    if not all((os.path.isfile(path) for path in paths)):
        return None
    index = load(*paths)
    index.cad_error = ''
    try:
        index.cad = load_cad(os.path.join(folder, CAD_TABLE))
    except IndexError_ as error:
        index.cad = pd.DataFrame(columns=['Folder', 'Name'])
        index.cad_error = f'CAD list unreadable: {error}'
    return index

def _stack(production, prototype):
    frames = [frame.assign(Source=source) for (frame, source) in ((production, PRODUCTION), (prototype, PROTOTYPE)) if frame is not None and (not frame.empty)]
    if not frames:
        return production
    return pd.concat(frames, ignore_index=True)

def search_both(index, term, limit=None, scope=None):
    production = index.search(term, limit=limit, scope=scope)
    return _stack(production, _prototype_rows(index, 'search', term, limit, scope))

def search_components_both(index, term, limit=None, scope=None):
    production = index.search_components(term, limit=limit, scope=scope)
    return _stack(production, _prototype_rows(index, 'search_components', term, limit, scope))

def _prototype_rows(index, method, term, limit, scope):
    other = getattr(index, 'prototype', None)
    if other is None:
        return None
    try:
        rows = getattr(other, method)(term, limit=limit, scope=scope)
    except Exception as error:
        index.prototype_error = f'{type(error).__name__}: {error}'
        return None
    index.prototype_error = ''
    return rows

def _read(path, what):
    if not os.path.exists(path):
        raise IndexError_(f'The {what} file is missing:\n\n{path}')
    try:
        table = feather.read_table(path)
    except Exception as error:
        raise IndexError_(f'The {what} file could not be read:\n\n{path}\n\n{error}')
    metadata = {key.decode(): value.decode() for (key, value) in (table.schema.metadata or {}).items() if key.decode().startswith('demo_')}
    return (table.to_pandas(), metadata)

def load(cells_path, files_path, pdfs_path, projects_path=None, project_folders_path=None, obsolete_path=None, dfmea_path=None, prototype_dir=None):
    (cells, info) = _read(cells_path, 'BOM contents')
    (files, _) = _read(files_path, 'BOM list')
    (pdfs, _) = _read(pdfs_path, 'drawing list')
    version = info.get('demo_format_version')
    if version and version != SUPPORTED_FORMAT:
        raise IndexError_(f'This index is in format {version}, and this version of BOM Search Demo reads format {SUPPORTED_FORMAT}.\n\nEither the index or the app needs updating.')
    for (frame, name, needed) in ((cells, 'BOM contents', ['Folder', 'Filename', 'Sheet', 'Row', 'Col', 'ColHeader', 'IsHeader', 'Value']), (files, 'BOM list', ['Folder', 'Filename', 'BOM #']), (pdfs, 'drawing list', ['Folder', 'PDF_Name'])):
        missing = [column for column in needed if column not in frame.columns]
        if missing:
            raise IndexError_(f"The {name} file is not in the expected shape.\n\nMissing: {', '.join(missing)}")
    index = Index(cells, files, pdfs, info)
    index.projects = load_projects(projects_path)
    index.project_folders = load_project_folders(project_folders_path)
    index.obsolete = load_obsolete(obsolete_path)
    index.dfmea = load_dfmea(dfmea_path)
    try:
        index.prototype = load_prototype(prototype_dir)
    except Exception as error:
        index.prototype = None
        index.prototype_error = f'{type(error).__name__}: {error}'
    return index

def load_obsolete(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        (table, _) = _read(path, 'obsolete BOM list')
    except IndexError_:
        return pd.DataFrame()
    return table
