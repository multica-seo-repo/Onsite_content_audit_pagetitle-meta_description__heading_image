"""Versioned image evidence transport. Shared verbatim by producer and consumer."""
import hashlib
import io
import json
from pathlib import Path
import zipfile

SCHEMA = 'image-evidence-v1'
MARKER = '[COMPACT_IMAGE_EVIDENCE_V1]'
DEFAULTS = {'row_limit': 50000, 'preview_rows': 1000, 'target_bytes': 8 * 1024 * 1024}


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def settings(values=None):
    result = dict(DEFAULTS)
    if values:
        if set(values) - set(result):
            raise ValueError('Unknown image delivery setting')
        result.update(values)
    if any(type(v) is not int or v <= 0 for v in result.values()):
        raise ValueError('Image delivery settings require positive integers')
    if result['preview_rows'] > 1000 or result['target_bytes'] < 2048:
        raise ValueError('Image preview limit is 1000; packaging target must be at least 2048 bytes')
    return result


def safe_file(folder, name):
    if not isinstance(name, str) or not name or Path(name).name != name or '/' in name or '\\' in name or name in {'.', '..'}:
        raise ValueError('Evidence filename must be a relative basename')
    p = (Path(folder) / name).resolve()
    if p.parent != Path(folder).resolve():
        raise ValueError('Evidence path escapes selected directory')
    return p


def literal(value):
    return '' if value is None else str(value)


def package_check(check, folder, stem, options, created):
    """Bound chunk buffers by uncompressed bytes; preserve row order and duplicates."""
    parts, lines, size, start = [], [], 0, 1
    target = options['target_bytes']
    headers = check['headers']
    pages, assets = set(), set()
    def flush():
        nonlocal lines, size, start
        if not lines:
            return
        name = f'{stem}_{check["id"]}_evidence_v1_{len(parts)+1:04}.zip'
        member = name[:-4] + '.jsonl'
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr(member, b''.join(lines))
        raw = buffer.getvalue()
        if len(raw) > target:
            raise ValueError('Evidence package exceeds configured target')
        path = safe_file(folder, name)
        with path.open('xb') as f:
            created.append(path)
            f.write(raw)
        parts.append({'file': name, 'member': member, 'sha256': digest(path), 'size_bytes': len(raw),
                      'uncompressed_bytes': size, 'first_row': start, 'last_row': start + len(lines)-1, 'row_count': len(lines)})
        start += len(lines)
        lines, size = [], 0
    for index, row in enumerate(check['rows'], 1):
        values = dict(zip(headers, row))
        if len(row) > len(headers):
            raise ValueError('Invalid evidence row width')
        pages.add(literal(values.get('Source'))); assets.add(literal(values.get('Destination')))
        line = (json.dumps({'row': index, 'values': row}, ensure_ascii=False, allow_nan=False) + '\n').encode('utf-8')
        if len(line) > target - 1024:
            raise ValueError('One evidence row exceeds packaging target; increase target explicitly')
        if size + len(line) > target - 1024:
            flush()
        lines.append(line); size += len(line)
    flush()
    return {'headers': headers, 'reference_rows': len(check['rows']), 'unique_assets': len(assets - {''}),
            'affected_pages': len(pages - {''}), 'parts': parts}


def verify_manifest(path, workbook=None, row_validator=None):
    path = Path(path)
    m = json.loads(path.read_text(encoding='utf-8'))
    if m.get('schema') != SCHEMA or not m.get('run_id') or not m.get('site_name') or not m.get('site_url'):
        raise ValueError('Invalid image evidence identity/schema')
    opts = settings(m['settings'])
    if not m.get('checks') or set(m['checks']) - {'20.1', '20.2'}:
        raise ValueError('Invalid image evidence check coverage')
    if workbook and (Path(workbook).name != m['workbook']['file'] or digest(workbook) != m['workbook']['sha256']):
        raise ValueError('Image evidence workbook hash/identity mismatch')
    seen = set()
    for ident, c in m['checks'].items():
        headers = c['headers']
        if not isinstance(headers,list) or len(headers)!=len(set(headers)) or not {'Source','Destination'}.issubset(headers):
            raise ValueError('Invalid image evidence headers')
        count, preview, pages, assets = 0, [], set(), set()
        for part in c['parts']:
            if part['file'] in seen:
                raise ValueError('Duplicate evidence package')
            seen.add(part['file'])
            p = safe_file(path.parent, part['file'])
            if p.stat().st_size != part['size_bytes'] or p.stat().st_size > opts['target_bytes'] or digest(p) != part['sha256']:
                raise ValueError('Image evidence package checksum/size mismatch')
            if part['first_row'] != count+1:
                raise ValueError('Image evidence row range gap/overlap')
            n = 0
            with zipfile.ZipFile(p) as z:
                if z.namelist() != [part['member']] or Path(part['member']).name != part['member']:
                    raise ValueError('Unexpected evidence ZIP member')
                info = z.getinfo(part['member'])
                if info.file_size != part['uncompressed_bytes'] or info.file_size > opts['target_bytes']:
                    raise ValueError('Evidence expanded size exceeds bounded contract')
                with z.open(info) as f:
                    for line in f:
                        record = json.loads(line)
                        count += 1; n += 1
                        row = record['values']
                        if record['row'] != count or not isinstance(row,list) or len(row)>len(headers) or any(isinstance(v,(dict,list)) for v in row):
                            raise ValueError('Invalid evidence row order/schema')
                        values = dict(zip(headers,row))
                        if row_validator:
                            row_validator(ident, values)
                        pages.add(literal(values.get('Source'))); assets.add(literal(values.get('Destination')))
                        if len(preview)<opts['preview_rows']:
                            preview.append(row)
            if n != part['row_count'] or count != part['last_row']:
                raise ValueError('Image evidence range/count mismatch')
        if not count or count != c['reference_rows'] or len(assets-{''})!=c['unique_assets'] or len(pages-{''})!=c['affected_pages']:
            raise ValueError('Image evidence totals mismatch')
        if preview != c['preview'] or c['displayed_rows'] != len(preview):
            raise ValueError('Image evidence preview mismatch')
    return m
