"""Regenerate the small synthetic dataset used by the portfolio demo."""

from datetime import datetime
from pathlib import Path

import pandas as pd
import pyarrow as pa
import pyarrow.feather as feather
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'demo_data'
INDEX = DATA / 'index'
FILES = DATA / 'files'

BOMS = [
    ('01-Assemblies', '7001234001-A', '7001234001', 'Gearbox Assembly', 'Demo Customer One', '61001', [
        ('7001234101', 'Drive Gear', '7001234', 'Steel', '0.125', '0.140', 'Black oxide', '1'),
        ('7001234201', 'Support Bracket', '7001234', 'Aluminum', '0.080', '0.095', 'Clear coat', '2'),
    ]),
    ('01-Assemblies', '7100001001-B', '7100001001', 'Lift Module', 'Demo Customer One', '61002', [
        ('7001234101', 'Drive Gear', '7001234', 'Steel', '0.125', '0.140', 'Black oxide', '4'),
        ('7100001101', 'Pivot Pin', '7100001', 'Alloy steel', '0.240', '0.250', 'Zinc plate', '2'),
    ]),
    ('02-Motion Systems', '7200045001-C', '7200045001', 'Linkage Kit', 'Demo Customer Two', '62011', [
        ('7200045101', 'Connecting Link', '7200045', 'Steel', '0.095', '0.100', 'E-coat', '1'),
        ('7100001101', 'Pivot Pin', '7100001', 'Alloy steel', '0.240', '0.250', 'Zinc plate', '2'),
    ]),
    ('02-Motion Systems', '7300066001-A', '7300066001', 'Actuator Assembly', 'Demo Customer Two', '62012', [
        ('7300066101', 'Actuator Rod', '7300066', 'Stainless steel', '0.300', '0.310', 'Passivated', '1'),
        ('7001234201', 'Support Bracket', '7001234', 'Aluminum', '0.080', '0.095', 'Clear coat', '2'),
    ]),
    ('03-Mounting Hardware', '7400088001-D', '7400088001', 'Mounting Set', 'Demo Customer Three', '63021', [
        ('7400088101', 'Mounting Plate', '7400088', 'Steel', '0.180', '0.190', 'Powder coat', '1'),
        ('7100001101', 'Pivot Pin', '7100001', 'Alloy steel', '0.240', '0.250', 'Zinc plate', '2'),
    ]),
    ('03-Mounting Hardware', '7500099001-B', '7500099001', 'Guard Assembly', 'Demo Customer Three', '63022', [
        ('7500099101', 'Safety Guard', '7500099', 'Stainless steel', '0.060', '0.070', 'Brushed', '1'),
        ('7400088101', 'Mounting Plate', '7400088', 'Steel', '0.180', '0.190', 'Powder coat', '4'),
    ]),
]


def write_pdf(path, title):
    path.parent.mkdir(parents=True, exist_ok=True)
    text = title.replace('(', '[').replace(')', ']')
    stream = f'BT /F1 18 Tf 72 720 Td ({text}) Tj ET'.encode('ascii', 'replace')
    objects = [
        b'<< /Type /Catalog /Pages 2 0 R >>',
        b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        b'<< /Length ' + str(len(stream)).encode() + b' >>\nstream\n' + stream + b'\nendstream',
    ]
    out = bytearray(b'%PDF-1.4\n')
    offsets = [0]
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out.extend(f'{number} 0 obj\n'.encode())
        out.extend(body + b'\nendobj\n')
    xref = len(out)
    out.extend(f'xref\n0 {len(objects) + 1}\n'.encode())
    out.extend(b'0000000000 65535 f \n')
    for offset in offsets[1:]:
        out.extend(f'{offset:010d} 00000 n \n'.encode())
    out.extend(f'trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode())
    path.write_bytes(out)


def categorized(frame, names):
    for name in names:
        if name in frame.columns:
            frame[name] = frame[name].astype('category')
    return frame


def main():
    INDEX.mkdir(parents=True, exist_ok=True)
    cells = []
    file_rows = []
    pdf_rows = []
    drawings = set()
    headers = ['Part Number', 'Part Name', 'Drawing Number', 'Qty',
               'Material Specification', 'Thk-min', 'Thk-max',
               'Surface Finish', 'Level']
    for folder, bom, assembly, name, customer, ecn, components in BOMS:
        filename = bom + '.xlsx'
        file_rows.append({
            'Folder': folder, 'Filename': filename, 'BOM #': bom,
            'Internal Assembly No:': assembly, 'Assembly Name:': name,
            'Customer & Platform:': customer + ' / Demo Platform',
            'Customer Assembly No:': 'C-' + assembly[-5:],
            'Customer Service No:': 'S-' + assembly[-5:],
            'ECN#:': ecn, 'ML': 'PLANT-' + str(1 + len(file_rows) % 3),
            'Date': '2026-09-15', 'Signed By': 'Demo Reviewer',
        })
        all_rows = [(assembly, name, assembly[:7], '1', '—', '', '', '—', '0')] + components
        for col, header in enumerate(headers):
            cells.append({'Folder': folder, 'Filename': filename, 'Sheet': 'BOM',
                          'Row': 1, 'Col': col, 'ColHeader': header,
                          'IsHeader': True, 'Value': header})
        for row_number, values in enumerate(all_rows, 2):
            for col, (header, value) in enumerate(zip(headers, values)):
                cells.append({'Folder': folder, 'Filename': filename, 'Sheet': 'BOM',
                              'Row': row_number, 'Col': col, 'ColHeader': header,
                              'IsHeader': False, 'Value': str(value)})
        for _, _, drawing, *_ in components:
            drawings.add((folder, drawing))
        xlsx_path = FILES / 'boms' / folder / filename
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        book = Workbook()
        sheet = book.active
        sheet.title = 'BOM'
        sheet.append(headers)
        sheet.append([assembly, name, assembly[:7], '1', '—', '', '', '—', '0'])
        for component in components:
            sheet.append(list(component))
        book.save(xlsx_path)

    for folder, drawing in sorted(drawings):
        pdf_name = drawing + '-A.pdf'
        pdf_rows.append({'Folder': folder, 'PDF_Name': pdf_name})
        write_pdf(FILES / 'drawings' / folder / pdf_name,
                  f'Synthetic drawing {drawing}')
        cad = FILES / 'cad' / folder / (drawing + '-A.step')
        cad.parent.mkdir(parents=True, exist_ok=True)
        cad.write_text('ISO-10303-21;\n/* synthetic demo geometry */\n', encoding='ascii')

    metadata = {b'demo_format_version': b'2',
                b'demo_built_at': datetime.now().isoformat().encode()}
    cells_frame = categorized(pd.DataFrame(cells),
                              ['Folder', 'Filename', 'Sheet', 'ColHeader', 'Value'])
    files_frame = categorized(pd.DataFrame(file_rows), ['Folder', 'Filename'])
    pdfs_frame = categorized(pd.DataFrame(pdf_rows), ['Folder', 'PDF_Name'])
    for frame, filename, meta in ((cells_frame, 'bom_cells.feather', metadata),
                                  (files_frame, 'bom_files.feather', None),
                                  (pdfs_frame, 'bom_pdfs.feather', None)):
        table = pa.Table.from_pandas(frame, preserve_index=False)
        if meta:
            table = table.replace_schema_metadata(meta)
        feather.write_feather(table, INDEX / filename)

    projects = pd.DataFrame([
        {'ECN': '61001', 'Project': '700101', 'Source': 'Change 61001.docx'},
        {'ECN': '62011', 'Project': '700201', 'Source': 'Change 62011.docx'},
    ])
    feather.write_feather(pa.Table.from_pandas(projects, preserve_index=False),
                          INDEX / 'change_projects.feather')
    project_folders = pd.DataFrame([
        {'Project': '700101', 'Path': 'demo_data/files/projects/active/700101 Demo Project'},
        {'Project': '700201', 'Path': 'demo_data/files/projects/closed/700201 Demo Project'},
    ])
    feather.write_feather(pa.Table.from_pandas(project_folders, preserve_index=False),
                          INDEX / 'project_folders.feather')
    for project in ('700101 Demo Project', '700201 Demo Project'):
        (FILES / 'projects' / ('active' if project.startswith('700101') else 'closed') / project).mkdir(parents=True, exist_ok=True)
    print(f'Wrote {len(file_rows)} synthetic BOMs, {len(cells_frame)} indexed cells, and {len(pdfs_frame)} synthetic drawings.')


if __name__ == '__main__':
    main()
