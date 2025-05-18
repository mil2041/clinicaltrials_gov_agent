import zipfile
import xml.etree.ElementTree as ET
import csv

INPUT_FILE = 'data/ctg-studies_CAR.parsed.xlsx'
OUTPUT_FILE = 'data/ctg-studies_CAR.classified.csv'

# keywords for blood cancers and solid tumors
blood_keywords = [
    'leukemia', 'lymphoma', 'myeloma', 'hematologic', 'hematological',
    'lymphoblastic', 'lymphoid', 'myeloid', 'plasma cell',
]
solid_keywords = [
    'cancer', 'carcinoma', 'sarcoma', 'tumor', 'glioblastoma', 'glioma'
]

def load_sheet_rows(path):
    z = zipfile.ZipFile(path)
    shared_xml = z.read('xl/sharedStrings.xml')
    root = ET.fromstring(shared_xml)
    shared = [el.text if el.text else ''
              for el in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]
    sheet = z.read('xl/worksheets/sheet1.xml')
    root = ET.fromstring(sheet)
    rows = []
    for row in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
        cells = []
        for c in row.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            t = c.get('t')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if t == 's' and val:
                val = shared[int(val)]
            cells.append(val)
        rows.append(cells)
    return rows

def classify_condition(text):
    lower = text.lower()
    for kw in blood_keywords:
        if kw in lower:
            return 'blood cancer'
    for kw in solid_keywords:
        if kw in lower:
            return 'solid tumor'
    return 'others'

def main():
    rows = load_sheet_rows(INPUT_FILE)
    header = rows[0]
    nct_idx = header.index('nctId')
    cond_idx = header.index('conditions')

    with open(OUTPUT_FILE, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['nctId', 'conditions', 'classification'])
        for r in rows[1:]:
            if len(r) <= max(nct_idx, cond_idx):
                continue
            nct = r[nct_idx]
            cond = r[cond_idx]
            cls = classify_condition(cond)
            writer.writerow([nct, cond, cls])

if __name__ == '__main__':
    main()
